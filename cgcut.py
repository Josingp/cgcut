#!/usr/bin/env python3
"""cgcut — CG 리스트(xlsx)의 TC를 읽어 원본 영상에서 해당 샷만 잘라내는 도구.

사용 예:
    python cgcut.py CG리스트.xlsx --video 편집본.mov -o out/
    python cgcut.py CG리스트.xlsx --video-dir ./편집본폴더/ -o out/ --codec prores --handles 12
    python cgcut.py CG리스트.xlsx --dry-run
    python cgcut.py CG리스트.xlsx --clips-dir ./카메라원본/ -o out/   # Scan Name(클립명)으로 원본 클립 모으기

자세한 설명은 README.md 참고.
"""
from __future__ import annotations

import argparse
import csv
import json
import os
import re
import shutil
import subprocess
import sys
import time
from dataclasses import dataclass, field
from fractions import Fraction
from pathlib import Path
from typing import Iterable, Optional

__version__ = "0.1.0"

VIDEO_EXTS = {".mov", ".mp4", ".mxf", ".mkv", ".avi", ".m4v", ".mts", ".m2ts", ".mpg", ".mpeg", ".webm", ".wmv"}

# ---------------------------------------------------------------------------
# 콘솔 출력 (Windows cp949 콘솔에서도 한글이 깨지지 않게)
# ---------------------------------------------------------------------------
for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8", errors="replace")  # type: ignore[attr-defined]
    except Exception:
        pass


LOG_HOOK = None  # GUI 등에서 교체: callable(str)


def log(msg: str = "") -> None:
    if LOG_HOOK:
        LOG_HOOK(msg)
    else:
        print(msg, flush=True)


def warn(msg: str) -> None:
    if LOG_HOOK:
        LOG_HOOK(f"[경고] {msg}")
    else:
        print(f"[경고] {msg}", file=sys.stderr, flush=True)


class CutError(Exception):
    """치명적 오류 (CLI에서는 메시지 출력 후 종료, GUI에서는 창으로 표시)."""


# ---------------------------------------------------------------------------
# 타임코드
# ---------------------------------------------------------------------------
# 01:04:24:08 / 01:04:24;08(드롭프레임) / 01:04:24.08 / 01:04:24(프레임 생략)
TC_RE = re.compile(r"(?<![\d.])(\d{1,2})[:;.](\d{2})[:;.](\d{2})(?:([:;.,])(\d{2,3}))?(?!\d)")


@dataclass(frozen=True)
class TC:
    h: int
    m: int
    s: int
    f: int
    drop: bool = False

    @staticmethod
    def parse(text: str) -> Optional["TC"]:
        m = TC_RE.search(text or "")
        return TC._from_match(m) if m else None

    @staticmethod
    def find_all(text: str) -> list["TC"]:
        return [TC._from_match(m) for m in TC_RE.finditer(text or "")]

    @staticmethod
    def _from_match(m: "re.Match[str]") -> "TC":
        h, mm, s = int(m.group(1)), int(m.group(2)), int(m.group(3))
        sep, f = m.group(4), m.group(5)
        return TC(h, mm, s, int(f) if f is not None else 0, drop=(sep == ";"))

    @staticmethod
    def _drop_per_min(tc_fps: int) -> int:
        return {30: 2, 60: 4}.get(tc_fps, 0)

    def to_frames(self, tc_fps: int) -> int:
        """타임코드 → 프레임 번호 (00:00:00:00 기준)."""
        n = ((self.h * 60 + self.m) * 60 + self.s) * tc_fps + self.f
        if self.drop:
            d = self._drop_per_min(tc_fps)
            total_min = self.h * 60 + self.m
            n -= d * (total_min - total_min // 10)
        return n

    @staticmethod
    def from_frames(n: int, tc_fps: int, drop: bool = False) -> "TC":
        n = max(0, int(n))
        if drop and TC._drop_per_min(tc_fps):
            d = TC._drop_per_min(tc_fps)
            per_10min = tc_fps * 600 - 9 * d
            per_min = tc_fps * 60 - d
            tens, rem = divmod(n, per_10min)
            n += d * 9 * tens
            if rem >= d:
                n += d * ((rem - d) // per_min)
        f = n % tc_fps
        s = (n // tc_fps) % 60
        m = (n // (tc_fps * 60)) % 60
        h = n // (tc_fps * 3600)
        return TC(h, m, s, f, drop=drop)

    def __str__(self) -> str:
        sep = ";" if self.drop else ":"
        return f"{self.h:02d}:{self.m:02d}:{self.s:02d}{sep}{self.f:02d}"

    def filename(self) -> str:
        return f"{self.h:02d}-{self.m:02d}-{self.s:02d}-{self.f:02d}"


def tc_base(fps: Fraction) -> int:
    """실제 fps(23.976 등) → 타임코드 계산용 정수 fps(24 등)."""
    return max(1, int(round(float(fps))))


# ---------------------------------------------------------------------------
# xlsx 읽기
# ---------------------------------------------------------------------------
HEADER_ROLES: dict[str, set[str]] = {
    "no": {"no", "no.", "num", "#", "번호", "순번", "n", "idx"},
    "ep": {"ep", "ep.", "episode", "에피소드", "화", "ep#", "ep no", "ep no."},
    "ep_alt": {"회차", "회"},
    "tc": {"tc", "timecode", "time code", "tc in", "tc_in", "tc(in)", "in", "in tc", "타임코드", "tc(시작)", "시작 tc", "start tc", "src tc", "rec tc"},
    "tc_out": {"tc out", "tc_out", "tc(out)", "out", "out tc", "종료 tc", "끝 tc", "end tc"},
    "clip": {"scan name", "scanname", "scan", "clip", "clip name", "clipname", "클립", "클립명", "reel", "source clip", "원본 클립", "원본클립", "파일명", "file name", "filename"},
    "desc": {"work order", "cg description", "description", "cg 내용", "내용", "작업 내용", "작업내용", "요청 사항", "요청사항", "cg desc", "desc"},
    "shot": {"shotname", "shot name", "shot", "샷", "샷명", "shot no", "shot id"},
    "scene": {"s#", "scene", "sc", "씬", "scene#", "s"},
    "cut": {"c#", "cut", "컷", "cut#", "c"},
}


def _norm_header(v) -> str:
    s = str(v if v is not None else "").replace("\n", " ").strip().lower()
    return re.sub(r"\s+", " ", s)


@dataclass
class ShotRow:
    sheet: str
    row: int
    no: str = ""
    ep: Optional[int] = None
    tc_in: Optional[TC] = None
    tc_out: Optional[TC] = None
    tc_raw: str = ""
    clips: list[str] = field(default_factory=list)
    clip_raw: str = ""
    desc: str = ""
    note: str = ""

    @property
    def label(self) -> str:
        ep = f"EP{self.ep:02d}" if self.ep is not None else (_safe(self.sheet) or "sheet")
        no = _safe(self.no) or f"r{self.row}"
        return f"{ep}_{no}"


def _safe(s) -> str:
    s = str(s).strip()
    s = re.sub(r"[\\/:*?\"<>|\s]+", "_", s)
    return s.strip("._")


def _cell_text(v) -> str:
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    if hasattr(v, "strftime"):
        # 엑셀이 "01:04:24" 같은 값을 시간으로 바꿔버린 경우
        try:
            return v.strftime("%H:%M:%S")
        except Exception:
            return str(v)
    return str(v)


def _parse_ep(v) -> Optional[int]:
    t = _cell_text(v)
    m = re.search(r"(\d{1,3})", t)
    return int(m.group(1)) if m else None


def ep_from_name(name: str) -> Optional[int]:
    """시트명/파일명에서 회차 번호 추출: EP04, EP_04, E04, 4화, 4회, 제4화 …"""
    pats = [
        r"(?i)(?<![A-Z0-9])EP\s*[_\-\.]?\s*(\d{1,3})(?!\d)",
        r"(?i)S\d{1,2}\s*E(\d{1,3})(?!\d)",
        r"(?i)(?<![A-Z0-9])E(\d{2,3})(?![0-9])",
        r"제?\s*(\d{1,3})\s*(?:화|회|부)(?!\d)",
        r"(?i)episode\s*(\d{1,3})",
    ]
    for p in pats:
        m = re.search(p, name)
        if m:
            return int(m.group(1))
    return None


CLIP_RE = re.compile(r"(?<![A-Z0-9])([A-Z])\s*(\d{2,4})\s*[_\-]?\s*C\s*(\d{2,4})(?!\d)", re.I)


def clip_tokens(text: str) -> list[str]:
    """'A027 C023, A028 C022' → ['A027C023', 'A028C022'] (대문자, 중복 제거).
    'A008 C013 ~ A008 C017' 처럼 같은 롤의 범위는 C013~C017로 펼친다."""
    text = text or ""
    out: list[str] = []
    matches = list(CLIP_RE.finditer(text))
    skip: set[int] = set()
    for i, m in enumerate(matches):
        if i in skip:
            continue
        roll, clip = f"{m.group(1).upper()}{m.group(2)}", int(m.group(3))
        toks = [f"{roll}C{m.group(3)}"]
        if i + 1 < len(matches):
            n = matches[i + 1]
            between = text[m.end():n.start()]
            same_roll = f"{n.group(1).upper()}{n.group(2)}" == roll
            if same_roll and re.fullmatch(r"\s*[~\-–—]\s*", between) and int(n.group(3)) > clip and int(n.group(3)) - clip <= 50:
                width = len(m.group(3))
                toks = [f"{roll}C{c:0{width}d}" for c in range(clip, int(n.group(3)) + 1)]
                skip.add(i + 1)
        for t in toks:
            if t not in out:
                out.append(t)
    return out


def _detect_header(ws) -> tuple[Optional[int], dict[str, int]]:
    """헤더 행 번호와 {역할: 열번호}를 돌려준다. 못 찾으면 (None, {})."""
    best: tuple[int, Optional[int], dict[str, int]] = (0, None, {})
    for r in range(1, min(ws.max_row, 40) + 1):
        roles: dict[str, int] = {}
        for c in range(1, min(ws.max_column, 60) + 1):
            h = _norm_header(ws.cell(r, c).value)
            if not h:
                continue
            for role, names in HEADER_ROLES.items():
                if h in names and role not in roles:
                    roles[role] = c
                    break
        score = len(roles) + (3 if "tc" in roles else 0) + (2 if "clip" in roles else 0)
        if ("tc" in roles or "clip" in roles) and score > best[0]:
            best = (score, r, roles)
    _, row, roles = best
    if row is None:
        return None, {}
    if "ep" not in roles and "ep_alt" in roles:
        roles["ep"] = roles["ep_alt"]
    roles.pop("ep_alt", None)
    return row, roles


def _guess_tc_column(ws, start_row: int = 1) -> Optional[int]:
    """헤더가 없을 때: 4자리(HH:MM:SS:FF) 타임코드가 가장 많은 열."""
    counts: dict[int, int] = {}
    for r in range(start_row, min(ws.max_row, 2000) + 1):
        for c in range(1, min(ws.max_column, 60) + 1):
            v = ws.cell(r, c).value
            if isinstance(v, str) and re.search(r"\d{1,2}[:;.]\d{2}[:;.]\d{2}[:;.,]\d{2,3}", v):
                counts[c] = counts.get(c, 0) + 1
    if not counts:
        return None
    col, n = max(counts.items(), key=lambda kv: kv[1])
    return col if n >= 1 else None


def parse_workbook(path: Path, sheets: Optional[list[str]] = None) -> list[ShotRow]:
    import openpyxl  # 지연 임포트 (설치 안내를 위해)

    wb = openpyxl.load_workbook(str(path), data_only=True)
    rows: list[ShotRow] = []
    for ws in wb.worksheets:
        if sheets and ws.title not in sheets:
            continue
        if ws.sheet_state != "visible" and not sheets:
            continue
        hdr, roles = _detect_header(ws)
        if hdr is None:
            col = _guess_tc_column(ws)
            if col is None:
                continue
            roles = {"tc": col}
            hdr = 0
        sheet_ep = ep_from_name(ws.title)
        tc_col, out_col, clip_col = roles.get("tc"), roles.get("tc_out"), roles.get("clip")
        for r in range(hdr + 1, ws.max_row + 1):
            tc_raw = _cell_text(ws.cell(r, tc_col).value) if tc_col else ""
            out_raw = _cell_text(ws.cell(r, out_col).value) if out_col else ""
            clip_raw = _cell_text(ws.cell(r, clip_col).value) if clip_col else ""
            no = _cell_text(ws.cell(r, roles["no"]).value) if "no" in roles else ""
            if not (tc_raw.strip() or clip_raw.strip() or no.strip()):
                continue  # 완전히 빈 행
            ep = _parse_ep(ws.cell(r, roles["ep"]).value) if "ep" in roles else None
            if ep is None:
                ep = sheet_ep
            desc = _cell_text(ws.cell(r, roles["desc"]).value) if "desc" in roles else ""
            tcs = TC.find_all(tc_raw)
            tc_in = tcs[0] if tcs else None
            tc_out = tcs[1] if len(tcs) > 1 else (TC.parse(out_raw) if out_raw else None)
            if "tc" in roles and hdr == 0 and tc_in is None:
                continue  # 헤더 없이 추정한 시트는 TC 있는 행만
            rows.append(ShotRow(
                sheet=ws.title, row=r, no=no.strip(), ep=ep, tc_in=tc_in, tc_out=tc_out,
                tc_raw=tc_raw.strip(), clips=clip_tokens(clip_raw), clip_raw=clip_raw.strip(),
                desc=desc.strip(),
            ))
    return rows


def parse_text(text: str, default_ep: Optional[int] = None, sheet: str = "TC") -> list[ShotRow]:
    """줄 단위 TC 목록. '01:04:24:08' / '01:04:24:08 ~ 01:04:30:00' / '이름 01:04:24:08' / 'EP04 3 01:04:24:08'."""
    rows: list[ShotRow] = []
    cur_ep = default_ep
    n = 0
    for i, line in enumerate(text.splitlines(), 1):
        raw = line.strip()
        if not raw or raw.startswith("#"):
            continue
        m_ep = re.search(r"(?i)(?<![A-Z0-9])EP\s*(\d{1,3})(?!\d)", raw)
        if m_ep and not TC_RE.search(raw):
            cur_ep = int(m_ep.group(1))
            continue
        tcs = TC.find_all(raw)
        if not tcs:
            continue
        n += 1
        head = TC_RE.split(raw)[0]
        name = head
        for pat in (r"(?i)(?<![A-Z0-9])EP\s*\d{1,3}(?!\d)", r"^\s*(?:no\.?|#)\s*", r"[\s,;:\-–~]+$"):
            name = re.sub(pat, " ", name)
        name = name.strip(" \t.-_:,")
        ep = int(m_ep.group(1)) if m_ep else cur_ep
        rows.append(ShotRow(sheet=sheet, row=i, no=name or f"{n:02d}", ep=ep, tc_in=tcs[0],
                            tc_out=tcs[1] if len(tcs) > 1 else None, tc_raw=raw, desc=raw))
    return rows


def parse_pdf(path: Path) -> list[ShotRow]:
    """PDF로 받은 CG 리스트: 페이지 글자에서 TC를 찾는다. 'EP04' 표기를 만나면 이후 행의 회차로 쓴다."""
    try:
        from pypdf import PdfReader
    except ImportError:
        raise CutError("[오류] PDF를 읽으려면 pypdf가 필요합니다:  pip install pypdf")
    rows: list[ShotRow] = []
    cur_ep = ep_from_name(path.stem)
    reader = PdfReader(str(path))
    pending_range = False  # 직전 줄이 'TC ~' 로 끝났거나 '~' 한 줄
    for pno, page in enumerate(reader.pages, 1):
        text = page.extract_text() or ""
        for line in text.splitlines():
            raw = line.strip()
            m_ep = re.search(r"(?i)(?<![A-Z0-9])EP\s*[_\-]?\s*(\d{1,3})(?!\d)", raw)
            if m_ep:
                cur_ep = int(m_ep.group(1))
            tcs = TC.find_all(raw)
            if not tcs:
                if raw in ("~", "-", "–"):
                    pending_range = True
                continue
            head = TC_RE.split(raw)[0]
            # 'TC ~' 다음 줄, 또는 '~ TC' 로 시작하는 줄 → 직전 행의 OUT
            if rows and rows[-1].tc_out is None and (pending_range or re.match(r"^[~\-–]\s*$", head)):
                rows[-1].tc_out = tcs[0]
                rows[-1].tc_raw += f" ~ {tcs[0]}"
                pending_range = False
                continue
            range_start = bool(re.search(r"[~\-–]\s*$", raw))
            m_no = re.match(r"\s*(\d{1,3}(?:-\d{1,2})?)\b", head)
            head_rest = re.sub(r"(?i)EP\s*[_\-]?\s*\d{1,3}", "", head)
            tail = TC_RE.split(raw)[-1]
            if not m_no and not range_start and (re.search(r"[^\s.\-_]", head_rest) or re.search(r"[가-힣A-Za-z]{2,}", tail)):
                continue  # 번호 없이 설명문 안에 들어 있는 TC(예: '(01:00:16:17 부터 그래픽 등장)')는 컷이 아님
            pending_range = range_start
            no = m_no.group(1) if m_no else ""
            rows.append(ShotRow(sheet=f"p{pno}", row=pno, no=no, ep=cur_ep, tc_in=tcs[0],
                                tc_out=tcs[1] if len(tcs) > 1 else None, tc_raw=raw, desc=raw))
    return rows


EDL_EVENT_RE = re.compile(
    r"^\s*(\d{1,6})\s+(\S+)\s+(\S+)\s+(C|D|W\d*|K\s*B?|K)\s*(\d*)\s+"
    r"(\d{1,2}[:;.]\d{2}[:;.]\d{2}[:;.]\d{2,3})\s+(\d{1,2}[:;.]\d{2}[:;.]\d{2}[:;.]\d{2,3})\s+"
    r"(\d{1,2}[:;.]\d{2}[:;.]\d{2}[:;.]\d{2,3})\s+(\d{1,2}[:;.]\d{2}[:;.]\d{2}[:;.]\d{2,3})\s*$")


def parse_edl(path: Path, use_source: bool = False) -> list[ShotRow]:
    """CMX3600 EDL. 기본은 레코드(타임라인) IN/OUT을 쓰고, use_source면 소스 IN/OUT.
    OUT은 EDL 관례대로 '마지막 프레임 다음'이므로 1프레임 당겨 포함 OUT으로 바꾼다.
    '* FROM CLIP NAME:' 또는 '* COMMENT:' 줄이 있으면 파일명에 넣는다."""
    text = path.read_text(encoding="utf-8-sig", errors="replace")
    rows: list[ShotRow] = []
    ep = ep_from_name(path.stem)
    title = ""
    for line in text.splitlines():
        raw = line.rstrip()
        if raw.upper().startswith("TITLE:"):
            title = raw[6:].strip()
            ep = ep_from_name(title) or ep
            continue
        m = EDL_EVENT_RE.match(raw)
        if m:
            num, reel = m.group(1), m.group(2)
            if reel.upper() in ("BL", "BLACK"):
                continue
            tc_in_s, tc_out_s = (m.group(6), m.group(7)) if use_source else (m.group(8), m.group(9))
            tc_in, tc_out = TC.parse(tc_in_s), TC.parse(tc_out_s)
            if tc_in is None or tc_out is None:
                continue
            if rows and rows[-1].no == num:  # 같은 이벤트의 A/B 트랙(디졸브 등)은 한 번만
                continue
            rows.append(ShotRow(sheet=path.stem, row=len(rows) + 1, no=num, ep=ep, tc_in=tc_in, tc_out=tc_out,
                                tc_raw=f"{tc_in} ~ {tc_out} (EDL)", desc=raw.strip()))
            rows[-1]._edl_out_exclusive = True  # type: ignore[attr-defined]
            continue
        mc = re.match(r"^\*\s*(?:FROM CLIP NAME|CLIP NAME|COMMENT|LOC)\s*:\s*(.+)$", raw, re.I)
        if mc and rows and not getattr(rows[-1], "_named", False):
            name = mc.group(1).strip()
            name = re.sub(r"\.[A-Za-z0-9]{2,4}$", "", name)  # 확장자 제거
            if mc.group(0).upper().startswith("* LOC"):
                name = name.split()[-1] if name.split() else name
            rows[-1].no = f"{rows[-1].no}_{name}"
            rows[-1]._named = True  # type: ignore[attr-defined]
    return rows


def load_shots(source: Optional[Path], tcs: Optional[list[str]] = None, sheets: Optional[list[str]] = None) -> list[ShotRow]:
    shots: list[ShotRow] = []
    if source:
        if not source.exists():
            raise CutError(f"[오류] 파일이 없습니다: {source}")
        ext = source.suffix.lower()
        if ext in (".xlsx", ".xlsm"):
            try:
                import openpyxl  # noqa: F401
            except ImportError:
                raise CutError("[오류] openpyxl이 없습니다:  pip install openpyxl")
            shots += parse_workbook(source, sheets)
        elif ext == ".pdf":
            shots += parse_pdf(source)
        elif ext == ".edl":
            shots += parse_edl(source, use_source=bool(getattr(load_shots, "edl_source", False)))
        elif ext in (".txt", ".csv", ".md"):
            shots += parse_text(source.read_text(encoding="utf-8-sig", errors="replace"), ep_from_name(source.stem), sheet=source.stem)
        else:
            raise CutError(f"[오류] 지원하지 않는 리스트 형식: {source.name} (xlsx / pdf / edl / txt)")
    if tcs:
        shots += parse_text("\n".join(tcs))
    return shots


# ---------------------------------------------------------------------------
# ffmpeg / ffprobe
# ---------------------------------------------------------------------------
def _bundled(name: str) -> Optional[str]:
    exe = name + (".exe" if os.name == "nt" else "")
    dirs = [Path(getattr(sys, "_MEIPASS", "")), Path(sys.executable).parent, Path(__file__).resolve().parent]
    for d in dirs:
        if str(d) and (d / exe).is_file():
            return str(d / exe)
        if str(d) and (d / "ffmpeg" / exe).is_file():
            return str(d / "ffmpeg" / exe)
    return None


def _find_tool(name: str, explicit: Optional[str]) -> str:
    cand = explicit or _bundled(name) or shutil.which(name)
    if not cand:
        if explicit is None and os.name == "nt":
            hint = "  winget install Gyan.FFmpeg  (설치 후 터미널을 다시 열 것)"
        else:
            hint = "  macOS: brew install ffmpeg / Ubuntu: sudo apt install ffmpeg"
        raise CutError(f"[오류] {name}을(를) 찾을 수 없습니다. ffmpeg를 설치하거나 --ffmpeg 옵션으로 경로를 지정하세요.\n{hint}")
    return cand


def _run(cmd: list[str], timeout: Optional[float] = None) -> subprocess.CompletedProcess:
    return subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=timeout)


@dataclass
class VideoInfo:
    path: Path
    fps: Fraction
    duration: float
    nb_frames: int
    width: int
    height: int
    codec: str
    start_tc: Optional[TC]  # 파일에 들어 있는 타임코드(없으면 None)
    color: dict[str, str] = field(default_factory=dict)

    @property
    def tc_fps(self) -> int:
        return tc_base(self.fps)


def probe(ffprobe: str, path: Path) -> VideoInfo:
    cp = _run([ffprobe, "-v", "error", "-print_format", "json", "-show_format", "-show_streams", str(path)])
    if cp.returncode != 0:
        raise RuntimeError(f"ffprobe 실패: {path}\n{cp.stderr.strip()}")
    data = json.loads(cp.stdout or "{}")
    streams = data.get("streams", [])
    v = next((s for s in streams if s.get("codec_type") == "video" and not s.get("disposition", {}).get("attached_pic")), None)
    if v is None:
        raise RuntimeError(f"비디오 스트림이 없습니다: {path}")
    fps_str = v.get("r_frame_rate") or v.get("avg_frame_rate") or "0/0"
    try:
        fps = Fraction(fps_str)
    except (ValueError, ZeroDivisionError):
        fps = Fraction(0)
    if fps <= 0:
        avg = v.get("avg_frame_rate", "0/0")
        try:
            fps = Fraction(avg)
        except (ValueError, ZeroDivisionError):
            fps = Fraction(0)
    if fps <= 0:
        raise RuntimeError(f"프레임레이트를 알 수 없습니다: {path}")
    fmt = data.get("format", {})
    duration = float(fmt.get("duration") or v.get("duration") or 0)
    nb = int(v.get("nb_frames") or 0)
    if nb <= 0:
        nb = int(round(duration * float(fps)))
    tc_text = (fmt.get("tags", {}) or {}).get("timecode")
    if not tc_text:
        for s in streams:
            t = (s.get("tags", {}) or {}).get("timecode")
            if t:
                tc_text = t
                break
    start_tc = TC.parse(tc_text) if tc_text else None
    color = {k: v[k] for k in ("color_primaries", "color_trc", "color_space", "color_range") if v.get(k) and v[k] != "unknown"}
    return VideoInfo(
        path=path, fps=fps, duration=duration, nb_frames=nb,
        width=int(v.get("width") or 0), height=int(v.get("height") or 0),
        codec=str(v.get("codec_name") or ""), start_tc=start_tc, color=color,
    )


_PTS_RE = re.compile(r"pts_time:\s*(-?\d+(?:\.\d+)?)")


def detect_scenes(ffmpeg: str, info: VideoInfo, threshold: float,
                  t0: Optional[float] = None, dur: Optional[float] = None) -> list[int]:
    """장면 전환 프레임 번호 목록(파일 시작 = 0). t0/dur을 주면 그 구간만 검사."""
    cmd = [ffmpeg, "-hide_banner", "-nostats", "-v", "info"]
    if t0 is not None:
        cmd += ["-ss", f"{max(0.0, t0):.6f}"]
    if dur is not None:
        cmd += ["-t", f"{max(0.0, dur):.6f}"]
    cmd += ["-i", str(info.path), "-an", "-sn", "-dn",
            "-vf", f"select='gt(scene,{threshold})',showinfo", "-f", "null", "-"]
    cp = _run(cmd)
    if cp.returncode != 0:
        raise RuntimeError(f"장면 검출 실패: {info.path}\n{cp.stderr.strip()[-800:]}")
    base = max(0.0, t0) if t0 is not None else 0.0
    fps = float(info.fps)
    frames = sorted({int(round((base + float(m.group(1))) * fps)) for m in _PTS_RE.finditer(cp.stderr)})
    return [f for f in frames if f > 0]


class SceneIndex:
    """영상별 장면 전환 캐시. 샷이 많으면 전체를 한 번 훑고, 적으면 TC 주변만 검사."""

    def __init__(self, ffmpeg: str, info: VideoInfo, threshold: float, cache_dir: Optional[Path]):
        self.ffmpeg, self.info, self.thr = ffmpeg, info, threshold
        self.cache_file = (cache_dir / f"{info.path.name}.scenes.json") if cache_dir else None
        self.full: Optional[list[int]] = None
        self.windows: list[tuple[int, int, list[int]]] = []  # (start, end, frames)
        self._load_cache()

    def _key(self) -> dict:
        st = self.info.path.stat()
        return {"threshold": self.thr, "size": st.st_size, "mtime": int(st.st_mtime), "fps": str(self.info.fps)}

    def _load_cache(self) -> None:
        if not self.cache_file or not self.cache_file.exists():
            return
        try:
            d = json.loads(self.cache_file.read_text(encoding="utf-8"))
            if d.get("key") == self._key():
                self.full = list(d.get("scenes", []))
        except Exception:
            pass

    def scan_full(self) -> None:
        if self.full is not None:
            return
        log(f"  장면 검출(전체 1회): {self.info.path.name} ({self.info.duration/60:.1f}분) …")
        t = time.time()
        self.full = detect_scenes(self.ffmpeg, self.info, self.thr)
        log(f"  → 전환 {len(self.full)}곳, {time.time()-t:.0f}초")
        if self.cache_file:
            try:
                self.cache_file.parent.mkdir(parents=True, exist_ok=True)
                self.cache_file.write_text(json.dumps({"key": self._key(), "scenes": self.full}), encoding="utf-8")
            except Exception:
                pass

    def around(self, frame: int, before: int, after: int) -> list[int]:
        """frame 앞뒤 [frame-before, frame+after] 구간의 전환 프레임."""
        lo, hi = max(0, frame - before), frame + after
        if self.full is not None:
            return [f for f in self.full if lo <= f <= hi]
        for s, e, fr in self.windows:
            if s <= lo and hi <= e:
                return [f for f in fr if lo <= f <= hi]
        fps = float(self.info.fps)
        fr = detect_scenes(self.ffmpeg, self.info, self.thr, t0=lo / fps, dur=(hi - lo + 2) / fps)
        self.windows.append((lo, hi, fr))
        return [f for f in fr if lo <= f <= hi]


# ---------------------------------------------------------------------------
# 컷 계획
# ---------------------------------------------------------------------------
@dataclass
class CutPlan:
    shot: ShotRow
    video: Optional[VideoInfo] = None
    in_frame: int = 0          # 파일 기준 프레임(포함)
    out_frame: int = 0         # 파일 기준 프레임(미포함)
    mode: str = ""             # shot / range / fixed
    out_path: Optional[Path] = None
    status: str = ""
    note: str = ""

    @property
    def n_frames(self) -> int:
        return max(0, self.out_frame - self.in_frame)


PART_RES = [r"(\d{1,3})\s*부(?![가-힣])", r"(?i)part\s*[_\-.]?\s*(\d{1,3})(?!\d)", r"(?i)(?<![A-Z])pt\.?\s*[_\-]?\s*(\d{1,3})(?!\d)",
            r"(?i)(?<![A-Z0-9])p(\d{1,2})(?![0-9A-Z])", r"[_\-\s](\d{1,2})$"]


def part_from_name(stem: str) -> tuple[Optional[int], str]:
    """'쩔해주세요_2부' → (2, '쩔해주세요_'). 부/part/pt/p2/끝 숫자. 못 찾으면 (None, stem)."""
    for pat in PART_RES:
        m = re.search(pat, stem)
        if m:
            return int(m.group(1)), (stem[:m.start()] + stem[m.end():]).strip(" _-")
    return None, stem


def detect_parts(videos: list[Path]) -> Optional[list[Path]]:
    """여러 파일이 같은 작품의 1부·2부·3부… 라면 순서대로 돌려준다. 아니면 None."""
    if len(videos) < 2:
        return None
    parsed = [(part_from_name(v.stem), v) for v in videos]
    if any(n is None for (n, _), _ in parsed):
        return None
    bases = {b for (_, b), _ in parsed}
    if len(bases) != 1:
        return None
    nums = sorted(n for (n, _), _ in parsed)
    if len(set(nums)) != len(nums):
        return None
    return [v for (_, _), v in sorted(parsed, key=lambda x: x[0][0])]


def chain_parts(infos: list[VideoInfo], shots: list[ShotRow], override: Optional[TC]) -> Optional[dict[Path, int]]:
    """부별 시작 프레임. 내장 TC가 있으면 그것, 없으면 앞 부에 이어 붙인다. 범위가 겹치면(각 부가 독립 TC) None."""
    starts: dict[Path, int] = {}
    prev_end: Optional[int] = None
    for i, info in enumerate(infos):
        base = info.tc_fps
        if i == 0:
            st, _ = _resolve_start_tc(info, shots, override)
        elif info.start_tc is not None:
            st = info.start_tc.to_frames(base)
        else:
            st = prev_end  # type: ignore[assignment]
        if prev_end is not None and st < prev_end:
            return None
        starts[info.path] = st
        prev_end = st + info.nb_frames
    return starts


def _resolve_start_tc(info: VideoInfo, shots: list[ShotRow], override: Optional[TC],
                      override_frames: Optional[int] = None) -> tuple[int, str]:
    """파일의 첫 프레임이 어느 타임코드인지(프레임 번호) 결정."""
    base = info.tc_fps
    if override_frames is not None:
        return override_frames, "부(part) 이어붙임"
    if override is not None:
        return override.to_frames(base), f"옵션 --start-tc {override}"
    if info.start_tc is not None:
        return info.start_tc.to_frames(base), f"파일 내장 TC {info.start_tc}"
    tcs = [s.tc_in.to_frames(base) for s in shots if s.tc_in]
    if tcs:
        lo = min(tcs)
        hour = 3600 * base
        if lo >= info.nb_frames and lo >= hour and (lo - hour) < info.nb_frames:
            return hour, "추정 01:00:00:00 (리스트 TC가 영상 길이를 넘어서 1시간 오프셋 가정)"
    return 0, "추정 00:00:00:00 (내장 TC 없음)"


def plan_cuts(shots: list[ShotRow], info: VideoInfo, args, scenes: Optional[SceneIndex],
              start_frames: Optional[int] = None) -> list[CutPlan]:
    base = info.tc_fps
    fps = float(info.fps)
    start, how = _resolve_start_tc(info, shots, args.start_tc_parsed, start_frames)
    log(f"  시작 TC: {TC.from_frames(start, base)} ({how}), {float(info.fps):.3f}fps, {info.nb_frames}프레임, {info.width}x{info.height} {info.codec}")
    handles = args.handles
    plans: list[CutPlan] = []
    for s in shots:
        p = CutPlan(shot=s, video=info)
        plans.append(p)
        if s.tc_in is None:
            p.status, p.note = "건너뜀", "TC 없음"
            continue
        if s.tc_in.f >= base or (s.tc_out and s.tc_out.f >= base):
            p.note = f"프레임 값({s.tc_in.f})이 {base}fps 범위를 넘음 — 리스트 fps 확인 필요; "
        F = s.tc_in.to_frames(base) - start
        if F < 0 or F >= info.nb_frames:
            p.status, p.note = "건너뜀", p.note + f"TC {s.tc_in}가 영상 범위 밖 (파일 {TC.from_frames(start, base)} ~ {TC.from_frames(start + info.nb_frames - 1, base)})"
            continue
        if s.tc_out is not None:
            O = s.tc_out.to_frames(base) - start + 1  # OUT TC는 마지막 프레임(포함)
            if getattr(s, "_edl_out_exclusive", False):
                O -= 1  # EDL의 OUT은 마지막 프레임 다음
            if O <= F:
                p.status, p.note = "건너뜀", p.note + f"OUT TC({s.tc_out})가 IN보다 앞섬"
                continue
            p.mode, p.in_frame, p.out_frame = "range", F, min(O, info.nb_frames)
        elif args.mode == "fixed":
            p.mode = "fixed"
            p.in_frame = max(0, F - int(round(args.before * fps)))
            p.out_frame = min(info.nb_frames, F + int(round(args.after * fps)))
        else:
            assert scenes is not None
            max_b, max_a = int(round(args.max_before * fps)), int(round(args.max_after * fps))
            cuts = scenes.around(F, max_b, max_a)
            # TC가 전환 직전 프레임(1~2프레임 오차)이면 다음 샷으로 간주
            snap = [c for c in cuts if F < c <= F + args.snap]
            anchor = snap[0] if snap else F
            before = [c for c in cuts if c <= anchor]
            after = [c for c in cuts if c > anchor]
            p.mode = "shot"
            if before:
                p.in_frame = before[-1]
            elif anchor - max_b <= 0:
                p.in_frame = 0  # 파일 첫 샷
            else:
                p.in_frame = anchor - max_b
                p.note += f"앞 전환 못 찾음(최대 {args.max_before:g}초로 자름); "
            if after:
                p.out_frame = after[0]
            elif anchor + max_a >= info.nb_frames:
                p.out_frame = info.nb_frames  # 파일 마지막 샷
            else:
                p.out_frame = anchor + max_a
                p.note += f"뒤 전환 못 찾음(최대 {args.max_after:g}초로 자름); "
            if snap:
                p.note += f"TC가 전환 {anchor - F}프레임 앞 → 다음 샷으로 맞춤; "
        if handles:
            p.in_frame = max(0, p.in_frame - handles)
            p.out_frame = min(info.nb_frames, p.out_frame + handles)
        if p.n_frames <= 0:
            p.status, p.note = "건너뜀", p.note + "길이 0"
            continue
        p.status = "예정"
    return plans


# ---------------------------------------------------------------------------
# ffmpeg 컷 실행
# ---------------------------------------------------------------------------
CODECS = {
    "prores": (".mov", ["-c:v", "prores_ks", "-profile:v", "3", "-vendor", "apl0", "-pix_fmt", "yuv422p10le", "-c:a", "pcm_s16le"]),
    "proreshq": (".mov", ["-c:v", "prores_ks", "-profile:v", "3", "-vendor", "apl0", "-pix_fmt", "yuv422p10le", "-c:a", "pcm_s16le"]),
    "prores422": (".mov", ["-c:v", "prores_ks", "-profile:v", "2", "-vendor", "apl0", "-pix_fmt", "yuv422p10le", "-c:a", "pcm_s16le"]),
    "h264": (".mp4", ["-c:v", "libx264", "-preset", "medium", "-crf", "16", "-pix_fmt", "yuv420p", "-c:a", "aac", "-b:a", "192k", "-movflags", "+faststart"]),
    "copy": (None, ["-c", "copy", "-avoid_negative_ts", "make_zero"]),
    "png": (None, []),
}


def output_path(p: CutPlan, out_dir: Path, codec: str, start_frames: int, prefix: str = "") -> Path:
    s, info = p.shot, p.video
    assert info is not None
    in_tc = TC.from_frames(start_frames + p.in_frame, info.tc_fps, drop=bool(s.tc_in and s.tc_in.drop))
    name = f"{s.label}_{in_tc.filename()}"
    if prefix:
        name = f"{_safe(prefix)}_{name}"
    ext = CODECS[codec][0]
    if codec == "copy":
        ext = info.path.suffix.lower() or ".mov"
    if codec == "png":
        return out_dir / name
    return out_dir / f"{name}{ext}"


def run_cut(ffmpeg: str, p: CutPlan, codec: str, start_frames: int, frame_start: int, no_audio: bool) -> None:
    info = p.video
    assert info is not None and p.out_path is not None
    fps = float(info.fps)
    base = info.tc_fps
    n = p.n_frames
    t_in = max(0.0, (p.in_frame - 0.5) / fps)  # 반 프레임 앞에서 찾기 시작 → 부동소수 오차로 한 프레임 밀리는 것 방지
    off = p.in_frame / fps - t_in              # 첫 프레임까지의 간격(보통 반 프레임) — 오디오를 이만큼 잘라 싱크를 맞춤
    dur = off + (n + 0.4) / fps
    in_tc = TC.from_frames(start_frames + p.in_frame, base, drop=bool(p.shot.tc_in and p.shot.tc_in.drop))
    cmd = [ffmpeg, "-hide_banner", "-nostats", "-loglevel", "error", "-y", "-ss", f"{t_in:.6f}", "-i", str(info.path), "-t", f"{dur:.6f}"]
    if codec == "png":
        p.out_path.mkdir(parents=True, exist_ok=True)
        cmd += ["-an", "-frames:v", str(n), "-vf", "setpts=PTS-STARTPTS", "-start_number", str(frame_start), "-f", "image2",
                str(p.out_path / f"{p.out_path.name}.%06d.png")]
    else:
        ext, cargs = CODECS[codec]
        cmd += list(cargs)
        if codec != "copy":
            # 타임스탬프를 0부터 다시 매겨 플레이어/편집툴에서 첫 프레임이 밀리거나 중복되지 않게
            cmd += ["-frames:v", str(n), "-vf", "setpts=PTS-STARTPTS",
                    "-af", f"atrim=start={off:.6f},asetpts=PTS-STARTPTS"]
            for k, v in info.color.items():
                cmd += [f"-{k}" if k != "color_space" else "-colorspace", v]
        if no_audio:
            cmd += ["-an"]
        if p.out_path.suffix.lower() in (".mov", ".mp4"):
            cmd += ["-timecode", str(in_tc)]
        cmd += ["-map_metadata", "-1", "-metadata", f"comment=cgcut {p.shot.sheet} no.{p.shot.no} TC {p.shot.tc_raw.replace(chr(10), ' ')}"]
        cmd += [str(p.out_path)]
    cp = _run(cmd)
    if cp.returncode != 0:
        raise RuntimeError(cp.stderr.strip()[-600:] or f"ffmpeg 종료코드 {cp.returncode}")


# ---------------------------------------------------------------------------
# 영상 파일 찾기 / 클립 복사
# ---------------------------------------------------------------------------
def _under(p: Path, d: Path) -> bool:
    try:
        p.resolve().relative_to(d.resolve())
        return True
    except ValueError:
        return False


def collect_videos(args) -> list[Path]:
    out_dir = Path(args.out)
    files: list[Path] = []
    for v in args.video or []:
        pth = Path(v)
        if pth.is_dir():
            files += sorted(p for p in pth.rglob("*") if p.suffix.lower() in VIDEO_EXTS)
        elif pth.exists():
            files.append(pth)
        else:
            raise CutError(f"[오류] 영상 파일이 없습니다: {pth}")
    if args.video_dir:
        d = Path(args.video_dir)
        if not d.is_dir():
            raise CutError(f"[오류] 폴더가 없습니다: {d}")
        files += sorted(p for p in d.rglob("*") if p.suffix.lower() in VIDEO_EXTS and not p.name.startswith("."))
    seen: set[Path] = set()
    out: list[Path] = []
    for f in files:
        rf = f.resolve()
        if _under(rf, out_dir):  # 이전 실행 결과물은 원본으로 쓰지 않음
            continue
        if rf not in seen:
            seen.add(rf)
            out.append(f)
    return out


def assign_videos(shots: list[ShotRow], videos: list[Path]) -> tuple[dict[Path, list[ShotRow]], list[ShotRow]]:
    """샷을 영상에 배정. (영상→샷들, 배정 못 한 샷들)"""
    by_video: dict[Path, list[ShotRow]] = {}
    unassigned: list[ShotRow] = []
    if not videos:
        return by_video, list(shots)
    if len(videos) == 1:
        by_video[videos[0]] = list(shots)
        return by_video, unassigned
    ep_map: dict[int, Path] = {}
    for v in videos:
        ep = ep_from_name(v.stem)
        if ep is None:
            continue
        if ep in ep_map:
            warn(f"EP{ep:02d} 영상이 여러 개: {ep_map[ep].name}, {v.name} → 앞의 것 사용")
            continue
        ep_map[ep] = v
    for s in shots:
        v = ep_map.get(s.ep) if s.ep is not None else None
        if v is None:
            unassigned.append(s)
        else:
            by_video.setdefault(v, []).append(s)
    return by_video, unassigned


def _norm_name(s: str) -> str:
    return re.sub(r"[^A-Z0-9]+", "_", s.upper())


def find_clip_files(clips_dir: Path, token: str) -> list[Path]:
    """'A027C023' 토큰에 맞는 카메라 원본 파일 찾기 (A027C023_…, A027_C023_…, A027_0913_C023 …)."""
    m = re.fullmatch(r"([A-Z])(\d+)C(\d+)", token)
    if not m:
        return []
    roll_letter, roll, clip = m.group(1), int(m.group(2)), int(m.group(3))
    pat = re.compile(rf"(?<![A-Z0-9]){roll_letter}0*{roll}(?!\d)[A-Z0-9_]*?C0*{clip}(?!\d)")
    found = []
    for p in clips_dir.rglob("*"):
        if p.is_file() and not p.name.startswith(".") and pat.search(_norm_name(p.stem)):
            found.append(p)
    return sorted(found)


# ---------------------------------------------------------------------------
# 보고서
# ---------------------------------------------------------------------------
def write_report(path: Path, plans: list[CutPlan], start_map: dict[Path, int]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["시트", "No", "EP", "리스트 TC", "영상", "모드", "IN TC", "OUT TC(마지막 프레임)", "프레임수", "길이(초)", "출력", "상태", "비고"])
        for p in plans:
            s = p.shot
            if p.video and p.status not in ("건너뜀",) and p.n_frames > 0:
                base = p.video.tc_fps
                st = start_map.get(p.video.path, 0)
                drop = bool(s.tc_in and s.tc_in.drop)
                in_tc = str(TC.from_frames(st + p.in_frame, base, drop))
                out_tc = str(TC.from_frames(st + p.out_frame - 1, base, drop))
                nf, secs = p.n_frames, f"{p.n_frames / float(p.video.fps):.2f}"
            else:
                in_tc = out_tc = ""
                nf, secs = "", ""
            w.writerow([s.sheet, s.no, s.ep if s.ep is not None else "", s.tc_raw.replace("\n", " "),
                        p.video.path.name if p.video else "", p.mode, in_tc, out_tc, nf, secs,
                        p.out_path.name if p.out_path else "", p.status, p.note.strip("; ")])


# ---------------------------------------------------------------------------
# main
# ---------------------------------------------------------------------------
def build_parser() -> argparse.ArgumentParser:
    ap = argparse.ArgumentParser(
        prog="cgcut",
        description="CG 리스트(xlsx)의 TC를 읽어 원본 영상에서 해당 샷만 잘라냅니다.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "예시:\n"
            "  cgcut CG리스트.xlsx --video 편집본.mov -o out/\n"
            "  cgcut CG리스트.xlsx --video-dir ./편집본/ -o out/ --codec prores --handles 12\n"
            "  cgcut CG리스트.xlsx --dry-run\n"
            "  cgcut CG리스트.xlsx --clips-dir ./카메라원본/ -o out/\n"
        ),
    )
    ap.add_argument("xlsx", nargs="?", metavar="리스트", help="CG 리스트 파일 (.xlsx / .pdf / .edl / .txt) — --tc만 쓸 때는 생략")
    ap.add_argument("--edl-source", action="store_true", help="EDL의 레코드(타임라인) TC 대신 소스 TC를 사용")
    ap.add_argument("--no-parts", action="store_true", help="1부·2부… 파일을 한 편으로 이어 보는 기능 끄기")
    ap.add_argument("--tc", action="append", help="TC 직접 지정 (여러 번). 예: --tc 01:04:24:08  --tc \"01:04:24:08 ~ 01:04:30:00\"  --tc \"창문 01:04:24:08\"")
    ap.add_argument("--prefix", default="", help="출력 파일명 앞에 붙일 이름 (예: 쩔해주세요)")
    ap.add_argument("-o", "--out", default="cgcut_out", help="출력 폴더 (기본: cgcut_out)")
    ap.add_argument("--video", action="append", help="원본 영상 파일 (여러 번 지정 가능, 폴더도 가능)")
    ap.add_argument("--video-dir", help="원본 영상 폴더 — 파일명의 EP01/1화 등으로 회차 매칭")
    ap.add_argument("--sheet", action="append", help="특정 시트만 처리 (여러 번 지정 가능)")
    ap.add_argument("--mode", choices=["shot", "fixed"], default="shot",
                    help="shot: 장면 전환을 찾아 샷 전체를 자름(기본) / fixed: TC 기준 고정 길이")
    ap.add_argument("--before", type=float, default=1.0, help="fixed 모드: TC 앞 몇 초 (기본 1)")
    ap.add_argument("--after", type=float, default=5.0, help="fixed 모드: TC 뒤 몇 초 (기본 5)")
    ap.add_argument("--handles", type=int, default=0, help="앞뒤 여유 프레임(핸들) 수 (기본 0)")
    ap.add_argument("--scene-threshold", type=float, default=0.35, help="장면 전환 감도 0~1, 낮을수록 민감 (기본 0.35)")
    ap.add_argument("--max-before", type=float, default=30.0, help="shot 모드: TC 앞으로 최대 몇 초까지 (기본 30)")
    ap.add_argument("--max-after", type=float, default=30.0, help="shot 모드: TC 뒤로 최대 몇 초까지 (기본 30)")
    ap.add_argument("--snap", type=int, default=2, help="TC가 전환 n프레임 앞이면 다음 샷으로 간주 (기본 2, 0이면 끔)")
    ap.add_argument("--codec", choices=list(CODECS), default="prores",
                    help="출력 코덱: prores(422 HQ, 기본) / prores422 / h264(mp4) / copy(재인코딩 없음, 키프레임 단위라 부정확) / png(이미지 시퀀스)")
    ap.add_argument("--frame-start", type=int, default=1001, help="png 모드 시작 프레임 번호 (기본 1001)")
    ap.add_argument("--no-audio", action="store_true", help="오디오 제외")
    ap.add_argument("--start-tc", help="영상 첫 프레임의 타임코드 지정 (예: 01:00:00:00). 기본은 파일 내장 TC")
    ap.add_argument("--clips-dir", help="카메라 원본 폴더 — 리스트의 Scan Name(클립명)과 같은 파일을 찾아 복사")
    ap.add_argument("--dry-run", action="store_true", help="자르지 않고 계획만 출력(보고서 CSV는 생성)")
    ap.add_argument("--overwrite", action="store_true", help="이미 있는 출력 파일도 다시 만듦")
    ap.add_argument("--ffmpeg", help="ffmpeg 실행 파일 경로")
    ap.add_argument("--ffprobe", help="ffprobe 실행 파일 경로")
    ap.add_argument("--version", action="version", version=f"cgcut {__version__}")
    return ap


def drop_args(paths: list[str]) -> list[str]:
    """드래그앤드롭(cgcut_drag.bat)으로 받은 파일들을 옵션으로 바꾼다.
    .xlsx → 리스트, 영상 파일 → --video, 폴더 → --video-dir. 출력은 리스트 옆 cgcut_out 폴더."""
    xlsx = [p for p in paths if p.lower().endswith((".xlsx", ".xlsm"))]
    if len(xlsx) != 1:
        raise CutError("[오류] CG 리스트 엑셀(.xlsx) 1개와 원본 영상(또는 영상 폴더)을 함께 끌어다 놓으세요.")
    out = [xlsx[0], "-o", str(Path(xlsx[0]).resolve().parent / "cgcut_out")]
    for p in paths:
        if p in xlsx:
            continue
        if Path(p).is_dir():
            out += ["--video-dir", p]
        elif Path(p).suffix.lower() in VIDEO_EXTS:
            out += ["--video", p]
        else:
            warn(f"무시: {p}")
    return out


def main(argv: Optional[list[str]] = None) -> int:
    argv = list(sys.argv[1:] if argv is None else argv)
    if argv and argv[0] == "--drop":
        rest = argv[1:]
        files = [a for a in rest if not a.startswith("-") and Path(a).exists()]
        opts = [a for a in rest if a not in files]
        argv = drop_args(files) + opts
    args = build_parser().parse_args(argv)
    try:
        return run(args)
    except CutError as e:
        print(str(e), file=sys.stderr, flush=True)
        return 2


def run(args, cancel=None) -> int:
    """args: build_parser()로 만든 네임스페이스. cancel: 호출하면 True를 돌려주는 함수(중단 요청)."""
    if not args.xlsx and not args.tc:
        raise CutError("[오류] CG 리스트 파일(.xlsx/.pdf/.txt)이나 --tc 를 지정하세요.")
    xlsx = Path(args.xlsx) if args.xlsx else None
    out_dir = Path(args.out)
    args.start_tc_parsed = TC.parse(args.start_tc) if args.start_tc else None
    if args.start_tc and args.start_tc_parsed is None:
        raise CutError(f"[오류] --start-tc 형식이 잘못됐습니다: {args.start_tc} (예: 01:00:00:00)")

    load_shots.edl_source = bool(getattr(args, "edl_source", False))  # type: ignore[attr-defined]
    shots = load_shots(xlsx, args.tc, args.sheet)
    n_tc = sum(1 for s in shots if s.tc_in)
    n_clip = sum(1 for s in shots if s.clips)
    log(f"리스트: {xlsx.name if xlsx else '직접 입력'} — 행 {len(shots)}개 (TC 있음 {n_tc}, 클립명 있음 {n_clip})")
    if not shots:
        raise CutError("[오류] TC나 클립명이 있는 행을 찾지 못했습니다. 헤더에 'TC' 또는 'Scan Name' 열이 있는지 확인하세요.")

    out_dir.mkdir(parents=True, exist_ok=True)
    rc = 0
    if n_tc == 0 and n_clip > 0 and not args.clips_dir:
        log("이 리스트에는 TC가 없고 클립명(Scan Name)만 있습니다. 카메라 원본 폴더를 주면 클립을 찾아 모아 드립니다:")
        log(f"  cgcut \"{xlsx}\" --clips-dir 카메라원본폴더 -o {out_dir}")
        return 0

    # --- 클립명 모드 -------------------------------------------------------
    if args.clips_dir:
        cdir = Path(args.clips_dir)
        if not cdir.is_dir():
            raise CutError(f"[오류] 폴더가 없습니다: {cdir}")
        dest = out_dir / "clips"
        rows_out = []
        log(f"\n클립 찾기: {cdir}")
        for s in shots:
            if not s.clips:
                continue
            for tok in s.clips:
                files = find_clip_files(cdir, tok)
                if not files:
                    log(f"  {s.label:<14} {tok:<10} → 못 찾음")
                    rows_out.append([s.sheet, s.no, s.ep or "", tok, "", "못 찾음"])
                    continue
                for f in files:
                    target = dest / f"{s.label}_{f.name}"
                    status = "예정" if args.dry_run else "복사"
                    if not args.dry_run:
                        dest.mkdir(parents=True, exist_ok=True)
                        if target.exists() and not args.overwrite and target.stat().st_size == f.stat().st_size:
                            status = "이미 있음"
                        else:
                            shutil.copy2(f, target)
                    log(f"  {s.label:<14} {tok:<10} → {f.name} ({status})")
                    rows_out.append([s.sheet, s.no, s.ep or "", tok, str(f), status])
        rep = out_dir / "cgcut_clips_report.csv"
        with rep.open("w", newline="", encoding="utf-8-sig") as fh:
            w = csv.writer(fh)
            w.writerow(["시트", "No", "EP", "클립", "원본 파일", "상태"])
            w.writerows(rows_out)
        log(f"클립 보고서: {rep}")
        if n_tc == 0:
            return 0

    # --- TC 모드 -----------------------------------------------------------
    tc_shots = [s for s in shots if s.tc_in or not s.clips]
    videos = collect_videos(args)
    if not videos:
        if args.dry_run:
            log("\n영상 없이 리스트만 확인합니다 (--video 또는 --video-dir 지정 시 실제 위치 계산).")
            for s in tc_shots:
                tc = f"{s.tc_in}" + (f" ~ {s.tc_out}" if s.tc_out else "") if s.tc_in else "(TC 없음)"
                log(f"  {s.sheet:<8} no.{s.no:<5} EP{s.ep if s.ep is not None else '?':<4} {tc:<26} {s.desc[:40].replace(chr(10), ' ')}")
            return 0
        raise CutError("[오류] 원본 영상을 지정하세요: --video 파일 또는 --video-dir 폴더 (계획만 보려면 --dry-run)")

    ffmpeg = _find_tool("ffmpeg", args.ffmpeg)
    ffprobe = _find_tool("ffprobe", args.ffprobe or (str(Path(ffmpeg).with_name("ffprobe" + Path(ffmpeg).suffix)) if args.ffmpeg else None))

    plans: list[CutPlan] = []
    start_map: dict[Path, int] = {}
    part_starts: dict[Path, int] = {}
    infos: dict[Path, VideoInfo] = {}
    parts = None if getattr(args, "no_parts", False) else detect_parts(videos)
    if parts:
        # 1부·2부·3부… 를 이어진 한 편으로: TC가 어느 부에 들어가는지로 배정
        try:
            pinfos = [probe(ffprobe, v) for v in parts]
        except RuntimeError as e:
            raise CutError(str(e))
        chain = chain_parts(pinfos, tc_shots, args.start_tc_parsed)
        if chain:
            infos.update({i.path: i for i in pinfos})
            part_starts = chain
            log("부(part)로 나뉜 파일을 한 편으로 인식:")
            for i in pinfos:
                st = chain[i.path]
                log(f"  {i.path.name}: {TC.from_frames(st, i.tc_fps)} ~ {TC.from_frames(st + i.nb_frames - 1, i.tc_fps)}")
            by_video: dict[Path, list[ShotRow]] = {}
            unassigned: list[ShotRow] = []
            for s in tc_shots:
                if s.tc_in is None:
                    unassigned.append(s)
                    continue
                f = s.tc_in.to_frames(pinfos[0].tc_fps)
                hit = next((i for i in pinfos if chain[i.path] <= f < chain[i.path] + i.nb_frames), None)
                (by_video.setdefault(hit.path, []) if hit else unassigned).append(s)
            for s in unassigned:
                plans.append(CutPlan(shot=s, status="건너뜀", note="TC 없음" if s.tc_in is None else f"TC {s.tc_in}가 어느 부에도 없음"))
        else:
            log("부 파일들의 TC 범위가 겹쳐 각각 별개 영상으로 처리합니다.")
            parts = None
    if not parts:
        by_video, unassigned = assign_videos(tc_shots, videos)
        for s in unassigned:
            plans.append(CutPlan(shot=s, status="건너뜀", note=f"EP{s.ep}에 해당하는 영상 파일 없음" if s.ep is not None else "회차를 알 수 없어 영상 배정 불가"))

    for vpath, vshots in by_video.items():
        log(f"\n영상: {vpath}")
        try:
            info = infos.get(vpath) or probe(ffprobe, vpath)
        except RuntimeError as e:
            warn(str(e))
            for s in vshots:
                plans.append(CutPlan(shot=s, status="실패", note="ffprobe 실패"))
            rc = 1
            continue
        scenes = None
        if args.mode == "shot" and not args.dry_run:
            scenes = SceneIndex(ffmpeg, info, args.scene_threshold, out_dir / ".cache")
            n_need = sum(1 for s in vshots if s.tc_in and s.tc_out is None)
            if n_need * (args.max_before + args.max_after) > info.duration * 0.6 or n_need >= 8:
                scenes.scan_full()
        elif args.mode == "shot" and args.dry_run:
            # dry-run: 장면 검출 없이 TC 위치만 확인
            class _NoScenes:
                def around(self, frame, before, after):
                    return []
            scenes = _NoScenes()  # type: ignore[assignment]
        vplans = plan_cuts(vshots, info, args, scenes, part_starts.get(vpath))
        start_map[vpath] = _resolve_start_tc(info, vshots, args.start_tc_parsed, part_starts.get(vpath))[0]
        for p in vplans:
            if p.status == "예정":
                p.out_path = output_path(p, out_dir, args.codec, start_map[vpath], getattr(args, "prefix", ""))
        plans += vplans

    todo = [p for p in plans if p.status == "예정"]
    log(f"\n자를 샷: {len(todo)}개 / 건너뜀 {sum(1 for p in plans if p.status == '건너뜀')}개")
    for p in plans:
        s = p.shot
        if p.status == "예정" and p.video:
            st = start_map[p.video.path]
            base = p.video.tc_fps
            rng = f"{TC.from_frames(st + p.in_frame, base)} ~ {TC.from_frames(st + p.out_frame - 1, base)}"
            log(f"  {s.label:<14} {str(s.tc_in):<12} {p.mode:<5} {rng}  {p.n_frames:>5}f  {p.note.strip('; ')}")
        else:
            log(f"  {s.label:<14} {str(s.tc_in) if s.tc_in else '-':<12} {p.status:<5} {p.note.strip('; ')}")

    if args.dry_run:
        rep = out_dir / "cgcut_report.csv"
        write_report(rep, plans, start_map)
        log(f"\n(dry-run) 보고서: {rep}" + ("  ※ shot 모드의 실제 IN/OUT은 장면 검출 후 정해집니다." if args.mode == "shot" else ""))
        return rc

    done = 0
    t0 = time.time()
    for i, p in enumerate(todo, 1):
        assert p.out_path is not None and p.video is not None
        if cancel and cancel():
            log("중단됨.")
            rc = 3
            break
        exists = p.out_path.exists() if args.codec != "png" else (p.out_path.is_dir() and any(p.out_path.iterdir()))
        if exists and not args.overwrite:
            p.status, p.note = "이미 있음", p.note + "출력 파일 있음(--overwrite로 다시 만듦)"
            log(f"[{i}/{len(todo)}] {p.out_path.name} — 이미 있음, 건너뜀")
            continue
        log(f"[{i}/{len(todo)}] {p.out_path.name} ({p.n_frames}f, {p.n_frames / float(p.video.fps):.1f}s) …")
        try:
            run_cut(ffmpeg, p, args.codec, start_map[p.video.path], args.frame_start, args.no_audio)
            p.status = "완료"
            done += 1
        except RuntimeError as e:
            p.status, p.note = "실패", p.note + str(e).replace("\n", " ")[:200]
            warn(f"{p.out_path.name}: {e}")
            rc = 1
    rep = out_dir / "cgcut_report.csv"
    write_report(rep, plans, start_map)
    log(f"\n완료 {done}개, {time.time() - t0:.0f}초. 출력: {out_dir.resolve()}\n보고서: {rep}")
    return rc


if __name__ == "__main__":
    sys.exit(main())
