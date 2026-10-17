#!/usr/bin/env python3
"""테스트용 샘플 생성: 장면 전환 위치를 아는 합성 영상 + 쩔해주세요식 CG 리스트 xlsx.

    python tests/make_sample.py sample/
→ sample/EP01_sample.mov (내장 TC 01:00:00:00), sample/EP02_sample.mp4 (내장 TC 없음), sample/CG_list_sample.xlsx
"""
from __future__ import annotations

import subprocess
import sys
from pathlib import Path

FPS = 24
# 각 샷의 프레임 수. 전환 위치 = 누적합: 48, 120, 150, 246, 306, 426, 466, 666, 끝 722
SEGMENTS = [48, 72, 30, 96, 60, 120, 40, 200, 56]
# 각 샷의 회색 밝기(0~255). 인접 샷끼리 차이가 커서 장면 검출에 걸린다.
GRAYS = [20, 230, 80, 200, 40, 160, 100, 240, 60]
W, H = 320, 180


def boundaries() -> list[int]:
    acc, out = 0, []
    for n in SEGMENTS:
        acc += n
        out.append(acc)
    return out


def make_video(path: Path, timecode: str | None) -> None:
    cmd = ["ffmpeg", "-hide_banner", "-loglevel", "error", "-y"]
    parts = []
    for i, (n, g) in enumerate(zip(SEGMENTS, GRAYS)):
        cmd += ["-f", "lavfi", "-i", f"color=c=0x{g:02x}{g:02x}{g:02x}:s={W}x{H}:r={FPS}"]
        parts.append(f"[{i}:v]trim=end_frame={n},setpts=PTS-STARTPTS[v{i}]")
    n_in = len(SEGMENTS)
    cmd += ["-f", "lavfi", "-i", "sine=frequency=440:sample_rate=48000"]
    fc = ";".join(parts) + ";" + "".join(f"[v{i}]" for i in range(n_in)) + f"concat=n={n_in}:v=1:a=0[v]"
    cmd += ["-filter_complex", fc, "-map", "[v]", "-map", f"{n_in}:a", "-shortest",
            "-c:v", "libx264", "-preset", "veryfast", "-crf", "12", "-pix_fmt", "yuv420p",
            # 키프레임이 장면 전환과 겹치지 않게 (정확한 탐색 검증용)
            "-x264-params", "keyint=37:min-keyint=37:scenecut=0",
            "-c:a", "aac", "-b:a", "96k"]
    if timecode:
        cmd += ["-timecode", timecode]
    cmd += [str(path)]
    subprocess.run(cmd, check=True)


def make_xlsx(path: Path) -> None:
    import openpyxl

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    def sheet(name: str, rows: list[tuple[str, str, str]]):
        ws = wb.create_sheet(name)
        ws["A1"] = f"샘플_{name}"
        ws.merge_cells("A1:H1")
        ws["F2"], ws["G2"], ws["H2"] = "Total", "Done", "Progress"
        ws["F4"], ws["G4"], ws["H4"] = "COMP", "MOTION", "3D"
        ws.append([])
        for c, h in enumerate(["no.", "EP", "Shotname", "Work Order", "TC", "Comp Status", "INDEX", "Worker"], 1):
            ws.cell(6, c, h)
        r = 7
        for no, tc, desc in rows:
            ws.cell(r, 1, no)
            ws.cell(r, 2, name)
            ws.cell(r, 4, desc)
            ws.cell(r, 5, tc)
            ws.cell(r, 7, "C")
            r += 1

    sheet("EP01", [
        ("1", "01:00:02:00", "샷2 첫 프레임 (48)"),
        ("2", "01:00:08:08", "샷4 중간 (200)"),
        ("3", "01:00:10:06\n~\n01:00:12:05", "범위 지정 246~293"),
        ("4", "01:00:04:23", "샷2 마지막 프레임(119) → 다음 샷으로 스냅"),
        ("5", "02:00:00:00", "범위 밖"),
        ("6", "00:00:05:00", "시작 TC 이전"),
        ("7", "", "TC 없음"),
        ("8", "01:00:00:00", "파일 첫 샷"),
        ("9", "01:00:29:10", "파일 마지막 샷 (706)"),
    ])
    sheet("EP02", [
        ("1", "01:00:02:00", "내장 TC 없는 mp4 — 01:00:00:00 추정"),
        ("2-1", "01:00:17:18;", "드롭프레임 표기 아님(세미콜론 뒤) → 426 샷7"),
    ])
    wb.save(path)


def main() -> None:
    out = Path(sys.argv[1] if len(sys.argv) > 1 else "sample")
    out.mkdir(parents=True, exist_ok=True)
    make_video(out / "EP01_sample.mov", "01:00:00:00")
    make_video(out / "EP02_sample.mp4", None)
    make_xlsx(out / "CG_list_sample.xlsx")
    print(f"생성: {out.resolve()}  전환 프레임: {boundaries()}")


if __name__ == "__main__":
    main()
