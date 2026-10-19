#!/usr/bin/env python3
"""cgcut GUI — 원본 영상 + (TC 직접 입력 / CG 리스트 xlsx·pdf) → 샷 자르기. exe 빌드용 진입점."""
from __future__ import annotations

import os
import queue
import subprocess
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import cgcut

APP = "cgcut — CG 샷 자르기"
VIDEO_TYPES = [("영상", "*.mov *.mp4 *.mxf *.mkv *.avi *.m4v *.mts *.m2ts"), ("모든 파일", "*.*")]
LIST_TYPES = [("CG 리스트", "*.xlsx *.xlsm *.pdf *.edl *.txt"), ("EDL", "*.edl"), ("모든 파일", "*.*")]
CODECS = [("ProRes 422 HQ (.mov) — CG 작업용", "prores"), ("ProRes 422 (.mov)", "prores422"),
          ("H.264 (.mp4) — 확인용", "h264"), ("PNG 이미지 시퀀스", "png"), ("재인코딩 없음 (빠름, 앞뒤 몇 프레임 오차)", "copy")]


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP)
        self.minsize(760, 640)
        try:
            self.tk.call("tk", "scaling", 1.25)
        except Exception:
            pass
        self.q: queue.Queue[str] = queue.Queue()
        self.worker: threading.Thread | None = None
        self.cancel_flag = False
        self.videos: list[str] = []
        self._build()
        self.after(100, self._pump)

    # ------------------------------------------------------------------ UI
    def _build(self):
        pad = {"padx": 10, "pady": 4}
        root = ttk.Frame(self, padding=10)
        root.pack(fill="both", expand=True)
        root.columnconfigure(1, weight=1)

        # 1. 원본 영상
        ttk.Label(root, text="1. 원본 영상", font=("", 11, "bold")).grid(row=0, column=0, columnspan=3, sticky="w", **pad)
        self.video_var = tk.StringVar(value="(선택 안 됨)")
        ttk.Label(root, textvariable=self.video_var, wraplength=520).grid(row=1, column=1, sticky="w", padx=10)
        b = ttk.Frame(root)
        b.grid(row=1, column=2, sticky="e", padx=10)
        ttk.Button(b, text="영상 파일…", command=self.pick_videos).pack(side="left")
        ttk.Button(b, text="폴더…", command=self.pick_video_dir).pack(side="left", padx=(6, 0))

        # 2. TC
        ttk.Label(root, text="2. 자를 TC", font=("", 11, "bold")).grid(row=2, column=0, columnspan=3, sticky="w", **pad)
        ttk.Label(root, text="리스트 파일").grid(row=3, column=0, sticky="w", padx=10)
        self.list_var = tk.StringVar()
        ttk.Entry(root, textvariable=self.list_var).grid(row=3, column=1, sticky="ew", padx=10)
        ttk.Button(root, text="xlsx / pdf / edl…", command=self.pick_list).grid(row=3, column=2, sticky="e", padx=10)
        ttk.Label(root, text="직접 입력\n(한 줄에 하나)").grid(row=4, column=0, sticky="nw", padx=10, pady=4)
        self.tc_text = tk.Text(root, height=6, font=("Consolas", 11), undo=True)
        self.tc_text.grid(row=4, column=1, columnspan=2, sticky="ew", padx=10, pady=4)
        ttk.Label(root, foreground="#666", text=(
            "예)  01:04:24:08          ← 그 TC가 있는 샷을 통째로\n"
            "      01:36:10:07 ~ 01:36:20:03   ← 이 구간만\n"
            "      창문리무브 01:04:24:08   ← 앞에 이름을 쓰면 파일명에 들어감   /   EP04 줄을 넣으면 그 아래는 EP04 영상에서")
        ).grid(row=5, column=1, columnspan=2, sticky="w", padx=10)

        # 3. 옵션
        ttk.Label(root, text="3. 옵션", font=("", 11, "bold")).grid(row=6, column=0, columnspan=3, sticky="w", **pad)
        opt = ttk.Frame(root)
        opt.grid(row=7, column=0, columnspan=3, sticky="ew", padx=10)
        for c in (1, 3):
            opt.columnconfigure(c, weight=1)

        ttk.Label(opt, text="저장 폴더").grid(row=0, column=0, sticky="w")
        self.out_var = tk.StringVar()
        ttk.Entry(opt, textvariable=self.out_var).grid(row=0, column=1, columnspan=2, sticky="ew", padx=6)
        ttk.Button(opt, text="…", width=3, command=self.pick_out).grid(row=0, column=3, sticky="w")

        ttk.Label(opt, text="파일명 앞머리").grid(row=1, column=0, sticky="w", pady=4)
        self.prefix_var = tk.StringVar()
        ttk.Entry(opt, textvariable=self.prefix_var).grid(row=1, column=1, sticky="ew", padx=6)
        ttk.Label(opt, text="형식").grid(row=1, column=2, sticky="e")
        self.codec_var = tk.StringVar(value=CODECS[0][0])
        ttk.Combobox(opt, textvariable=self.codec_var, values=[c[0] for c in CODECS], state="readonly", width=34).grid(row=1, column=3, sticky="ew")

        ttk.Label(opt, text="자르는 방식").grid(row=2, column=0, sticky="w")
        self.mode_var = tk.StringVar(value="shot")
        mf = ttk.Frame(opt)
        mf.grid(row=2, column=1, columnspan=3, sticky="w", padx=6)
        ttk.Radiobutton(mf, text="샷 전체 (장면 전환 자동 감지)", variable=self.mode_var, value="shot", command=self._mode).pack(side="left")
        ttk.Radiobutton(mf, text="TC 기준 고정 길이", variable=self.mode_var, value="fixed", command=self._mode).pack(side="left", padx=(12, 4))
        self.before_var, self.after_var = tk.StringVar(value="1"), tk.StringVar(value="5")
        self.fixed_box = ttk.Frame(mf)
        ttk.Label(self.fixed_box, text="앞").pack(side="left")
        ttk.Entry(self.fixed_box, textvariable=self.before_var, width=4).pack(side="left", padx=2)
        ttk.Label(self.fixed_box, text="초  뒤").pack(side="left")
        ttk.Entry(self.fixed_box, textvariable=self.after_var, width=4).pack(side="left", padx=2)
        ttk.Label(self.fixed_box, text="초").pack(side="left")

        self.edl_src_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(opt, text="EDL은 레코드(타임라인) TC 대신 소스 TC 사용", variable=self.edl_src_var).grid(row=5, column=0, columnspan=4, sticky="w")
        ttk.Label(opt, text="핸들(여유 프레임)").grid(row=3, column=0, sticky="w", pady=4)
        self.handles_var = tk.StringVar(value="0")
        ttk.Entry(opt, textvariable=self.handles_var, width=6).grid(row=3, column=1, sticky="w", padx=6)
        ttk.Label(opt, text="영상 첫 프레임 TC").grid(row=3, column=2, sticky="e")
        self.start_var = tk.StringVar()
        e = ttk.Entry(opt, textvariable=self.start_var, width=14)
        e.grid(row=3, column=3, sticky="w")
        ttk.Label(opt, foreground="#666", text="비우면 파일에 든 TC 사용 (없으면 01:00:00:00 추정)").grid(row=4, column=2, columnspan=2, sticky="e")

        # 실행
        run = ttk.Frame(root)
        run.grid(row=8, column=0, columnspan=3, sticky="ew", padx=10, pady=(10, 4))
        self.run_btn = ttk.Button(run, text="▶  자르기", command=self.start)
        self.run_btn.pack(side="left")
        self.stop_btn = ttk.Button(run, text="중단", command=self.stop, state="disabled")
        self.stop_btn.pack(side="left", padx=6)
        ttk.Button(run, text="저장 폴더 열기", command=self.open_out).pack(side="left", padx=6)
        self.status_var = tk.StringVar(value="")
        ttk.Label(run, textvariable=self.status_var).pack(side="left", padx=10)

        self.log = tk.Text(root, height=14, font=("Consolas", 10), state="disabled", wrap="none")
        self.log.grid(row=9, column=0, columnspan=3, sticky="nsew", padx=10, pady=(0, 6))
        root.rowconfigure(9, weight=1)
        sb = ttk.Scrollbar(root, command=self.log.yview)
        sb.grid(row=9, column=3, sticky="ns")
        self.log.configure(yscrollcommand=sb.set)
        self._mode()

    def _mode(self):
        if self.mode_var.get() == "fixed":
            self.fixed_box.pack(side="left")
        else:
            self.fixed_box.pack_forget()

    # ------------------------------------------------------------------ pickers
    def pick_videos(self):
        fs = filedialog.askopenfilenames(title="원본 영상 선택", filetypes=VIDEO_TYPES)
        if fs:
            self.videos = list(fs)
            self._after_pick()

    def pick_video_dir(self):
        d = filedialog.askdirectory(title="원본 영상 폴더")
        if d:
            self.videos = [d]
            self._after_pick()

    def _after_pick(self):
        names = [Path(v).name for v in self.videos]
        self.video_var.set(", ".join(names[:4]) + (f" 외 {len(names) - 4}개" if len(names) > 4 else ""))
        base = Path(self.videos[0])
        if not self.out_var.get():
            self.out_var.set(str((base if base.is_dir() else base.parent) / "cgcut_out"))
        if not self.prefix_var.get() and len(self.videos) == 1 and base.is_file():
            stem = base.stem
            self.prefix_var.set(stem.split("_")[0] if "_" in stem else stem)

    def pick_list(self):
        f = filedialog.askopenfilename(title="CG 리스트", filetypes=LIST_TYPES)
        if f:
            self.list_var.set(f)
            if not self.out_var.get():
                self.out_var.set(str(Path(f).parent / "cgcut_out"))

    def pick_out(self):
        d = filedialog.askdirectory(title="저장 폴더")
        if d:
            self.out_var.set(d)

    def open_out(self):
        d = self.out_var.get().strip()
        if not d or not Path(d).exists():
            messagebox.showinfo(APP, "아직 저장 폴더가 없습니다.")
            return
        if sys.platform == "win32":
            os.startfile(d)  # type: ignore[attr-defined]
        elif sys.platform == "darwin":
            subprocess.Popen(["open", d])
        else:
            subprocess.Popen(["xdg-open", d])

    # ------------------------------------------------------------------ run
    def _args(self) -> list[str]:
        argv: list[str] = []
        lst = self.list_var.get().strip()
        if lst:
            argv.append(lst)
        for line in self.tc_text.get("1.0", "end").splitlines():
            if line.strip():
                argv += ["--tc", line.strip()]
        if not lst and not any(a == "--tc" for a in argv):
            raise ValueError("리스트 파일을 고르거나 TC를 입력하세요.")
        if not self.videos:
            raise ValueError("원본 영상을 선택하세요.")
        for v in self.videos:
            argv += ["--video-dir" if Path(v).is_dir() else "--video", v]
        out = self.out_var.get().strip()
        if not out:
            raise ValueError("저장 폴더를 정하세요.")
        argv += ["-o", out]
        codec = dict(CODECS)[self.codec_var.get()]
        argv += ["--codec", codec]
        if self.prefix_var.get().strip():
            argv += ["--prefix", self.prefix_var.get().strip()]
        h = self.handles_var.get().strip() or "0"
        if not h.isdigit():
            raise ValueError("핸들은 숫자(프레임 수)여야 합니다.")
        if int(h):
            argv += ["--handles", h]
        if self.mode_var.get() == "fixed":
            try:
                b, a = float(self.before_var.get()), float(self.after_var.get())
            except ValueError:
                raise ValueError("고정 길이의 앞/뒤 초는 숫자여야 합니다.")
            argv += ["--mode", "fixed", "--before", str(b), "--after", str(a)]
        if self.edl_src_var.get():
            argv.append("--edl-source")
        st = self.start_var.get().strip()
        if st:
            if not cgcut.TC.parse(st):
                raise ValueError("영상 첫 프레임 TC 형식이 잘못됐습니다. 예: 01:00:00:00")
            argv += ["--start-tc", st]
        return argv

    def start(self):
        try:
            argv = self._args()
        except ValueError as e:
            messagebox.showwarning(APP, str(e))
            return
        self._log_clear()
        self.cancel_flag = False
        self.run_btn.configure(state="disabled")
        self.stop_btn.configure(state="normal")
        self.status_var.set("작업 중…")
        cgcut.LOG_HOOK = self.q.put
        self.worker = threading.Thread(target=self._work, args=(argv,), daemon=True)
        self.worker.start()

    def _work(self, argv: list[str]):
        try:
            args = cgcut.build_parser().parse_args(argv)
            rc = cgcut.run(args, cancel=lambda: self.cancel_flag)
            self.q.put(f"__DONE__{rc}")
        except cgcut.CutError as e:
            self.q.put(str(e))
            self.q.put("__DONE__2")
        except SystemExit as e:
            self.q.put("__DONE__2")
        except Exception as e:  # noqa: BLE001
            self.q.put(f"[오류] {type(e).__name__}: {e}")
            self.q.put("__DONE__2")

    def stop(self):
        self.cancel_flag = True
        self.status_var.set("중단 요청… 현재 컷이 끝나면 멈춥니다")

    def _pump(self):
        try:
            while True:
                m = self.q.get_nowait()
                if m.startswith("__DONE__"):
                    rc = int(m[8:])
                    self.run_btn.configure(state="normal")
                    self.stop_btn.configure(state="disabled")
                    self.status_var.set({0: "완료", 3: "중단됨"}.get(rc, "오류 — 아래 내용을 확인하세요"))
                    if rc == 0:
                        self.open_out()
                else:
                    self._log(m)
        except queue.Empty:
            pass
        self.after(100, self._pump)

    def _log(self, s: str):
        self.log.configure(state="normal")
        self.log.insert("end", s + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    def _log_clear(self):
        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.configure(state="disabled")


def main():
    if len(sys.argv) > 1:  # 명령줄 인자가 있으면 CLI로 동작
        sys.exit(cgcut.main())
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
