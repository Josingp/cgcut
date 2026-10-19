"""GUI 스모크 테스트 (CI 윈도우 러너에서 실행): 창을 띄우지 않고 App을 만들어 샘플 영상을 실제로 자른다.

    python tests/gui_smoke.py <작업폴더>
"""
from __future__ import annotations

import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "tests"))

import make_sample  # noqa: E402
import cgcut_gui  # noqa: E402


def main() -> int:
    work = Path(sys.argv[1] if len(sys.argv) > 1 else "gui_smoke")
    work.mkdir(parents=True, exist_ok=True)
    make_sample.make_video(work / "EP01_sample.mov", "01:00:00:00")
    make_sample.make_xlsx(work / "CG_list_sample.xlsx")

    app = cgcut_gui.App()
    app.withdraw()
    app.videos = [str(work / "EP01_sample.mov")]
    app._after_pick()
    assert app.out_var.get().endswith("cgcut_out"), app.out_var.get()
    app.out_var.set(str(work / "out_tc"))
    app.tc_text.insert("1.0", "01:00:02:00\n창문 01:00:08:08 ~ 01:00:09:00\n")
    app.codec_var.set(cgcut_gui.CODECS[2][0])  # h264
    app.handles_var.set("0")
    app.open_out = lambda: None  # 폴더 창 열지 않기
    app.start()
    t = time.time()
    while app.worker and app.worker.is_alive() and time.time() - t < 180:
        app.update()
        time.sleep(0.05)
    for _ in range(20):
        app.update()
        time.sleep(0.05)
    log = app.log.get("1.0", "end")
    print(log)
    assert app.status_var.get() == "완료", app.status_var.get()
    outs = sorted(p.name for p in (work / "out_tc").glob("*.mp4"))
    assert outs == ["EP01_TC_01_01-00-02-00.mp4", "EP01_TC_창문_01-00-08-08.mp4"], outs

    # 리스트 파일(xlsx) 경로
    app.tc_text.delete("1.0", "end")
    app.list_var.set(str(work / "CG_list_sample.xlsx"))
    app.out_var.set(str(work / "out_xlsx"))
    app.start()
    t = time.time()
    while app.worker and app.worker.is_alive() and time.time() - t < 180:
        app.update()
        time.sleep(0.05)
    for _ in range(20):
        app.update()
        time.sleep(0.05)
    assert app.status_var.get() == "완료", app.log.get("1.0", "end")
    n = len(list((work / "out_xlsx").glob("*.mp4")))
    assert n == 8, n
    app.destroy()
    print("GUI smoke OK")
    return 0


if __name__ == "__main__":
    sys.exit(main())
