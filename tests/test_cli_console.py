import io
import os
import subprocess
import sys
from pathlib import Path

from pptx import Presentation
from pptx.util import Inches

from broadcast_ppt import cli

ENV = {k: v for k, v in os.environ.items() if k != "GEMINI_API_KEY"}
WARNING = {"counts": {"critical": 0, "warning": 1, "info": 0},
           "findings": [{"level": "warning", "out_slide": 3, "check": "lines_fit", "message": "x"}]}


def _deck(path: Path) -> Path:
    prs = Presentation()
    s = prs.slides.add_slide(prs.slide_layouts[6])
    s.shapes.add_textbox(Inches(1), Inches(1), Inches(8), Inches(1)).text_frame.text = "1 태초에 말씀이 계시니라 이 말씀이 하나님과 함께 계셨으니"
    path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(path))
    return path


def _plan(tmp_path: Path, monkeypatch) -> Path:
    monkeypatch.delenv("GEMINI_API_KEY", raising=False)
    out = tmp_path / "p" / "plan.json"
    assert cli.main(["plan", str(_deck(tmp_path / "src.pptx")), "--out", str(out)]) == 0
    return out


def test_plan_survives_a_cp949_pipe_and_a_trailing_space_name(tmp_path):
    # Korean Windows with redirected output encodes stdout as cp949; '—' is not in cp949, and NTFS drops a trailing space from a folder name
    src = _deck(tmp_path / "입력" / "주일 테스트 — 000000 .pptx")
    work = tmp_path / "work"
    r = subprocess.run([sys.executable, "-m", "broadcast_ppt.cli", "plan", str(src), "--work", str(work)],
                       env={**ENV, "PYTHONIOENCODING": "cp949"}, capture_output=True)
    assert r.returncode == 0 and b"Traceback" not in r.stderr, r.stderr.decode("utf-8", "replace")
    plans = list(work.glob("*/plan.json"))
    assert len(plans) == 1 and plans[0].parent.name == src.stem.rstrip(" .")


def test_build_without_template_is_a_message(tmp_path, monkeypatch, capsys):
    plan = _plan(tmp_path, monkeypatch)
    assert cli.main(["build", str(plan), "--template", str(tmp_path / "없음.pptx")]) == 2
    assert "템플릿" in capsys.readouterr().err


def test_build_into_a_locked_deck_is_a_message(tmp_path, monkeypatch, capsys):
    plan = _plan(tmp_path, monkeypatch)
    out = plan.with_name("broadcast.pptx")

    def locked(*args, **kwargs):
        raise PermissionError(13, "Permission denied", str(out))

    monkeypatch.setattr("broadcast_ppt.builder.build.build_deck", locked)
    assert cli.main(["build", str(plan), "--template", str(plan)]) == 1
    assert "PowerPoint" in capsys.readouterr().err


def test_report_validation_survives_a_cp949_stdout(monkeypatch):
    buf = io.TextIOWrapper(io.BytesIO(), encoding="cp949")
    monkeypatch.setattr(sys, "stdout", buf)
    cli._utf8_console()
    cli._report_validation(WARNING)
    buf.flush()
    assert "lines_fit — x" in buf.buffer.getvalue().decode("utf-8")


def test_overrides_saved_as_ansi_is_a_message(tmp_path, monkeypatch, capsys):
    monkeypatch.delenv("GEMINI_API_KEY", raising=False)
    f = tmp_path / "overrides.yaml"
    f.write_bytes('slides:\n  1: {ref: "요 1:1"}\n'.encode("cp949"))
    assert cli.main(["plan", str(_deck(tmp_path / "src.pptx")), "--overrides", str(f), "--out", str(tmp_path / "p.json")]) == 2
    assert "UTF-8" in capsys.readouterr().err


def test_plan_of_a_non_pptx_file_is_a_message(tmp_path, monkeypatch, capsys):
    monkeypatch.delenv("GEMINI_API_KEY", raising=False)
    f = tmp_path / "._주일.pptx"  # the AppleDouble sidecar a Mac leaves on a USB stick
    f.write_bytes(b"\x00\x05\x16\x07 not a zip")
    assert cli.main(["plan", str(f), "--out", str(tmp_path / "p.json")]) == 2
    assert ".pptx로 읽을 수 없습니다" in capsys.readouterr().err
