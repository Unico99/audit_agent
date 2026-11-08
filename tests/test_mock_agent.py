import pandas as pd
from openpyxl import load_workbook

from agent.loop import AuditAgent
from agent.mock import MockClient


def _setup(tmp_path, monkeypatch):
    data, out = tmp_path / "data", tmp_path / "out"
    data.mkdir()
    pd.DataFrame(
        {
            "كود الحساب": ["1", "2"],
            "اسم الحساب": ["نقدية", "رأس مال"],
            "مدين": [105, 0],
            "دائن": [0, 100],
            "السنة السابقة": [90, -90],
        }
    ).to_excel(data / "tb.xlsx", index=False)
    monkeypatch.setenv("AUDIT_DATA_DIR", str(data))
    monkeypatch.setenv("AUDIT_OUTPUT_DIR", str(out))
    return out


def test_full_loop_with_mock_writes_workpaper(tmp_path, monkeypatch):
    out = _setup(tmp_path, monkeypatch)
    agent = AuditAgent(client=MockClient(), approver=lambda n, a: True, log=lambda s: None)
    answer = agent.ask("افحص ملف tb.xlsx")

    assert "تم حفظ ورقة العمل" in answer
    files = list(out.glob("*.xlsx"))
    assert len(files) == 1
    ws = load_workbook(files[0]).active
    texts = " ".join(str(c.value) for row in ws.iter_rows() for c in row if c.value)
    assert "فرق" in texts  # ملاحظة عدم التوازن (فرق 5)


def test_rejection_means_no_file(tmp_path, monkeypatch):
    out = _setup(tmp_path, monkeypatch)
    agent = AuditAgent(client=MockClient(), approver=lambda n, a: False, log=lambda s: None)
    answer = agent.ask("افحص ملف tb.xlsx")

    assert "ما اتكتبتش" in answer
    assert not list(out.glob("*.xlsx"))


def test_second_question_in_same_session(tmp_path, monkeypatch):
    _setup(tmp_path, monkeypatch)
    agent = AuditAgent(client=MockClient(), approver=lambda n, a: True, log=lambda s: None)
    agent.ask("افحص ملف tb.xlsx")
    assert "تم حفظ" in agent.ask("افحص ملف tb.xlsx تاني")
