"""以 python-docx 產生獨立的 Word 差異報告（與網頁使用同一份比對結果）。"""
from docx import Document
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor

FONT = "Microsoft JhengHei"
GREEN = RGBColor(0x1B, 0x7F, 0x3B)
RED = RGBColor(0xC6, 0x28, 0x28)
GRAY = RGBColor(0x55, 0x55, 0x55)
NAVY = RGBColor(0x1F, 0x2A, 0x44)


def _fmt(run, size=None, bold=None, color=None, strike=None, italic=None):
    run.font.name = FONT
    run._element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
    if size:
        run.font.size = Pt(size)
    if bold is not None:
        run.font.bold = bold
    if italic is not None:
        run.font.italic = italic
    if color is not None:
        run.font.color.rgb = color
    if strike is not None:
        run.font.strike = strike


def _add(par, text, **kw):
    run = par.add_run(text)
    _fmt(run, **kw)
    return run


def _heading(doc, text, level):
    h = doc.add_heading(level=level)
    _fmt(h.add_run(text), color=NAVY)
    return h


def _segments(par, segments, color, strike):
    for s in segments:
        _add(par, s["text"], color=color, strike=strike, bold=(s["t"] != "equal"))


def _cell(cell, text, bold=False):
    _add(cell.paragraphs[0], text, bold=bold)


def build_report(result: dict, path) -> None:
    doc = Document()
    normal = doc.styles["Normal"]
    normal.font.name = FONT
    normal.font.size = Pt(11)
    normal.element.rPr.rFonts.set(qn("w:eastAsia"), FONT)

    title = doc.add_heading(level=0)
    _fmt(title.add_run("企業文件新舊版本差異分析報告"), color=NAVY)

    for label, value in (("舊版本文件", result["old_name"]), ("新版本文件", result["new_name"]),
                         ("比對日期時間", result["created_at"])):
        p = doc.add_paragraph()
        _add(p, f"{label}：", bold=True)
        _add(p, value)

    stats = result["stats"]
    _heading(doc, "一、差異統計", 1)
    table = doc.add_table(rows=2, cols=4, style="Table Grid")
    for c, h in enumerate(("新增", "刪除", "修改", "合計")):
        _cell(table.rows[0].cells[c], h, bold=True)
    for c, k in enumerate(("added", "deleted", "modified", "total")):
        _cell(table.rows[1].cells[c], str(stats[k]))

    _heading(doc, "二、差異內容", 1)
    if not result["items"]:
        _add(doc.add_paragraph(), "兩份文件未發現段落差異", bold=True)
    for it in result["items"]:
        _heading(doc, f"#{it['id']}　{it['type_label']}　{it['position']}", 2)
        if it["type"] == "added":
            p = doc.add_paragraph()
            _add(p, "【新增】", bold=True, color=GREEN)
            _segments(p, it["new_segments"], GREEN, False)
        elif it["type"] == "deleted":
            p = doc.add_paragraph()
            _add(p, "【刪除】", bold=True, color=RED, strike=False)
            _segments(p, it["old_segments"], RED, True)
        else:
            p = doc.add_paragraph()
            _add(p, "【修改前】", bold=True, color=RED, strike=False)
            _segments(p, it["old_segments"], RED, True)
            p = doc.add_paragraph()
            _add(p, "【修改後】", bold=True, color=GREEN)
            _segments(p, it["new_segments"], GREEN, False)
        if it["note"]:
            _add(doc.add_paragraph(), "備註：" + it["note"], italic=True, color=GRAY, size=10)

    _heading(doc, "三、Gemini AI 影響分析", 1)
    ai = result["ai"]
    if ai["status"] == "ok":
        d = ai["data"]
        _add(doc.add_paragraph(), "以下內容由 AI 依實際差異產生，其中標示「AI推論」者為推測，僅供參考，請由人員審閱確認。",
             italic=True, color=GRAY, size=10)
        if ai.get("partial_note"):
            _add(doc.add_paragraph(), ai["partial_note"], bold=True, color=RED)
        _heading(doc, "變更摘要", 2)
        _add(doc.add_paragraph(), d["summary"])
        _heading(doc, "可能影響", 2)
        for imp in d["impacts"] or []:
            _add(doc.add_paragraph(style="List Bullet"), f"【{imp['area']}／{imp['basis']}】{imp['description']}")
        if not d["impacts"]:
            _add(doc.add_paragraph(), "模型未列出具體影響。")
        _heading(doc, "風險評估", 2)
        p = doc.add_paragraph()
        _add(p, "風險程度：", bold=True)
        _add(p, d["risk_level"], bold=True)
        p = doc.add_paragraph()
        _add(p, "判斷依據與理由：", bold=True)
        _add(p, d["risk_reasons"])
        _heading(doc, "改善建議", 2)
        for r in d["recommendations"] or ["模型未提供具體建議。"]:
            _add(doc.add_paragraph(style="List Number"), r)
        p = doc.add_paragraph()
        _add(p, "資訊不足之處：", bold=True)
        _add(p, d["information_gaps"])
    elif ai["status"] == "skipped":
        _add(doc.add_paragraph(), ai["message"])
    else:
        _add(doc.add_paragraph(), "AI 影響分析未完成", bold=True, color=RED, size=13)
        _add(doc.add_paragraph(), ai["message"])
        if ai.get("partial_note"):
            _add(doc.add_paragraph(), ai["partial_note"])

    _heading(doc, "四、限制與揭露", 1)
    for text in result["limitations"]:
        _add(doc.add_paragraph(style="List Bullet"), text, size=10)
    _add(doc.add_paragraph(), "差異內容（經敏感資料遮蔽後）可能已傳送至外部 AI 服務（Google Gemini）進行分析。",
         size=10, color=GRAY)
    doc.save(str(path))
