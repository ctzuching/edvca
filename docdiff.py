"""Word 文件讀取與段落差異比對（純 Python，不呼叫任何外部服務）。"""
import io
import re
import zipfile
from difflib import SequenceMatcher

from docx import Document

PAIR_THRESHOLD = 0.5          # 判定為「修改」所需的最低文字相似度
MAX_PAIR_CELLS = 40_000       # 單一差異區塊的配對運算上限（舊段落數 x 新段落數）
MAX_INLINE_DIFF_CHARS = 4000  # 超過此長度的段落不做文字層級標示
MAX_UNCOMPRESSED_BYTES = 200 * 1024 * 1024  # 防止壓縮炸彈

TYPE_LABELS = {"added": "新增", "deleted": "刪除", "modified": "修改"}

STATIC_LIMITATIONS = [
    "比對單位為 .docx 的文字段落，只計入非空白段落；「第 N 段」為非空白段落的順序編號，不等同 Word 內顯示的頁碼或條號。",
    "表格、圖片、文字方塊、註解、追蹤修訂、頁首頁尾、腳註／章節附註等特殊元素未納入比對；其內容的變更不會出現在本報告中。",
    "段落移動無法可靠辨識：位置改變的段落會顯示為「刪除」加「新增」（若內容完全相同，會加註「疑似段落移動」）。",
    "「修改」以文字相似度（門檻 50%）判定；相似度低的大幅改寫會顯示為「刪除」加「新增」。",
    "字型、粗細、顏色、樣式等格式變更不會被偵測。",
]


class DocumentError(Exception):
    """訊息可直接顯示給使用者的錯誤。"""


def normalize(text: str) -> str:
    return re.sub(r"\s+", " ", text).strip()


def load_docx(data: bytes, label: str) -> dict:
    """讀取 .docx 位元組，回傳 {'paragraphs': [{'pos','text'}], 'features': {...}}。"""
    if not data:
        raise DocumentError(f"{label}：檔案內容為空，請重新選擇檔案。")
    if not data.startswith(b"PK"):
        raise DocumentError(f"{label}：檔案不是有效的 .docx 文件（內容格式不符），請確認檔案未被改名或損壞。")
    try:
        with zipfile.ZipFile(io.BytesIO(data)) as zf:
            names = set(zf.namelist())
            if "word/document.xml" not in names:
                raise DocumentError(f"{label}：檔案不是有效的 Word 文件（缺少主文件內容）。")
            if sum(i.file_size for i in zf.infolist()) > MAX_UNCOMPRESSED_BYTES:
                raise DocumentError(f"{label}：檔案解壓縮後過大，無法處理。")
    except zipfile.BadZipFile:
        raise DocumentError(f"{label}：檔案已損壞或不是有效的 .docx 文件。")

    try:
        doc = Document(io.BytesIO(data))
        paragraphs = []
        for p in doc.paragraphs:
            text = normalize(p.text)
            if text:
                paragraphs.append({"pos": len(paragraphs) + 1, "text": text})
        features = _detect_features(doc, names)
    except DocumentError:
        raise
    except Exception:
        raise DocumentError(f"{label}：無法解析此 Word 文件，檔案可能已損壞或格式不受支援。")

    if not paragraphs:
        raise DocumentError(f"{label}：文件中沒有可比對的文字段落（可能是空文件，或內容僅有表格、圖片）。")
    return {"paragraphs": paragraphs, "features": features}


def _detect_features(doc, zip_names) -> dict:
    feats = {"tables": len(doc.tables), "images": 0, "tracked_changes": 0,
             "comments": "word/comments.xml" in zip_names, "header_footer": False}
    try:
        feats["images"] = len(doc.inline_shapes)
    except Exception:
        pass
    try:
        feats["tracked_changes"] = len(doc.element.body.xpath(".//w:ins | .//w:del"))
    except Exception:
        pass
    try:
        for s in doc.sections:
            for part in (s.header, s.footer):
                if not part.is_linked_to_previous and any(p.text.strip() for p in part.paragraphs):
                    feats["header_footer"] = True
    except Exception:
        pass
    return feats


def feature_notes(label: str, f: dict) -> list:
    notes = []
    if f["tables"]:
        notes.append(f"{label}含 {f['tables']} 個表格，表格內容未納入比對。")
    if f["images"]:
        notes.append(f"{label}含 {f['images']} 個內嵌圖片，圖片未納入比對。")
    if f["tracked_changes"]:
        notes.append(f"{label}含追蹤修訂（{f['tracked_changes']} 處），修訂中的插入與刪除文字不會被讀取。")
    if f["comments"]:
        notes.append(f"{label}含註解，註解內容未納入比對。")
    if f["header_footer"]:
        notes.append(f"{label}含頁首或頁尾文字，未納入比對。")
    return notes


# ---------------------------------------------------------------- 比對 ----
def _similarity(a: str, b: str) -> float:
    sm = SequenceMatcher(None, a, b, autojunk=False)
    if sm.real_quick_ratio() < PAIR_THRESHOLD or sm.quick_ratio() < PAIR_THRESHOLD:
        return 0.0
    return sm.ratio()


def _tokens(text: str) -> list:
    # 英數字以「字詞」為單位，中文等其他字元逐字比對
    return re.findall(r"[A-Za-z0-9_]+|\s+|.", text, re.S)


def _inline_diff(old: str, new: str):
    """回傳 (舊版片段, 新版片段)；過長則回傳 None。"""
    if len(old) > MAX_INLINE_DIFF_CHARS or len(new) > MAX_INLINE_DIFF_CHARS:
        return None
    ta, tb = _tokens(old), _tokens(new)
    old_seg, new_seg = [], []
    for tag, i1, i2, j1, j2 in SequenceMatcher(None, ta, tb, autojunk=False).get_opcodes():
        if tag == "equal":
            s = "".join(ta[i1:i2])
            old_seg.append({"t": "equal", "text": s})
            new_seg.append({"t": "equal", "text": s})
            continue
        if i2 > i1:
            old_seg.append({"t": "del", "text": "".join(ta[i1:i2])})
        if j2 > j1:
            new_seg.append({"t": "ins", "text": "".join(tb[j1:j2])})
    return old_seg, new_seg


def _position(o, n) -> str:
    if o and n:
        return f"舊版第 {o} 段／新版第 {n} 段"
    return f"新版第 {n} 段" if n else f"舊版第 {o} 段"


def _make(kind: str, old=None, new=None) -> dict:
    item = {
        "type": kind, "type_label": TYPE_LABELS[kind],
        "old_pos": old["pos"] if old else None, "new_pos": new["pos"] if new else None,
        "old_text": old["text"] if old else "", "new_text": new["text"] if new else "",
        "old_segments": [], "new_segments": [], "note": "",
    }
    item["position"] = _position(item["old_pos"], item["new_pos"])
    if kind == "deleted":
        item["old_segments"] = [{"t": "equal", "text": item["old_text"]}]
    elif kind == "added":
        item["new_segments"] = [{"t": "equal", "text": item["new_text"]}]
    else:
        result = _inline_diff(item["old_text"], item["new_text"])
        if result is None:
            item["old_segments"] = [{"t": "equal", "text": item["old_text"]}]
            item["new_segments"] = [{"t": "equal", "text": item["new_text"]}]
            item["note"] = "段落過長，未進行文字層級標示，請對照修改前後完整內容。"
        else:
            item["old_segments"], item["new_segments"] = result
    return item


def _align_block(olds: list, news: list):
    """在一個「替換」區塊內，依順序以相似度做最佳配對（動態規劃）。回傳 (items, 是否因過大而略過配對)。"""
    m, n = len(olds), len(news)
    if m == 0 or n == 0 or m * n > MAX_PAIR_CELLS:
        items = [_make("deleted", old=o) for o in olds] + [_make("added", new=x) for x in news]
        return items, (m * n > MAX_PAIR_CELLS)
    sim = [[_similarity(o["text"], x["text"]) for x in news] for o in olds]
    dp = [[0.0] * (n + 1) for _ in range(m + 1)]
    for i in range(1, m + 1):
        for j in range(1, n + 1):
            best = max(dp[i - 1][j], dp[i][j - 1])
            if sim[i - 1][j - 1] >= PAIR_THRESHOLD:
                best = max(best, dp[i - 1][j - 1] + sim[i - 1][j - 1])
            dp[i][j] = best
    ops, i, j = [], m, n
    while i > 0 or j > 0:
        if i > 0 and j > 0 and sim[i - 1][j - 1] >= PAIR_THRESHOLD \
                and dp[i][j] == dp[i - 1][j - 1] + sim[i - 1][j - 1]:
            ops.append(("pair", i - 1, j - 1)); i -= 1; j -= 1
        elif j > 0 and (i == 0 or dp[i][j] == dp[i][j - 1]):
            ops.append(("ins", None, j - 1)); j -= 1
        else:
            ops.append(("del", i - 1, None)); i -= 1
    ops.reverse()
    items = []
    for op, a, b in ops:
        if op == "pair":
            items.append(_make("modified", old=olds[a], new=news[b]))
        elif op == "del":
            items.append(_make("deleted", old=olds[a]))
        else:
            items.append(_make("added", new=news[b]))
    return items, False


def compare(old_paras: list, new_paras: list) -> dict:
    """比對新舊段落。回傳 {'items': [...], 'stats': {...}, 'notes': [...]}。"""
    a = [p["text"] for p in old_paras]
    b = [p["text"] for p in new_paras]
    items, notes, skipped = [], [], False
    for tag, i1, i2, j1, j2 in SequenceMatcher(None, a, b, autojunk=False).get_opcodes():
        if tag == "equal":
            continue
        if tag == "delete":
            items += [_make("deleted", old=o) for o in old_paras[i1:i2]]
        elif tag == "insert":
            items += [_make("added", new=x) for x in new_paras[j1:j2]]
        else:
            block, too_big = _align_block(old_paras[i1:i2], new_paras[j1:j2])
            items += block
            skipped = skipped or too_big
    if skipped:
        notes.append("部分差異區塊過大，為控制運算量未進行「修改」配對，這些內容以「刪除」加「新增」呈現。")

    # 完全相同內容同時被刪除與新增 → 標註疑似移動（不改變類型）
    deleted = {}
    for it in items:
        if it["type"] == "deleted":
            deleted.setdefault(it["old_text"], []).append(it)
    for it in items:
        if it["type"] == "added" and deleted.get(it["new_text"]):
            match = deleted[it["new_text"]].pop(0)
            match["note"] = f"此段落內容與新增項目（{it['position']}）完全相同，疑似段落移動。"
            it["note"] = f"此段落內容與刪除項目（{match['position']}）完全相同，疑似段落移動。"

    for k, it in enumerate(items, 1):
        it["id"] = k
    stats = {
        "added": sum(i["type"] == "added" for i in items),
        "deleted": sum(i["type"] == "deleted" for i in items),
        "modified": sum(i["type"] == "modified" for i in items),
    }
    stats["total"] = sum(stats.values())
    return {"items": items, "stats": stats, "notes": notes}
