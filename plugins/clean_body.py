NAME = "AI正文清洗"

def run(doc):
    try:
        sel = doc.Application.Selection
        if sel.Type == 1:   # 没有选中文本
            return False, "未选中任何文本，请先选中要清洗的区域"
        original_start = sel.Start
        original_end = sel.End
        rng = sel.Range.Duplicate

        doc.Application.ScreenUpdating = False

        for i in range(1, rng.Paragraphs.Count + 1):
            para = rng.Paragraphs(1)
            if para.Style.NameLocal == "正文":
                orig = para.Range.Text
                cleaned = orig.replace(" ", "").replace(chr(12288), "").replace(chr(160), "") \
                               .replace(chr(8194), "").replace(chr(8195), "").replace("\t", "")
                if len(orig) > 0:
                    if orig[-1] == '\r':
                        if not cleaned.endswith('\r'):
                            cleaned += '\r'
                    else:
                        while cleaned.endswith('\r'):
                            cleaned = cleaned[:-1]
                    para.Range.Text = cleaned
                    para.Range.Style = "正文"
            if i < rng.Paragraphs.Count:
                rng.SetRange(para.Range.End, original_end)

        doc.Range(original_start, original_end).Select()
        doc.Application.ScreenUpdating = True
        return True, "正文清洗完成"
    except Exception as e:
        return False, f"正文清洗失败：{e}"