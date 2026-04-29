NAME = "检查目录序号"

def run(doc):
    import re
    try:
        # 更新所有目录域
        for fld in doc.Fields:
            if fld.Type == 37:   # wdFieldTOC
                fld.Update()

        toc_range = None
        for fld in doc.Fields:
            if fld.Type == 37:
                toc_range = fld.Result
                break
        if toc_range is None:
            return False, "未找到自动生成的目录域"

        items = []
        for para in toc_range.Paragraphs:
            text = para.Range.Text.strip()
            if text and not text.isdigit():
                # 去掉尾部页码（制表符+数字）
                text = re.sub(r'\t+\d+$', '', text)
                items.append(text)

        if not items:
            return False, "目录项为空"

        # 中文数字转阿拉伯
        def c2a(ch):
            m = {"一":1,"二":2,"三":3,"四":4,"五":5,"六":6,"七":7,"八":8,"九":9,"十":10}
            if ch in m:
                return m[ch]
            if len(ch)==2 and ch[0]=="十":
                return 10 + m.get(ch[1],0)
            return 0

        expected = 1
        errors = []
        for idx, item in enumerate(items, 1):
            if item.startswith("第") and "章" in item:
                try:
                    num_str = item[1:item.index("章")]
                    num = c2a(num_str)
                    if num != expected:
                        errors.append(f"第{idx}项：期望第{expected}章，实际{item}")
                    else:
                        expected += 1
                except:
                    errors.append(f"第{idx}项：格式错误 {item}")
        if errors:
            return False, "\n".join(errors)
        else:
            return True, "目录序号正确"
    except Exception as e:
        return False, f"检查目录失败：{e}"