# -*- coding: utf-8 -*-
NAME = "序号检查"

def run(doc):
    import re
    from tkinter import messagebox

    def clear_all_comments():
        while doc.Comments.Count > 0:
            doc.Comments(doc.Comments.Count).Delete()

    def add_comment(rng, text):
        try:
            doc.Comments.Add(rng, text)
        except:
            pass

    def chinese_to_int(ch_str):
        map_d = {'一':1,'二':2,'三':3,'四':4,'五':5,'六':6,'七':7,'八':8,'九':9,'十':10}
        if ch_str in map_d:
            return map_d[ch_str]
        if ch_str == '十':
            return 10
        if '十' in ch_str:
            parts = ch_str.split('十')
            if parts[0] == '':
                return 10 + map_d.get(parts[1], 0)
            else:
                return map_d.get(parts[0], 1) * 10 + map_d.get(parts[1], 0)
        return 0

    patterns = [
        (1, r'^[#\s]*第([一二三四五六七八九十]+)章', lambda m: chinese_to_int(m.group(1))),
        (2, r'^[#\s]*([一二三四五六七八九十]+)、', lambda m: chinese_to_int(m.group(1))),
        (3, r'^[#\s]*（([一二三四五六七八九十]+)）', lambda m: chinese_to_int(m.group(1))),
        (4, r'^[#\s]*(\d{1,3})\.(?!\d)', lambda m: int(m.group(1))),
        (5, r'^[#\s]*（(\d{1,3})）', lambda m: int(m.group(1))),
    ]

    def parse_heading(text):
        for level, pattern, extractor in patterns:
            m = re.match(pattern, text.strip())
            if m:
                num = extractor(m)
                if num > 0:
                    return level, num
        return None, None

    def check_headings(paragraphs):
        errors = []
        stack = []

        for para in paragraphs:
            text = para.Range.Text.rstrip('\r')
            if not text.strip():
                continue
            level, num = parse_heading(text)
            if level is None:
                continue

            if not stack:
                if level != 1 or num != 1:
                    errors.append((para.Range, "文档首个序号应为「第一章」"))
                else:
                    stack.append((level, num))
                continue

            prev_level, prev_num = stack[-1]

            if level == prev_level:
                expected = prev_num + 1
                if num != expected:
                    errors.append((para.Range, f"序号错误：{level}级序号应为{expected}"))
                else:
                    stack[-1] = (level, num)
            elif level > prev_level:
                if level != prev_level + 1:
                    errors.append((para.Range, f"层级跳级：不允许从{prev_level}级直接到{level}级"))
                elif num != 1:
                    errors.append((para.Range, f"序号错误：{level}级子标题起始编号应为1"))
                else:
                    stack.append((level, num))
            else:  # level < prev_level
                while stack and stack[-1][0] > level:
                    stack.pop()
                if not stack:
                    errors.append((para.Range, "层级混乱，无上层匹配"))
                    stack.append((level, num))
                else:
                    parent_level, parent_num = stack[-1]
                    if parent_level != level:
                        errors.append((para.Range, f"返回层级不匹配父级，期望{parent_level}级"))
                        stack.append((level, num))
                    else:
                        expected = parent_num + 1
                        if num != expected:
                            errors.append((para.Range, f"序号错误：{level}级序号应为{expected}"))
                        else:
                            stack[-1] = (level, num)

        return errors

    try:
        sections = doc.Sections
        if sections.Count < 4:
            return False, "文档节数不足4节，无法定位第四节至倒数第二节范围。"

        start_pos = sections(4).Range.Start
        end_pos = sections(sections.Count - 1).Range.End
        target_range = doc.Range(start_pos, end_pos)

        clear_all_comments()

        paras_to_check = []
        for para in target_range.Paragraphs:
            if para.Range.StoryType != 1:
                continue
            paras_to_check.append(para)

        errors = check_headings(paras_to_check)

        for rng, msg in errors:
            add_comment(rng, msg)

        if errors:
            messagebox.showinfo("序号检查", f"发现 {len(errors)} 处序号问题，已添加批注。")
        else:
            messagebox.showinfo("序号检查", "未发现序号错误。")

        return True, f"序号检查完成，发现 {len(errors)} 处问题"

    except Exception as e:
        messagebox.showerror("序号检查错误", str(e))
        return False, f"序号检查失败: {e}"
