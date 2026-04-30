# -*- coding: utf-8 -*-
NAME = "序号检查"

def run(doc):
    import re
    from tkinter import messagebox

    # ---------- 辅助函数 ----------
    def clear_all_comments():
        while doc.Comments.Count > 0:
            doc.Comments(doc.Comments.Count).Delete()

    def add_comment(rng, text):
        try:
            doc.Comments.Add(rng, text)
        except:
            pass

    # ---------- 中文数字转整数（支持 一~十、十一、二十等）----------
    def chinese_to_int(ch_str):
        map = {'一':1,'二':2,'三':3,'四':4,'五':5,'六':6,'七':7,'八':8,'九':9,'十':10}
        if ch_str in map:
            return map[ch_str]
        if ch_str == '十':
            return 10
        if '十' in ch_str:
            parts = ch_str.split('十')
            if parts[0] == '':
                return 10 + (map.get(parts[1], 0))
            else:
                return map.get(parts[0], 1) * 10 + map.get(parts[1], 0)
        return 0

    # ---------- 定义各级编号的正则和解析函数 ----------
    patterns = [
        (1, r'^第([一二三四五六七八九十]+)章', lambda m: chinese_to_int(m.group(1))),
        (2, r'^([一二三四五六七八九十]+)、', lambda m: chinese_to_int(m.group(1))),
        (3, r'^（([一二三四五六七八九十]+)）', lambda m: chinese_to_int(m.group(1))),
        (4, r'^(\d+)\.', lambda m: int(m.group(1))),
        (5, r'^（(\d+)）', lambda m: int(m.group(1))),
    ]

    def parse_heading(text):
        for level, pattern, extractor in patterns:
            m = re.match(pattern, text.strip())
            if m:
                num = extractor(m)
                if num > 0:
                    return level, num
        return None, None

    def check_headings(paragraphs, h1_style):
        """
        paragraphs: 文档的段落集合（已排除页眉页脚和一级标题样式）
        返回错误列表，每个错误为(段落Range, 错误消息)
        """
        errors = []
        stack = []  # 元素为 (level, last_number)

        for para in paragraphs:
            try:
                if h1_style is not None and para.Style.NameLocal == h1_style:
                    continue
            except:
                pass

            text = para.Range.Text.rstrip('\r')
            if not text.strip():
                continue

            level, num = parse_heading(text)
            if level is None:
                continue

            # 栈为空：必须是1级且编号为1
            if not stack:
                if level != 1 or num != 1:
                    errors.append((para.Range, "文档首个序号应为「第一章」"))
                else:
                    stack.append((level, num))
                continue

            prev_level, prev_num = stack[-1]

            # 同级
            if level == prev_level:
                expected = prev_num + 1
                if num != expected:
                    errors.append((para.Range, f"序号错误：{level}级序号应为{expected}"))
                else:
                    stack[-1] = (level, num)
            # 降级（进入子级）
            elif level > prev_level:
                if level != prev_level + 1:
                    errors.append((para.Range, f"层级跳级：不允许从{prev_level}级直接到{level}级"))
                elif num != 1:
                    errors.append((para.Range, f"序号错误：{level}级子标题起始编号应为1"))
                else:
                    stack.append((level, num))
            # 返回上级
            else:  # level < prev_level
                while stack and stack[-1][0] >= level:
                    stack.pop()
                if not stack:
                    errors.append((para.Range, "层级混乱，无上层匹配"))
                    stack.append((level, num))
                else:
                    parent_level, parent_num = stack[-1]
                    if level != parent_level:
                        errors.append((para.Range, f"返回层级不匹配父级，期望{parent_level}级"))
                    else:
                        expected = parent_num + 1
                        if num != expected:
                            errors.append((para.Range, f"序号错误：{level}级序号应为{expected}"))
                        else:
                            stack[-1] = (level, num)

        return errors

    # ---------- 主逻辑 ----------
    try:
        # 获取一级标题样式（用于排除）
        try:
            h1_style = doc.Styles("标题 1").NameLocal
        except:
            h1_style = None

        # 检查文档节数
        sections = doc.Sections
        if sections.Count < 4:
            return False, "文档节数不足4节，无法定位第四节至倒数第二节范围。"

        # 定位正文范围（第四节至倒数第二节）
        start_pos = sections(4).Range.Start
        end_pos = sections(sections.Count - 1).Range.End
        target_range = doc.Range(start_pos, end_pos)

        # 清空旧批注
        clear_all_comments()

        # 收集需要检查的段落（主文档故事，且非一级标题样式）
        paras_to_check = []
        for para in target_range.Paragraphs:
            # 跳过页眉页脚
            if para.Range.StoryType != 1:  # wdMainTextStory
                continue
            if h1_style is not None:
                try:
                    if para.Style.NameLocal == h1_style:
                        continue
                except:
                    pass
            paras_to_check.append(para)

        # 执行序号检查
        heading_errors = check_headings(paras_to_check, h1_style)

        # 添加批注
        for rng, msg in heading_errors:
            add_comment(rng, msg)

        if heading_errors:
            messagebox.showinfo("序号检查", f"发现 {len(heading_errors)} 处序号问题，已添加批注。")
        else:
            messagebox.showinfo("序号检查", "未发现序号错误。")

        return True, f"序号检查完成，发现 {len(heading_errors)} 处问题"

    except Exception as e:
        messagebox.showerror("序号检查错误", str(e))
        return False, f"序号检查失败: {e}"