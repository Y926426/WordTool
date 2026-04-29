# -*- coding: utf-8 -*-
NAME = "检查空格空行"

def run(doc):
    import re

    def clear_all_comments():
        while doc.Comments.Count > 0:
            doc.Comments(doc.Comments.Count).Delete()

    def add_comment(rng, text):
        try:
            doc.Comments.Add(rng, text)
        except:
            pass

    def check_empty_lines(rng, h1_style):
        find_rng = rng.Duplicate
        end_pos = rng.End
        issue_count = 0
        find = find_rng.Find
        find.ClearFormatting()
        find.Text = "^p^p"
        find.Forward = True
        find.Wrap = 0
        while find.Execute() and find_rng.End <= end_pos:
            try:
                para = find_rng.Paragraphs(1)
                if h1_style and para.Style.NameLocal == "标题 1":
                    find_rng.Collapse(0)
                    find_rng.MoveStart(1, 1)
                    find_rng.MoveEnd(1, end_pos - find_rng.Start)
                    continue
            except:
                pass
            add_comment(find_rng, "发现空行")   # 只提示“发现空行”，不计数
            issue_count += 1
            find_rng.Collapse(0)
            find_rng.MoveStart(1, 2)  # 跳过已检测的两个段落标记
            find_rng.MoveEnd(1, end_pos - find_rng.Start)
        return issue_count

    def check_spaces(rng, h1_style):
        find_rng = rng.Duplicate
        end_pos = rng.End
        issue_count = 0
        find = find_rng.Find
        find.ClearFormatting()
        find.Text = " "
        find.Forward = True
        find.Wrap = 0
        while find.Execute() and find_rng.End <= end_pos:
            try:
                para = find_rng.Paragraphs(1)
                if h1_style and para.Style.NameLocal == "标题 1":
                    find_rng.Collapse(0)
                    find_rng.MoveStart(1, 1)
                    find_rng.MoveEnd(1, end_pos - find_rng.Start)
                    continue
            except:
                pass
            add_comment(find_rng, "发现空格")   # 只提示“发现空格”
            issue_count += 1
            find_rng.Collapse(0)
            find_rng.MoveStart(1, 1)
            find_rng.MoveEnd(1, end_pos - find_rng.Start)
        return issue_count

    try:
        try:
            h1_style = doc.Styles("标题 1")
        except:
            h1_style = None

        sections = doc.Sections
        if sections.Count < 4:
            return False, "文档节数不足4节，无法定位第四节至倒数第二节范围。"

        start_pos = sections(4).Range.Start
        end_pos = sections(sections.Count - 1).Range.End
        target_range = doc.Range(start_pos, end_pos)

        clear_all_comments()
        empty_count = check_empty_lines(target_range, h1_style)
        space_count = check_spaces(target_range, h1_style)

        if empty_count + space_count == 0:
            return True, "未发现空行或空格问题"
        else:
            return True, f"检测完成：发现 {empty_count} 处空行，{space_count} 处空格。已添加批注。"

    except Exception as e:
        return False, f"检查失败: {e}"