# -*- coding: utf-8 -*-
NAME = "增加千分符"

def run(doc):
    import re

    def should_process_paragraph(para, h1_style):
        """判断段落是否应该处理：主文档故事、非一级标题、且不在表格内"""
        if para.Range.StoryType != 1:  # wdMainTextStory
            return False
        if h1_style is not None:
            try:
                if para.Style.NameLocal == h1_style:
                    return False
            except:
                pass
        # 检查段落是否在表格中
        try:
            if para.Range.Information(12):  # wdInTable = 12
                return False
        except:
            pass
        return True

    def process_paragraph(para):
        """在段落中查找数字并原地添加千分符（修订模式）"""
        text = para.Range.Text.rstrip('\r')
        if not text:
            return 0

        pattern = r'(?<![0-9,])(\d+)(?:\.(\d+))?(?![0-9])'
        matches = list(re.finditer(pattern, text))
        if not matches:
            return 0

        # 从后向前替换，避免位置偏移
        doc_range = para.Range
        replacement_count = 0
        for m in reversed(matches):
            integer_part = m.group(1)
            decimal_part = m.group(2)
            start, end = m.start(), m.end()
            num_len = len(integer_part)

            # 跳过年份：长度为4且在1900-2099范围
            if num_len == 4:
                num = int(integer_part)
                if 1900 <= num <= 2099:
                    continue
            if num_len < 4:
                continue

            # 生成带千分符的数字
            try:
                formatted_int = '{:,}'.format(int(integer_part))
            except:
                continue
            new_digit = formatted_int if decimal_part is None else formatted_int + '.' + decimal_part

            # 精确替换数字（保留格式，在修订模式下会生成修订标记）
            sub_range = doc_range.Duplicate
            sub_range.SetRange(doc_range.Start + start, doc_range.Start + end)
            sub_range.Text = new_digit
            replacement_count += 1

        return replacement_count

    try:
        # 保存原始修订状态
        original_revision_state = doc.TrackRevisions
        # 强制开启修订模式
        doc.TrackRevisions = True

        # 获取一级标题样式
        try:
            h1_style = doc.Styles("标题 1").NameLocal
        except:
            h1_style = None

        # 确认节数
        if doc.Sections.Count < 4:
            return False, "文档节数不足4节，无法定位正文范围（第四节至倒数第二节）。"

        start_pos = doc.Sections(4).Range.Start
        end_pos = doc.Sections(doc.Sections.Count - 1).Range.End
        target_range = doc.Range(start_pos, end_pos)

        modified_paragraphs = 0
        total_replacements = 0

        for para in target_range.Paragraphs:
            if not should_process_paragraph(para, h1_style):
                continue
            cnt = process_paragraph(para)
            if cnt > 0:
                modified_paragraphs += 1
                total_replacements += cnt

        # 恢复原来的修订状态（如果希望保持开启，可以注释下一行）
        doc.TrackRevisions = original_revision_state

        return True, f"千分符添加完成（修订模式），共修改 {modified_paragraphs} 个段落中的 {total_replacements} 个数字。如需确认，请使用Word的“审阅”功能接受或拒绝修订。"

    except Exception as e:
        return False, f"处理失败: {e}"
