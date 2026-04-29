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
        """在段落中查找数字并原地添加千分符（不改变段落格式）"""
        # 获取段落文本（不含末尾回车）
        text = para.Range.Text.rstrip('\r')
        if not text:
            return 0

        # 正则：匹配整数部分（至少一位），可选小数部分，避免匹配已有逗号的数字
        # 同时使用环视确保前后不是数字或逗号
        pattern = r'(?<![0-9,])(\d+)(?:\.(\d+))?(?![0-9])'
        matches = list(re.finditer(pattern, text))
        if not matches:
            return 0

        # 从后向前替换，避免影响后续位置偏移
        replacements = []
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

            # 长度小于4不添加
            if num_len < 4:
                continue

            # 生成带千分符的数字串
            try:
                formatted_int = '{:,}'.format(int(integer_part))
            except:
                continue
            new_digit = formatted_int if decimal_part is None else formatted_int + '.' + decimal_part

            # 记录替换位置和内容（在原文本中的位置）
            replacements.append((start, end, new_digit))

        # 执行替换（基于原文本位置，需转换为 Range 中的字符偏移）
        # 注意：段落范围是连续的，我们可以利用 Find 或直接操作 Range
        # 由于我们已经知道在原文本中的偏移量，我们可以通过 para.Range.Characters 定位
        # 但更稳定的方法是：逐个替换，每次替换后后续偏移变化 -> 从后向前替换
        # 下面使用从后向前的方式，创建每个匹配的 Range 并替换文本
        doc_range = para.Range
        for start, end, new_digit in replacements:
            # 创建匹配文本的 Range（相对于段落起始）
            # 注意：段落文本可能包含段落标记，但我们的 text 不包含，而 Range 包含。
            # 所以字符位置需要小心：para.Range 包含末尾回车，长度比 text 多1（如果末尾有回车）
            # 简单处理：我们基于 para.Range 建立一个子范围
            sub_range = doc_range.Duplicate
            sub_range.SetRange(doc_range.Start + start, doc_range.Start + end)
            # 替换文本
            sub_range.Text = new_digit

        return len(replacements)

    try:
        # 获取一级标题样式名称
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

        processed_count = 0
        modified_paragraphs = 0

        for para in target_range.Paragraphs:
            if not should_process_paragraph(para, h1_style):
                continue
            cnt = process_paragraph(para)
            if cnt > 0:
                modified_paragraphs += 1
                processed_count += cnt

        return True, f"千分符添加完成，共修改 {modified_paragraphs} 个段落中的 {processed_count} 个数字。"
    except Exception as e:
        return False, f"处理失败: {e}"
