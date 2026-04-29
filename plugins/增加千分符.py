# -*- coding: utf-8 -*-
NAME = "增加千分符"

def run(doc):
    import re

    def should_process_paragraph(para, h1_style):
        """判断段落是否应该处理：在主文档故事中，且不是一级标题"""
        if para.Range.StoryType != 1:  # wdMainTextStory
            return False
        if h1_style is not None:
            try:
                if para.Style.NameLocal == h1_style:
                    return False
            except:
                pass
        return True

    def add_thousands_separator(text):
        """
        在文本中为没有千分符的数字添加千分位逗号。
        匹配规则：整数部分至少4位，允许小数部分。
        已有逗号的数字跳过。
        """
        # 匹配数字：整数部分（可能带负号？忽略），可选小数
        # 使用正则：单词边界或非数字非字母前，捕获数字串（整数部分和小数部分）
        # 复杂一点：匹配整数部分（可能带负号），然后可选 .小数
        # 为了避免匹配到已带逗号的数字，先检查是否包含逗号。
        # 我们匹配的数字串中不能包含逗号。
        # 正则模式：
        # 边界： (?<![0-9,])  确保前面不是数字或逗号（避免部分匹配）
        # 然后匹配数字串：\d+ (?:\.\d+)?  但整数部分可能很长
        # 然后边界：(?![0-9]) 后面不是数字
        # 使用环视确保不匹配已带逗号的数字（因为已带逗号的数字串中会有逗号，我们的模式不匹配逗号，所以自然跳过）
        pattern = r'(?<![0-9,])(\d+)(?:\.(\d+))?(?![0-9])'
        def repl(match):
            integer_part = match.group(1)
            decimal_part = match.group(2)
            # 如果整数部分长度 >= 4 且没有逗号（肯定没有，因为没匹配到逗号），则添加千分符
            if len(integer_part) >= 4:
                # 添加千分符：从右向左每三位加逗号
                formatted_int = '{:,}'.format(int(integer_part))
                if decimal_part is not None:
                    return formatted_int + '.' + decimal_part
                else:
                    return formatted_int
            else:
                return match.group(0)  # 原样返回
        return re.sub(pattern, repl, text)

    try:
        # 获取一级标题样式名称
        try:
            h1_style = doc.Styles("标题 1").NameLocal
        except:
            h1_style = None

        # 确认节数
        if doc.Sections.Count < 4:
            return False, "文档节数不足4节，无法定位正文范围（第四节至倒数第二节）。"

        # 计算正文范围
        start_pos = doc.Sections(4).Range.Start
        end_pos = doc.Sections(doc.Sections.Count - 1).Range.End
        target_range = doc.Range(start_pos, end_pos)

        # 遍历范围内所有段落
        for para in target_range.Paragraphs:
            if not should_process_paragraph(para, h1_style):
                continue
            # 获取段落原始文本（不含段落标记）
            original_text = para.Range.Text.rstrip('\r')
            if not original_text.strip():
                continue
            new_text = add_thousands_separator(original_text)
            if new_text != original_text:
                # 更新段落文本（保留段落标记）
                if para.Range.Text.endswith('\r'):
                    new_text += '\r'
                para.Range.Text = new_text

        return True, "千分符添加完成"

    except Exception as e:
        return False, f"处理失败: {e}"