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

    def add_thousands_separator(text):
        """
        在文本中为没有千分符的数字添加千分位逗号，跳过年份（1900-2099的四位数字）。
        """
        # 匹配整数部分（至少一位），可选小数部分
        # 避免匹配到已有逗号的数字（自然跳过）
        pattern = r'(?<![0-9,])(\d+)(?:\.(\d+))?(?![0-9])'
        
        def repl(match):
            integer_part = match.group(1)
            decimal_part = match.group(2)
            num_len = len(integer_part)
            
            # 跳过年份：长度为4且在1900-2099范围
            if num_len == 4:
                num = int(integer_part)
                if 1900 <= num <= 2099:
                    return match.group(0)  # 原样返回
            
            # 长度小于4不添加
            if num_len < 4:
                return match.group(0)
            
            # 否则添加千分符
            # 注意：整数部分可能会很大，使用 int() 转换，但数值可能超出 Python int 范围？不会，Word文档中一般不会出现天文数字
            try:
                formatted_int = '{:,}'.format(int(integer_part))
            except:
                return match.group(0)
            if decimal_part is not None:
                return formatted_int + '.' + decimal_part
            else:
                return formatted_int
        
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

        start_pos = doc.Sections(4).Range.Start
        end_pos = doc.Sections(doc.Sections.Count - 1).Range.End
        target_range = doc.Range(start_pos, end_pos)

        processed_count = 0
        for para in target_range.Paragraphs:
            if not should_process_paragraph(para, h1_style):
                continue
            original_text = para.Range.Text.rstrip('\r')
            if not original_text.strip():
                continue
            new_text = add_thousands_separator(original_text)
            if new_text != original_text:
                # 更新段落文本，保留段落标记
                if para.Range.Text.endswith('\r'):
                    new_text += '\r'
                para.Range.Text = new_text
                processed_count += 1

        return True, f"千分符添加完成，共处理 {processed_count} 个段落"
    except Exception as e:
        return False, f"处理失败: {e}"
