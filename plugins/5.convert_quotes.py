# -*- coding: utf-8 -*-
NAME = "双引号修改"

def run(doc):
    import re

    def convert_quotes_in_range(rng):
        """在指定 Range 内执行引号转换"""
        # 获取范围中的文本（包含段落标记）
        text = rng.Text
        # 保留段落标记的长度（文本末尾可能有 \r）
        original_end = rng.End
        # 执行三步替换
        # 1. 成对英文引号 -> 中文引号
        text = re.sub(r'"(.*?)"', r'“\1”', text)
        # 2. 删除所有残余英文引号
        text = text.replace('"', '')
        # 3. 合并连续中文引号
        text = text.replace('““', '“').replace('””', '”')
        # 更新 Range 文本，注意保留原范围（段落标记可能变化）
        rng.Text = text
        return True

    try:
        # 仅处理主文档故事（排除页眉页脚、文本框等）
        main_story = 1  # wdMainTextStory
        # 由于 Word 的 StoryRanges 可能不连续，我们遍历所有段落，只处理主故事
        processed = 0
        for para in doc.Paragraphs:
            if para.Range.StoryType == main_story:
                # 对每个段落独立处理，避免跨段落匹配（但会影响段落内引号替换，已足够）
                # 更严谨的做法是处理整个主故事 Range，但考虑到性能，段落级已可。
                # 注意：若引号跨段落（很少见），则无法正确处理，但可接受。
                rng = para.Range
                convert_quotes_in_range(rng)
                processed += 1
        return True, f"双引号修改完成，共处理 {processed} 个段落。"
    except Exception as e:
        return False, f"双引号修改失败: {e}"