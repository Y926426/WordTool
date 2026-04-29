# -*- coding: utf-8 -*-
NAME = "1.调整图表"

def run(doc):
    import re

    try:
        # 1. 取消图/表域链接
        for fld in doc.Fields:
            code = fld.Code.Text
            if "图" in code or "表" in code:
                fld.Unlink()

        # 2. 统一图表标题前缀为“图表：”
        pattern = re.compile(r'^(?:图|表)\s*(?:\d+|[一二三四五六七八九十]+)\s*[：:]', re.UNICODE)
        for para in doc.Paragraphs:
            text = para.Range.Text
            m = pattern.match(text)
            if m:
                new_text = "图表：" + text[m.end():]
                if text.endswith('\r'):
                    new_text = new_text.rstrip('\r') + '\r'
                para.Range.Text = new_text

        # 3. 应用图表标题样式
        for para in doc.Paragraphs:
            if para.Range.Text.startswith("图表："):
                try:
                    para.Range.Style = "图表标题"
                except:
                    pass

        # 4. 图片样式：给所有包含图片的段落应用“图”样式
        for inline in doc.InlineShapes:
            try:
                para = inline.Range.Paragraphs(1)
                para.Range.Style = "图"
            except:
                pass
        for shape in doc.Shapes:
            try:
                if shape.Anchor:
                    para = shape.Anchor.Paragraphs(1)
                    para.Range.Style = "图"
            except:
                pass

        # 5. 图表单位样式（紧跟图表标题的单位行）
        paras = list(doc.Paragraphs)
        for i, para in enumerate(paras):
            if para.Range.Text.startswith("图表："):
                if i + 1 < len(paras):
                    next_para = paras[i + 1]
                    if next_para.Range.Text.strip().startswith("单位："):
                        try:
                            next_para.Range.Style = "图表单位"
                        except:
                            pass

        # 6. 资料来源/数据来源样式
        for para in doc.Paragraphs:
            text = para.Range.Text.strip()
            if text.startswith("资料来源：") or text.startswith("数据来源："):
                try:
                    para.Range.Style = "来源"
                except:
                    pass

        # 7. 表格处理
        for tbl in doc.Tables:
            # 7.1 表格整体水平居中
            try:
                tbl.Rows.Alignment = 1
            except:
                pass
            # 7.2 单元格内容水平垂直居中
            for row in tbl.Rows:
                try:
                    row.Range.ParagraphFormat.Alignment = 1   # 水平居中
                except:
                    pass
                try:
                    row.Range.Cells.VerticalAlignment = 1     # 垂直居中
                except:
                    pass
            # 7.3 设置首行为标题行（跨页重复）
            if tbl.Rows.Count >= 1:
                try:
                    tbl.Rows(1).HeadingFormat = True
                except:
                    pass
            # 7.4 应用表格样式
            if tbl.Rows.Count >= 1:
                try:
                    tbl.Rows(1).Range.Style = "表格（首行）"
                except:
                    pass
            for i in range(2, tbl.Rows.Count + 1):
                try:
                    tbl.Rows(i).Range.Style = "表格（内容）"
                except:
                    pass
            # 7.5 设置首行填充色（深蓝色 #1D4B81，WPS 需用 BGR 0x814B1D）
            if tbl.Rows.Count >= 1:
                target_color = 0x814B1D   # WPS 兼容值
                for cell in tbl.Rows(1).Cells:
                    cell.Shading.BackgroundPatternColor = target_color

        return True, "调整图表完成"

    except Exception as e:
        return False, f"调整图表失败: {e}"
