# -*- coding: utf-8 -*-
NAME = "调整图表"

def run(doc):
    import re
    from tkinter import messagebox

    def add_comment(rng, text):
        try:
            doc.Comments.Add(rng, text)
        except:
            pass

    try:
        modified = False

        # 1. 取消图/表域链接（无批注）
        for fld in doc.Fields:
            code = fld.Code.Text
            if "图" in code or "表" in code:
                fld.Unlink()

        # 2. 统一图表标题前缀为“图表：”并应用样式
        pattern = re.compile(r'^(?:图|表)\s*(?:\d+|[一二三四五六七八九十]+)\s*[：:]', re.UNICODE)
        for para in doc.Paragraphs:
            text = para.Range.Text
            m = pattern.match(text)
            if m:
                old_prefix = text[:m.end()]
                new_text = "图表：" + text[m.end():]
                if text.endswith('\r'):
                    new_text = new_text.rstrip('\r') + '\r'
                para.Range.Text = new_text
                modified = True
                add_comment(para.Range, f"已将「{old_prefix}」修改为「图表：」")
                try:
                    if para.Range.Style.NameLocal != "图表标题":
                        para.Range.Style = "图表标题"
                        add_comment(para.Range, "应用样式「图表标题」")
                except:
                    pass

        # 2.1 已经是“图表：”但样式不对的段落
        for para in doc.Paragraphs:
            if para.Range.Text.startswith("图表："):
                try:
                    if para.Range.Style.NameLocal != "图表标题":
                        para.Range.Style = "图表标题"
                        modified = True
                        add_comment(para.Range, "应用样式「图表标题」")
                except:
                    pass

        # 3. 图片样式
        for inline in doc.InlineShapes:
            try:
                para = inline.Range.Paragraphs(1)
                if para.Range.Style.NameLocal != "图":
                    para.Range.Style = "图"
                    modified = True
                    add_comment(para.Range, "为图片所在的段落应用样式「图」")
            except:
                pass
        for shape in doc.Shapes:
            try:
                if shape.Anchor:
                    para = shape.Anchor.Paragraphs(1)
                    if para.Range.Style.NameLocal != "图":
                        para.Range.Style = "图"
                        modified = True
                        add_comment(para.Range, "为浮动图片所在的段落应用样式「图」")
            except:
                pass

        # 4. 图表单位样式
        paras = list(doc.Paragraphs)
        for i, para in enumerate(paras):
            if para.Range.Text.startswith("图表："):
                if i + 1 < len(paras):
                    next_para = paras[i + 1]
                    if next_para.Range.Text.strip().startswith("单位："):
                        if next_para.Range.Style.NameLocal != "图表单位":
                            try:
                                next_para.Range.Style = "图表单位"
                                modified = True
                                add_comment(next_para.Range, "应用样式「图表单位」")
                            except:
                                pass

        # 5. 资料来源/数据来源样式
        for para in doc.Paragraphs:
            text = para.Range.Text.strip()
            if text.startswith("资料来源：") or text.startswith("数据来源："):
                if para.Range.Style.NameLocal != "来源":
                    try:
                        para.Range.Style = "来源"
                        modified = True
                        add_comment(para.Range, "应用样式「来源」")
                    except:
                        pass

        # 6. 表格处理
        for tbl in doc.Tables:
            changes = []
            # 6.1 表格整体水平居中
            try:
                if tbl.Rows.Alignment != 1:
                    tbl.Rows.Alignment = 1
                    changes.append("表格整体水平居中")
            except:
                pass
            # 6.2 单元格内容水平垂直居中
            has_align = False
            for row in tbl.Rows:
                for cell in row.Cells:
                    try:
                        if cell.Range.ParagraphFormat.Alignment != 1:
                            cell.Range.ParagraphFormat.Alignment = 1
                            has_align = True
                    except:
                        pass
                    try:
                        if cell.VerticalAlignment != 1:
                            cell.VerticalAlignment = 1
                            has_align = True
                    except:
                        pass
            if has_align:
                changes.append("单元格内容水平垂直居中")
            # 6.3 首行重复标题
            if tbl.Rows.Count >= 1:
                try:
                    if not tbl.Rows(1).HeadingFormat:
                        tbl.Rows(1).HeadingFormat = True
                        changes.append("首行作为标题行重复")
                except:
                    pass
            # 6.4 应用表格样式
            if tbl.Rows.Count >= 1:
                try:
                    if tbl.Rows(1).Range.Style.NameLocal != "表格（首行）":
                        tbl.Rows(1).Range.Style = "表格（首行）"
                        changes.append("首行应用样式「表格（首行）」")
                except:
                    pass
            for i in range(2, tbl.Rows.Count + 1):
                try:
                    if tbl.Rows(i).Range.Style.NameLocal != "表格（内容）":
                        tbl.Rows(i).Range.Style = "表格（内容）"
                        changes.append(f"第{i}行应用样式「表格（内容）」")
                except:
                    pass
            # 6.5 首行填充色（改进版：逐个单元格尝试，并统计成功数量）
            if tbl.Rows.Count >= 1:
                target_color = 0x814B1D
                success_count = 0
                for cell in tbl.Rows(1).Cells:
                    try:
                        if cell.Shading.BackgroundPatternColor != target_color:
                            cell.Shading.BackgroundPatternColor = target_color
                            success_count += 1
                    except:
                        pass
                if success_count > 0:
                    changes.append(f"首行底纹颜色设置为深蓝色（成功 {success_count}/{len(tbl.Rows(1).Cells)} 个单元格）")
            if changes:
                modified = True
                comment_range = tbl.Rows(1).Cells(1).Range
                comment_text = "表格格式调整：\n• " + "\n• ".join(changes)
                add_comment(comment_range, comment_text)

        # ========== 无修改时的处理 ==========
        if not modified:
            messagebox.showinfo("调整图表", "文档已完美符合规范，没有需要修改的地方。")
        else:
            messagebox.showinfo("调整图表", "修改已完成，请查看文档中的批注确认更改。")

        return True, "调整图表完成"

    except Exception as e:
        messagebox.showerror("调整图表错误", str(e))
        return False, f"调整图表失败: {e}"
