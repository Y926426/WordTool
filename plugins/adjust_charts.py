NAME = "调整图表"

def run(doc):
    try:
        # ---------- 以下是调整图表的核心逻辑 ----------
        # 1. 处理域代码
        for fld in doc.Fields:
            if "图" in fld.Code.Text or "表" in fld.Code.Text:
                fld.Unlink()

        # 2. 保护“单位：”
        find = doc.Content.Find
        find.ClearFormatting()
        find.Replacement.ClearFormatting()
        find.Text = "单位："
        find.Replacement.Text = "§UNIT§"
        find.Execute(Replace=2)

        # 3. 替换“图数字：” -> “图表：”
        find.MatchWildcards = True
        find.Text = "图[0-9]{1,3}："
        find.Replacement.Text = "图表："
        find.Execute(Replace=2)

        find.Text = "表[0-9]{1,3}："
        find.Execute(Replace=2)

        find.Text = "图[一二三四五六七八九十]{1,3}："
        find.Execute(Replace=2)

        find.Text = "表[一二三四五六七八九十]{1,3}："
        find.Execute(Replace=2)

        find.MatchWildcards = False
        find.Text = "图图表："
        find.Replacement.Text = "图表："
        find.Execute(Replace=2)

        # 4. 恢复“单位：”
        find.Text = "§UNIT§"
        find.Replacement.Text = "单位："
        find.Execute(Replace=2)

        # 5. 给图片应用“图”样式
        try:
            for i in doc.InlineShapes:
                if i.Type == 1:
                    i.Range.Paragraphs(1).Range.Style = "图"
            for s in doc.Shapes:
                if s.Type == 13 and s.Anchor:
                    s.Anchor.Paragraphs(1).Range.Style = "图"
        except:
            pass

        # 6. 设置文本样式
        def set_style(find_text, style_name):
            f = doc.Content.Find
            f.ClearFormatting()
            f.Replacement.ClearFormatting()
            f.Text = find_text
            f.Replacement.Text = find_text
            f.Replacement.Style = style_name
            f.Forward = True
            f.Wrap = 0
            f.Execute(Replace=2)

        set_style("图表：", "图表标题")
        set_style("单位：", "图表单位")
        set_style("来源：", "来源")

        # 7. 处理表格
        for tbl in doc.Tables:
            try:
                tbl.Rows.Alignment = 1   # 居中
            except:
                pass
            if tbl.Rows.Count >= 1:
                for cell in tbl.Rows(1).Cells:
                    cell.Shading.BackgroundPatternColor = 0x1D4B81  # 首行蓝色
            for i in range(2, tbl.Rows.Count + 1):
                for cell in tbl.Rows(i).Cells:
                    cell.Shading.BackgroundPatternColor = 0xFFFFFF  # 内容白色

        return True, "调整图表完成"
    except Exception as e:
        return False, f"调整图表失败：{e}"