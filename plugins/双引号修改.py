# -*- coding: utf-8 -*-
NAME = "双引号修改"

def run(doc):
    """
    将文档中的英文双引号对替换为中文双引号对，并清理多余引号。
    步骤：
    1. 通配符查找 "(*)" 替换为 “\1” （直接生成正确的中文引号对，无需二次删除英文引号）
       更准确的做法：查找 "([!"]*)" 或使用 "*" 然后删除英文引号。这里采用经典两步法。
    但为了兼容WPS/Word的通配符行为，采用原文描述：
       a) 通配符查找 "*" 替换为 “^&”
       b) 删除所有英文引号 "
       c) 修复连续中文引号
    """
    import re

    try:
        # 步骤1：取消“直引号替换为弯引号”的自动格式？这属于用户选项，不强制修改，因为查找替换可以独立工作。
        # 我们直接进行查找替换。

        # 获取查找对象
        find = doc.Content.Find
        find.ClearFormatting()
        find.Replacement.ClearFormatting()

        # ---------- 第一步：通配符替换，将 "内容" 变为 “"内容"” ----------
        find.Text = "\"*\""
        find.Replacement.Text = "“^&”"
        find.MatchWildcards = True
        find.Forward = True
        find.Wrap = 0  # wdFindStop
        # 执行全部替换
        find.Execute(Replace=2)  # wdReplaceAll = 2

        # ---------- 第二步：删除所有英文双引号 ----------
        find.MatchWildcards = False
        find.Text = "\""
        find.Replacement.Text = ""
        find.Execute(Replace=2)

        # ---------- 第三步：修复可能出现的中文连续双引号，如 ““内容”” 或 “内容”” 等 ----------
        # 处理两个连续左引号
        find.Text = "““"
        find.Replacement.Text = "“"
        find.Execute(Replace=2)

        # 处理两个连续右引号
        find.Text = "””"
        find.Replacement.Text = "”"
        find.Execute(Replace=2)

        # 可选：处理左引号后紧跟右引号的情况（空内容），如 “” -> 空，但不常见
        find.Text = "“”"
        find.Replacement.Text = ""
        find.Execute(Replace=2)

        return True, "双引号修改完成"
    except Exception as e:
        return False, f"修改失败: {e}"