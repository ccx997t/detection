#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
fill_statistic_to_word.py
---------------------------------------
功能：
    1. 读取 Excel 执行 statistic.py 中的巡检统计；
    2. 将统计结果（result）写入 Word 模板 {{汇总结果}} 段落；
    3. 输出生成可打印 Word 报告。
依赖：
    pip install docxtpl pandas openpyxl
"""

import os
from docxtpl import DocxTemplate
from statistic import get_excel_sheets, scan_excel_sheets  # 直接复用你的函数
import io
import sys
import re
import configparser
from jinja2 import Environment, DebugUndefined
# 全局参数
TEMPLATE_PATH = ""
INPUT_DIR = ""
OUTPUT_DIR = ""

def run_statistic_to_word():
    basename = os.path.basename(TEMPLATE_PATH)
    #print(f"basename = {basename}")
    # 去掉文件名中的“模板”，构成输出文件名。
    new_name = re.sub(r"模板\(.*?\)", "", basename).replace(".docx", "")
    #print(f"new_name = {new_name}")
    new_name = new_name.strip("-_ ") + ".docx"
    #print(f"new_name = {new_name}")
    # 构成输出文件全路径。
    output_path = os.path.join(OUTPUT_DIR, new_name)
    """执行统计并将结果写入 Word 模板"""
    print("📊 开始分析 Excel 巡检表...")
    sheet_names = get_excel_sheets(INPUT_DIR)

    # ✅ 获取返回值：汇总字符串 + 结构化结果列表
    summary_text = scan_excel_sheets(INPUT_DIR, sheet_names)

    # 清理日志格式
    summary_text = summary_text.replace("\r", "").strip()
    print(f"\n✅ 汇总结果提取完成（{len(summary_text)} 字）")


    # ✅ 写入 Word 模板
    print(f"\n✅ 读取word报告：{output_path}")
    jinja_env = Environment(undefined=DebugUndefined)
    doc = DocxTemplate(output_path)
    context = {"汇总结果": summary_text}
    doc.render(context,jinja_env=jinja_env)
    doc.save(output_path)
    print(f"\n✅ 已生成报告：{output_path}")


def run(config: configparser.ConfigParser):
    """ 模块主执行函数。 """
    # 提取配置文件参数项
    global TEMPLATE_PATH, INPUT_DIR, OUTPUT_DIR
    TEMPLATE_PATH = config.get("Path", "template_path")
    INPUT_DIR = config.get("Path", "input_path")
    OUTPUT_DIR = config.get("Path", "output_dir")
    run_statistic_to_word()
if __name__ == "__main__":
    TEMPLATE_PATH = "../template/实验性项目巡检报告模板(1.0).docx"
    INPUT_DIR = "../data/巡检报告数据集(1.0).xlsx"
    OUTPUT_DIR = "../out/"
    run_statistic_to_word()
