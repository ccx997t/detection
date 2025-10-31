#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
巡检报告模板插入表格图片模块（report_embedder.py）
--------------------------------------------------
功能说明：
    本模块用于将 Excel 表格截图（JPG 文件）自动嵌入到 Word 巡检报告模板中，
    按照模板中定义的占位符（如 {{表1}}、{{表2}} ... {{表7}}）依次替换为对应的图片。

核心流程：
    1. 加载 Word 模板文件；
    2. 遍历模板中的段落与表格，查找占位符；
    3. 根据占位符名称加载对应目录下的 JPG 文件；
    4. 在占位符处插入图片（自动居中、宽度固定）；
    5. 保存生成的最终报告文件。

输入输出：
    - 输入：Word 模板文件路径、表格截图目录（images_dir）
    - 输出：生成的完整巡检报告 Word 文件（保存在 output_dir）

依赖模块：
    - python-docx：Word 文档操作
    - util.Logger：自定义日志输出（来自 modules/util.py）
"""

import os
import re
import sys
from typing import Dict
from docxtpl import DocxTemplate, InlineImage      # 导入 docxtpl 模板类与插图类
from docx.shared import Inches                    # 导入单位类，用于设置图片宽度为英寸
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT

# ============================================================
# 修正模块搜索路径，确保可导入 modules 下的工具模块
# ============================================================
PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if PROJECT_ROOT not in sys.path:
    sys.path.append(PROJECT_ROOT)
print(f">> PROJECT_ROOT = {PROJECT_ROOT},  __file__ = {__file__}")

# ============================================================
# 项目模块 util


# ============================================================
# 模块函数定义
# ============================================================




def load_jpg_files(images_dir: str) -> Dict[str, str]:
    """
    加载指定目录下的所有表格截图文件，并构建占位符映射表。
    功能：
        - 自动扫描表1.jpg ~ 表7.jpg；
        - 生成占位符与对应图片路径的字典：
            例如：{"{{表1}}": "/path/to/表1.jpg", ...}
    参数：
        images_dir (str): 图片存放目录路径
    返回：
        Dict[str, str]: 占位符 → 图片路径 的映射表
    """
    image_map = {}
    # 检查图片目录是否存在
    if not os.path.isdir(images_dir):
        print(f"❌ 图片目录不存在：{images_dir}")
        return  # 退出函数

    # 获取所有 .jpg 文件名列表
    image_files = [f for f in os.listdir(images_dir) if f.lower().endswith(".jpg")]
    # 如果未发现任何图片，提示用户
    if not image_files:
        print("⚠️ 未找到任何 .jpg 文件，模板将不进行替换。")

    # 遍历所有图片文件
    for filename in image_files:
        key = os.path.splitext(filename)[0]  # 从文件名中提取变量名（去除扩展名）
        img_path = os.path.join(images_dir, filename)  # 拼接图片的完整路径

        # 再次确认文件存在（保险处理）
        if os.path.exists(img_path):
            image_map[key] =img_path
            print(f"✅ 已准备图片：{img_path} → 模板变量 {{ {key} }}")  # 打印图片加载成功信息
    return image_map





def find_placeholders_and_replace(doc: DocxTemplate, image_map: Dict[str, str]) -> None:
    """
    遍历整个文档（段落与表格单元格），匹配占位符并插入图片。
    逻辑：
        - 优先扫描所有段落；
        - 再扫描表格内的所有单元格；
        - 每当匹配到占位符（如 {{表3}}），则调用 replace_placeholder_with_image()。
    参数：
        DocxTemplate: Word 文档对象
        image_map (Dict[str, str]): 占位符 → 图片路径 映射表
    """

    context = {}  # 初始化上下文字典，用于存放变量名和图片对象
    for key, img_path in image_map.items():
        context[key] = InlineImage(doc, img_path, width=Inches(6.5))  # 设置图片宽度为 6.5 英寸
        print(f"键：{key}，值：{img_path}")
    doc.render(context)
    basename = os.path.basename(template_path)
    new_name = re.sub(r"模板\(.*?\)", "", basename).replace(".docx", "")
    new_name = new_name.strip("-_ ") + ".docx"
    output_path = os.path.join(output_dir, new_name)

    os.makedirs(output_dir, exist_ok=True)
    doc.save(output_path)
    print(f"✅ 生成报告成功：{output_path}")

def save_doc(doc: DocxTemplate, template_path: str, output_dir: str) -> str:
    """
    保存 Word 文档到指定目录。
    功能：
        - 根据模板文件名自动生成输出文件名；
        - 移除文件名中的“模板(版本号)”；
        - 确保输出目录存在；
        - 保存生成的 Word 报告。
    参数：
        doc (Document): Word 文档对象
        template_path (str): 模板文件路径
        output_dir (str): 输出目录
    返回：
        str: 生成报告的完整输出路径
    """
    # 例如：模板文件 “实验性项目巡检报告模板(1.0).docx”
    # → 输出文件 “实验性项目巡检报告.docx”
    basename = os.path.basename(template_path)
    new_name = re.sub(r"模板\(.*?\)", "", basename).replace(".docx", "")
    new_name = new_name.strip("-_ ") + ".docx"
    output_path = os.path.join(output_dir, new_name)

    os.makedirs(output_dir, exist_ok=True)
    doc.save(output_path)
    print(f"✅ 生成报告成功：{output_path}")
    return output_path


def run(template_path: str, images_dir: str, output_dir: str):
    """
    模块主执行函数。
    功能流程：
        1. 加载模板；
        2. 加载表格截图映射表；
        3. 查找占位符并替换为图片；
        4. 保存生成的新报告文件。
    参数：
        template_path (str): Word 模板路径
        images_dir (str): 存放图片的目录
        output_dir (str): 输出文件的保存目录
    """
    doc = DocxTemplate(template_path)  # 加载 Word 模板为 docxtpl 文档对象
    image_map = load_jpg_files(images_dir)
    find_placeholders_and_replace(doc, image_map)
    #save_doc(doc, template_path, output_dir)


# ============================================================
# 测试运行（仅在独立运行时触发）
# ============================================================
if __name__ == "__main__":
    template_path="../template/实验性项目巡检报告模板(1.0).docx"       # Word 模板路径
    images_dir="../tmp/images/"   # 图片目录路径
    output_dir="../out/"    # 最终生成报告的保存路径
    run(template_path, images_dir, output_dir)
