#!/usr/bin/env python3
# -*- coding: utf-8 -*-


import os, sys, subprocess, time, socket

import uno

import re

import configparser

from com.sun.star.beans import PropertyValue
import time
import subprocess

def ensure_soffice_service():
    """
    检查 soffice UNO 服务是否在运行，否则自动启动。
    """
    print("🔍 检查 LibreOffice UNO 服务状态...")
    result = subprocess.run(["pgrep", "-f", "soffice.*headless"], capture_output=True, text=True)
    if result.returncode == 0:
        print("✅ 检测到 soffice 服务已在运行。")
        return True

    print("⚠️ 未检测到 soffice 服务，尝试启动中...")
    cmd = [
        "soffice",
        "--headless",
        '--accept=socket,host=localhost,port=2002;urp;',
        "--norestore",
        "--nodefault",
        "--nolockcheck",
    ]
    subprocess.Popen(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    time.sleep(5)  # 等待 UNO 服务完全启动
    result = subprocess.run(["pgrep", "-f", "soffice.*headless"], capture_output=True, text=True)
    if result.returncode == 0:
        print("✅ 已启动 soffice UNO 服务。")
        return True
    else:
        print("❌ soffice 服务启动失败")
        return False
    


def update_docx_fields(template_path: str, output_dir: str):
    """
    通过 LibreOffice UNO 刷新 Word 文档的目录、页码等所有域。
    需先启动：
      soffice --headless --accept="socket,host=localhost,port=2002;urp;" --norestore &
    """
    ensure_soffice_service()
    basename = os.path.basename(template_path)
    #print(f"basename = {basename}")
    # 去掉文件名中的“模板”，构成输出文件名。
    new_name = re.sub(r"模板\(.*?\)", "", basename).replace(".docx", "")
    #print(f"new_name = {new_name}")
    new_name = new_name.strip("-_ ") + ".docx"
    #print(f"new_name = {new_name}")
    # 构成输出文件全路径。
    output_path = os.path.join(output_dir, new_name)
    output_path = os.path.abspath(output_path)
    # 连接到正在运行的 soffice 服务
    local_ctx = uno.getComponentContext()
    resolver = local_ctx.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver", local_ctx
    )
    ctx = resolver.resolve(
        "uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext"
    )
    smgr = ctx.ServiceManager
    desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)

    # 以隐藏方式加载文档
    props = (PropertyValue(Name="Hidden", Value=True),)
    url = uno.systemPathToFileUrl( output_path )
    doc = desktop.loadComponentFromURL(url, "_blank", 0, props)

    # --- 方法A：通过接口刷新（首选） ---
    try:
        # 1) 刷新所有文本域（页码、交叉引用、日期等）
        #    文档实现了 XTextFieldsSupplier 接口
        text_fields = doc.getTextFields()      # XEnumerationAccess
        text_fields.refresh()                  # 刷新所有域

        # 2) 刷新所有“文档索引”（目录、图表目录、表目录等）
        #    文档实现了 XDocumentIndexesSupplier 接口
        indexes = doc.getDocumentIndexes()     # XIndexAccess
        for i in range(indexes.getCount()):
            idx = indexes.getByIndex(i)        # XDocumentIndex
            idx.update()

        refreshed = True
    except Exception:
        refreshed = False

    # --- 方法B：Dispatcher 触发 .uno:UpdateAll（兜底） ---
    if not refreshed:
        try:
            frame = doc.getCurrentController().getFrame()
            dispatcher = smgr.createInstanceWithContext("com.sun.star.frame.DispatchHelper", ctx)
            # UpdateAll 会尝试更新所有域与索引
            dispatcher.executeDispatch(frame, ".uno:UpdateAll", "", 0, tuple())
            # 再明确触发 UpdateFields / UpdateAllIndexes，增强兼容性
            dispatcher.executeDispatch(frame, ".uno:UpdateFields", "", 0, tuple())
            dispatcher.executeDispatch(frame, ".uno:UpdateAllIndexes", "", 0, tuple())
        except Exception as e:
            # 两条路径都失败则抛出
            doc.close(True)
            raise RuntimeError(f"无法刷新目录/域：{e}")

    # 保存并关闭
    doc.store()
    doc.close(True)
    print(f"✅ 已更新目录与页码：{ output_path }")

def run(config: configparser.ConfigParser):
    """ 模块主执行函数。 """
    # 提取配置文件参数项
    global TEMPLATE_PATH, IMAGES_DIR, OUTPUT_DIR
    TEMPLATE_PATH = config.get("Path", "template_path")
    IMAGES_DIR = config.get("Path", "images_dir")
    OUTPUT_DIR = config.get("Path", "output_dir")
    update_docx_fields(TEMPLATE_PATH,OUTPUT_DIR)
if __name__ == "__main__":
    # 检查参数数量
    if len(sys.argv) != 3:
        print("❌ 参数错误：请提供输入模板路径和输出文件路径")
        print("用法：python3 update_word_toc_uno.py TEMPLATE_PATH OUTPUT_PATH")
        sys.exit(1)

    # 从命令行获取参数
    TEMPLATE_PATH = sys.argv[1]
    OUTPUT_DIR = sys.argv[2]

    # 调用主函数
    update_docx_fields(TEMPLATE_PATH, OUTPUT_DIR)


