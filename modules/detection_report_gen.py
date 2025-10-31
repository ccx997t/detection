# 程序名：detection_report_gen.py
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
巡检报告生成器主程序
--------------------------------
功能流程：
1. 读取配置文件（config/config.ini）
2. 解析 Excel 数据
3. 将每个表格转换为 JPG 图像
4. 可选：对图像进行裁剪或美化（由模块内部处理）
5. 将生成的图片嵌入 Word 模板
6. 输出最终报告文件到指定目录

模块结构：
- modules/util.py               → 日志与辅助函数（Logger）
- modules/excel_to_images.py    → Excel 转 JPG 模块
- modules/report_embedder.py    → Word 模板插入模块
"""

import os                                  # 提供文件和路径操作函数
import sys                                 # 提供系统级访问，如路径与退出
import configparser                        # 配置解释器。
import subprocess
# ========== 修正项目模块搜索路径 ==========
# 本文件位于 detection/modules/ 或 detection 根目录下
# PROJECT_ROOT 指向项目的根目录，以便导入 modules 下的自定义模块
PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if PROJECT_ROOT not in sys.path:
    sys.path.append(PROJECT_ROOT)
print(f">> PROJECT_ROOT = {PROJECT_ROOT},  __file__ = {__file__}")

# ============================================================
# 导入项目模块（采用容错方式，防止模块缺失时程序崩溃）
# ============================================================

# 工具模块（日志、配置打印等）
try:
    from modules import util as _ut
except Exception as e:
    _ut = None
    print(f"⚠️  未找到 util 模块：{e}")

# Excel 转 JPG 模块：用于将表格转化为高精度截图
try:
    from modules import excel_to_images as _excel_to_images
except Exception as e:
    _excel_to_images = None
    print(f"⚠️  未找到 excel_to_images 模块：{e}")

# Word 模板嵌入模块：将生成的图片插入到报告模板
try:
    from modules import report_embedder as _report_embedder
except Exception as e:
    _report_embedder = None
    print(f"⚠️  未找到 report_embedder 模块：{e}")

# 插入统计数据：将汇总的excel数据，插入最后章节
try:
    from modules import add_statistic_result as _add_statistic_result
except Exception as e:
    _report_embedder = None
    print(f"⚠️  未找到 add_statistic_result 模块：{e}")


# 更新生成报告目录：更新目录结构
try:
    from modules import update_dic_uno as _update_dic_uno
except Exception as e:
    _report_embedder = None
    print(f"⚠️  未找到 update_dic_uno 模块：{e}")

# 创建日志记录器实例
log = _ut.Logger()


def generate_report(config:configparser.ConfigParser()):
    # ---------- 3. Excel 数据表转换为 JPG 图像 ----------
    # 说明：
    # excel_to_images 模块应提供 run(input_path, pdfs_dir, images_dir) 接口
    # 功能：将 Excel 文件的每个工作表转换成对应的高分辨率图像文件
    log.info("开始执行 Excel → JPG 转换任务 ...", "ExcelToImages")
    _excel_to_images.run(config)
    log.info("Excel → JPG 转换任务完成", "ExcelToImages")

    # ---------- 4. Word 模板嵌入图片生成报告 ----------
    # 说明：
    # report_embedder 模块应提供 run(template_path, images_dir, output_dir) 接口
    # 功能：将生成的图片嵌入 Word 模板中的表格占位符位置，输出最终巡检报告
    log.info("开始执行 Word 模板嵌入任务 ...", "ReportEmbedder")
    _report_embedder.run(config)
    log.info("Word 模板嵌入任务完成", "ReportEmbedder")
    # ---------- 5. 添加统计汇总
    # 任务 ----------
    log.info("开始执行 添加统计汇总开始 ...", "AddStatisticResult")
    _add_statistic_result.run(config)
    log.info("添加统计汇总完成", "AddStatisticResult")
    # ---------- 6. 生成报告更新目录任务 ----------
    log.info("开始执行 更新目录任务 ...", "cmd call UpdateDicUno")
    generate_report_dic_cmd(config)
    log.info("更新目录任务完成", "cmd call UpdateDicUno")
    # ---------- 6. 结束 ----------
    log.info("=== 巡检报告生成器任务完成 ===")

# ============================================================
# 主程序入口函数
# ============================================================
def main():
    """
    巡检报告生成主入口
    --------------------------------
    步骤说明：
    1. 初始化路径与日志
    2. 加载配置文件 config.ini
    3. 执行服务。
    4. 执行生成巡检报告
    """
    print(f">> main()")
    log.info("=== 巡检报告生成器启动 ===")

    # ---------- 初始化 ----------
    global CONFIG
    # modules_dir：当前脚本所在目录（通常为 detection/modules/）
    modules_dir = os.path.dirname(os.path.abspath(__file__))
    # config_path：配置文件路径（项目根目录下的 config/config.ini）
    config_path = os.path.join(PROJECT_ROOT, "config", "config.ini")
    print(f">> modules_dir = {modules_dir} \n>> config_path = {config_path}")

    # ---------- 读取配置文件 ----------
    CONFIG = configparser.ConfigParser()
    if not os.path.exists(config_path):
        log.warn(f"配置文件未找到：{config_path}")
        sys.exit(1)
    # 加载配置文件并打印内容
    CONFIG.read(config_path, encoding="utf-8")
    log.info(f"配置文件读取成功：{config_path}，配置文件内容如下：", "config")
    log.show_config(CONFIG, "config")   # 调用 Logger 类的 show_config 方法打印配置详情

    # ---------- 执行操作 ----------
    # 获取服务器配置。
    server_run = CONFIG.get("ServerConf", "server")
    generate_report(CONFIG)
# ============================================================
# 程序启动入口
# ============================================================

def generate_report_dic_cmd(config:configparser.ConfigParser()):
    template_path = config.get("Path", "template_path")
    output_dir = config.get("Path", "output_dir")
    script_path = os.path.join(os.path.dirname(__file__), "update_dic_uno.py")
    subprocess.run(["python3", script_path, template_path, output_dir], check=True)

if __name__ == "__main__":
    main()
