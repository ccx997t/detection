#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
巡检报告生成服务进程（server_detection.py）
------------------------------------------------
功能说明：
    1. 接收 UI 端提交的巡检报告基础信息（JSON 格式）
    2. 分配唯一报告编号 report_id
    3. 自动触发巡检报告生成流程：生成封面、采集数据、汇总生成报告
启动方式：
    uvicorn server_detection:app --host 0.0.0.0 --port 8000 --reload
"""
# ============================================================
# 导入模块
# ============================================================
from fastapi import FastAPI                          # 导入 FastAPI 框架
from pydantic import BaseModel                       # 导入 Pydantic 用于定义请求模型
from datetime import date                            # 导入日期类型
import uuid                                          # 导入 uuid 用于生成报告编号
import uvicorn
import configparser
import os,sys

# PROJECT_ROOT 指向项目的根目录，以便导入 modules 下的自定义模块
PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
# 工具模块（日志、配置打印等）
try:
    from util import Logger
except Exception as e:
    print(f"⚠️  未找到 util 模块：{e}")
# 导入业务模块接口
try:
    from report_embedder import create_report_cover      # 封面生成模块
except Exception as e:
    print(f"⚠️  未找到 report_embedder 模块：{e}")
try:
    from get_data_for_sheet import run_data_fill_pipeline # 数据填表模块
except Exception as e:
    print(f"⚠️  未找到 get_data_for_sheet 模块：{e}")
try:
    from detection_report_gen import generate_report # 报告汇总模块
except Exception as e:
    print(f"⚠️  未找到 detection_report_gen 模块：{e}")

# 创建日志记录器实例
log = Logger()

# 全局变量
CONFIG = None

# ============================================================
# FastAPI 应用定义
# ============================================================
app = FastAPI(title="巡检报告生成服务进程", version="1.0")

# ============================================================
# 定义请求体模型
# ============================================================
class ReportInfo(BaseModel):
    """前端传入的巡检报告基础信息"""
    project_name: str          # 项目名称
    room_name: str             # 机房名称
    year: int                  # 年度
    quarter: str               # 季度
    report_date: date          # 上报日期
    report_person: str         # 上报责任人

# ============================================================
# 接口：提交巡检基础信息并自动生成报告
# ============================================================
@app.post("/api/report/basic-info")
def create_report(info: ReportInfo):
    try:
        report_id = "REP-" + uuid.uuid4().hex[:8].upper()
        print(f"✅ 创建任务报告编号: {report_id} ({info.project_name})")

        # Step 1: 生成封面
        if create_report_cover is None:
            raise RuntimeError("create_report_cover 未正确导入")


        # Step 1: 调用 report_embedder 生成封面
        create_report_cover(CONFIG,info.dict())

        # Step 2: 调用 get_data_for_sheet 采集数据并填表
        #run_data_fill_pipeline(report_id)

        # Step 3: 调用 detection_report_gen 汇总生成最终报告
        generate_report(CONFIG)


        return {"code": 200, "message": "巡检报告生成成功", "data": {
            "report_id": report_id
        }}
    except Exception as e:
        print(f"[ERROR] ❌ 生成失败: {e}")
        return {"code": 500, "message": f"生成失败: {str(e)}"}

def run(config: configparser.ConfigParser):
    """ 模块主执行函数。 """
    global CONFIG
    CONFIG = config
    main()
@app.on_event("startup")
def init_config():
    global CONFIG
    """加载配置文件"""
    # ---------- 1. 初始化 ----------
    # modules_dir：当前脚本所在目录（通常为 detection/modules/）
    modules_dir = os.path.dirname(os.path.abspath(__file__))
    # config_path：配置文件路径（项目根目录下的 config/config.ini）
    config_path = os.path.join(PROJECT_ROOT, "config", "config.ini")
    print(f">> modules_dir = {modules_dir} \n>> config_path = {config_path}")

    # ---------- 2. 读取配置文件 ----------
    config = configparser.ConfigParser()
    if not os.path.exists(config_path):
        log.warn(f"配置文件未找到：{config_path}")
        sys.exit(1)

    # 加载配置文件并打印内容
    config.read(config_path, encoding="utf-8")
    log.info(f"配置文件读取成功：{config_path}，配置文件内容如下：", "config")
    log.show_config(config, "config")   # 调用 Logger 类的 show_config 方法打印配置详情
    CONFIG = config
def main():
    """ 主函数 """
    # 启动 FastAPI 应用服务，监听 8100 端口
    log.info(f"服务监听端口：8100")
    uvicorn.run(f"server_detection:app", host="0.0.0.0", port=8100, reload=True)

# ============================================================
# 程序入口（支持直接运行 python3 server_detection.py）
# ============================================================
if __name__ == "__main__":
    main()
