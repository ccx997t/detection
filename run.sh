#!/bin/bash

# 进入工程根目录（/home/ubuntu(或yanght)/detection），如果已在detection可忽略
cd "$(dirname "$0")"

# 创建日志目录（如果没有）
mkdir -p log

# 启动主流程脚本
python3 modules/detection_report_gen.py "$@" | tee log/run.log
