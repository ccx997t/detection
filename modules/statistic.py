#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
statistic_inspection_table.py
-----------------------------------------
功能：
    对 data/巡检报告数据集(1.0).xlsx 中的“表1”进行统计分析。
    自动识别列名（表头可变），打印统计结果。
    不生成文件。
"""
import pandas as pd
from typing import List
import re
import io
import sys
# 清洗 Excel sheet 目前没有用
def clean_excel_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """通用清洗：只保留可见有效内容"""
    # 删除全空行和全空列
    df = df.dropna(how="all", axis=0)
    df = df.dropna(how="all", axis=1)

    # 去除列名与单元格空格、换行符
    df.columns = [str(c).strip().replace("\n", "").replace(" ", "") for c in df.columns]
    #df = df.applymap(lambda x: str(x).strip() if isinstance(x, str) else x)
    for c in df.columns:
        if df[c].dtype == "object":
            df[c] = df[c].map(lambda x: str(x).strip() if isinstance(x, str) else x)

    # 删除表尾残留：只保留到最后一个“检查结果”非空行
    target_cols = [c for c in df.columns if any(k in str(c) for k in ["检查", "检测", "结果", "结论"])]
    if target_cols:
        c_result = target_cols[0]
        last_valid_idx = df[df[c_result].notna()].index.max()
        if pd.notna(last_valid_idx):
            df = df.iloc[: last_valid_idx + 1]

    # 重新索引
    df = df.reset_index(drop=True)
    return df

#  sheet的表头名列表
def get_excel_sheets(excel_path: str) -> List[str]:
    # 获取 sheet的表头名列表
    xls = pd.ExcelFile(excel_path)
    print(f"成功加载 {excel_path} 文件")
    # 检查 Excel 文件的 sheet 工作区，获取 sheet 名字。
    sheet_names = [s for s in xls.sheet_names]
    # 打印显示 sheet 名字。
    print(f"📘 文件中共检测到 {len(sheet_names)} 个表：{sheet_names}")
    return sheet_names

# 获取表头字典的内容。
def get_columns_dict(df: pd.DataFrame, index: int) -> dict:
    """
    自动匹配关键列名，适配不同表头写法，例如：
    返回映射字典：{'技术指标':..., '说明':..., '检查结果':...}
    """
    if index in [1, 2, 3, 5]:
        # 统计字典：字典包含 3 个 Key-value 字段。
        col_map = {"技术指标": None, "说明": None, "检查结果": None}
        cols = [str(c).strip().replace("\n", "").replace(" ", "") for c in df.columns]
        # 打印显示表头信息
        print(f"📋 检测到表头共 {len(cols)} 项：{cols}")
        # 遍历列
        for c in cols:
            name = str(c)
            # 技术指标列（可识别“设备序列号”“机器序号”）
            if col_map["技术指标"] is None and any(k in name for k in ["指标", "项目", "检查项", "设备序列号", "序列号", "主机名", "机器序号"]):
                col_map["技术指标"] = c
            # 说明列（可识别“数据类型”“运行说明”等）
            elif col_map["说明"] is None and any(k in name for k in ["说明", "内容", "要求", "描述", "类型", "状态"]):
                col_map["说明"] = c
            # 检查结果列（兼容“运行状态”“检测结果”等）
            elif col_map["检查结果"] is None and any(k in name for k in ["检查", "检测", "结果", "结论", "运行状态"]):
                col_map["检查结果"] = c
        # 校验
        if not col_map["技术指标"] or not col_map["检查结果"]:
            raise ValueError(f"❌ 无法识别必要列，请检查表头：{cols}")
        print(f"📋 当前表的列映射 col_map = {col_map}")
    elif index in [6, 7]:
        # 统计字典：字典包含 3 个 Key-value 字段。
        col_map = {"统计指标1": None, "统计指标2": None, "统计指标3": None}
        cols = [str(c).strip().replace("\n", "").replace(" ", "") for c in df.columns]
        for c in cols:
            name = str(c)
            if col_map["统计指标1"] is None and any(k in name for k in ["数据中心"]):
                col_map["统计指标1"] = c
            elif col_map["统计指标2"] is None and any(k in name for k in ["设备类型"]):
                col_map["统计指标2"] = c
            elif col_map["统计指标3"] is None and any(k in name for k in ["设备型号"]):
                col_map["统计指标3"] = c    
            print(f"📋 当前表的列映射 col_map = {col_map}")
    return col_map

# 加载 Excel sheet 数据
def load_table(excel_path: str, sheet_name: str) -> pd.DataFrame:
    """读取 Excel 并进行基础清洗"""
    try:
        df = pd.read_excel(excel_path, sheet_name=sheet_name)
    except Exception as e:
        raise RuntimeError(f"❌ 无法读取文件: {e}")
    # 删除全空行
    df = df.dropna(how="all")
    df = df.reset_index(drop=True)
    return df

# 分析统计逻辑
def analyze_all(df: pd.DataFrame, col_map: dict, index: int) -> dict:
    if index in [1, 2, 3, 5]:
        result = analyze_12345(df, col_map, index)
    elif index in [6, 7]:
        result =analyze_67(df, col_map,index)
    return result

# 分析表1，表2，表3，表5。
def analyze_12345(df: pd.DataFrame, col_map: dict, index: int) -> dict:
    """ 执行巡检统计分析 """  
    print(f" .......... 对 sheet{index} 进行统计分析 ..........") 
    # 从列映射字典中提取关键列名
    c_item = col_map["技术指标"]
    c_desc = col_map["说明"]
    c_result = col_map["检查结果"]

    # 打印 DataFrame 的结构和前几行内容
    print(f" sheet{index}（前 3 行预览）:")
    print(df.head(3).to_string(index=False))
    # 将"检查结果c_result"列转换为字符串，去除空格、换行符，然后提交判断
    s = df[c_result].astype(str).fillna("").str.replace(r"\s+", "", regex=True)
    #print(f" s = {s}")
    # 对 s 异常判定条件
    abnormal_mask = s.apply(
        lambda x: (
            (not re.search(r"(?<!不)正常", x)) and
            any(k in x for k in ["不正常", "异常", "错误", "失败", "需检查", "告警"]) and
            not any(p in x for p in ["无告警"])
        )
    )
    # 对 s 正常判定条件
    normal_mask = s.apply(lambda x: re.search(r"(?<!不)正常", x) is not None) & ~abnormal_mask
    # 根据掩码提取正常和异常记录，异常记录: abnormal_df 项。
    abnormal_df = df[abnormal_mask]
    normal_df   = df[normal_mask]
    print(f"\n 正常记录数：{len(normal_df)} | ⚠️  异常记录数：{len(abnormal_df)}\n")
    # 打印部分样本以人工核查
    if not abnormal_df.empty:
        print("🚨 检测到的异常样本预览：")
        print(abnormal_df[[c_item, c_desc, c_result]].head(5).to_string(index=False))
    else:
        print("✅ 未检测到异常项目。")

    # 总项目数: total 项
    total = len(df)
    # 异常数: abnormal_count 项
    abnormal_count = len(abnormal_df)
    # 正常数: normal_count 项
    normal_count   = len(normal_df)
    # 正常率(%): normal_rate 项
    normal_rate    = round(normal_count / total * 100, 2) if total else 0
    # 异常率(%): abnormal_rate 项
    abnormal_rate  = round(abnormal_count / total * 100, 2) if total else 0
    print(f"\n统计比例 => 正常率: {normal_rate}% | 异常率: {abnormal_rate}% | 总项目: {total}")
    # 检查项": check_items 项
    check_items = "、".join(df[c_item].astype(str).tolist())
    # 异常详细: abnormal_detail 项
    if not abnormal_df.empty:
        abnormal_records = []
        for idx, (_, row) in enumerate(abnormal_df.iterrows(), start=1):
            item = str(row.get(c_item, "")).strip()
            res  = str(row.get(c_result, "")).strip()
            item_str = f"{idx}. {item}（{res}）"
            abnormal_records.append(item_str)
        abnormal_detail = "；".join(abnormal_records)
    else:
        abnormal_detail = ""
    print(f".......... sheet{index} 统计分析完毕 ..........") 
    # 返回 result 结果字典
    return {
        "总项目数": total,
        "检查项": check_items,
        "正常数": normal_count,
        "异常数": abnormal_count,
        "正常率(%)": normal_rate,
        "异常率(%)": abnormal_rate,
        "异常详细": abnormal_detail,
        "异常记录": abnormal_df
    }

# 分析表6，表75。
def analyze_67(df: pd.DataFrame, col_map: dict, index: int) -> dict:
    """
    对表6/表7执行三维度设备统计分析：
        ① 以数据中心为基点的统计
        ② 以设备类型为基点的统计（跨数据中心）
        ③ 以设备型号为基点的统计（跨数据中心）
    分析结果存入 result 字典。
    """
    print(f"\n.......... 对 sheet{index}（设备统计）进行分析 ..........")

    # 统一列名映射（确保兼容）
    df = df.rename(columns={
        col_map.get("统计指标1", "数据中心"): "数据中心",
        col_map.get("统计指标2", "设备类型"): "设备类型",
        col_map.get("统计指标3", "设备型号"): "设备型号"
    })

    # ========== ① 按数据中心统计 ==========
    center_stat = (
        df.groupby("数据中心")
          .size()
          .reset_index(name="设备总数")
          .sort_values(by="设备总数", ascending=False)
          .reset_index(drop=True)
    )

    # ========== ② 按设备类型统计（跨数据中心） ==========
    type_stat = (
        df.groupby("设备类型")
          .size()
          .reset_index(name="设备数量")
          .sort_values(by="设备数量", ascending=False)
          .reset_index(drop=True)
    )

    # ========== ③ 按设备型号统计（跨数据中心） ==========
    model_stat = (
        df.groupby(["设备型号", "设备类型"])
          .size()
          .reset_index(name="数量")
          .sort_values(by=["设备类型", "数量"], ascending=[True, False])
          .reset_index(drop=True)
    )

    # ========== 汇总结果 ==========
    result = {
        "中心统计": center_stat,
        "类型统计": type_stat,
        "型号统计": model_stat
    }

    print(f".......... sheet{index} 统计分析完毕 ..........")
    return result

def print_all(sheet_name: str, result: dict, index: int)-> str:
    """打印统计结果"""
    """捕获 print_1235 的所有输出为字符串"""
    buffer = io.StringIO()


            # 捕获所有 print 输出到字符串中
    from io import StringIO
    old_stdout = sys.stdout
    buffer = StringIO()
    sys.stdout = buffer
    if index in [1, 2, 3, 5]:
        print_1235(sheet_name, result)
    elif index in [6, 7]:
        print_67(sheet_name, result)
    sys.stdout = old_stdout
    summary_text = buffer.getvalue().strip()
    # 获取内容并返回
    return summary_text

# 打印表1，表2，表3，表5。
def print_1235(sheet_name: str, result: dict):
    print(f"\n====== {sheet_name} 巡检统计结果 ======")  
    print(f"总项目数：{result['总项目数']}")
    print(f"检查项：{result['检查项']}")
    print(f"正常数：{result['正常数']}")
    print(f"异常数：{result['异常数']}")
    print(f"正常率：{result['正常率(%)']}%")
    print(f"异常率：{result['异常率(%)']}%")
    if result["异常数"] > 0:
        print("\n--- 异常项目详细 ---")
        print(result["异常记录"].to_string(index=False))
        print(f"\n异常描述汇总：{result['异常详细']}")
    print("=============================")     

# 打印表6，表7。
def print_67(sheet_name: str, result: dict):
    print(f"\n====== {sheet_name} 巡检统计结果 ======")  
    # ① 打印数据中心层统计
    print("\n[Ⅰ] 按数据中心统计：")
    print(result["中心统计"].to_string(index=False))

    # ② 打印设备类型层统计
    print("\n[Ⅱ] 按设备类型统计（跨数据中心）：")
    print(result["类型统计"].to_string(index=False))

    # ③ 打印设备型号层统计
    print("\n[Ⅲ] 按设备型号统计（跨数据中心）：")
    print(result["型号统计"].to_string(index=False))

    # ④ 打印分布说明（结构化输出）
    print("\n📍 各数据中心设备类型分布：")
    df_center = result["中心统计"]
    for _, row in df_center.iterrows():
        print(f"  {row['数据中心']}：共 {row['设备总数']} 台设备")

    print("\n📍 各设备类型在数据中心的分布：")
    df_type = result["类型统计"]
    for _, row in df_type.iterrows():
        print(f"  {row['设备类型']}：共 {row['设备数量']} 台")

    print("\n📍 各型号在不同中心的分布：")
    df_model = result["型号统计"]
    for _, row in df_model.iterrows():
        print(f"  {row['设备型号']}（{row['设备类型']}） - 数量：{row['数量']}")
    print("=============================")

# 遍历 excel 的全部 sheet。
def scan_excel_sheets(excel_path: str, sheet_names: List[str])->str :
    """遍历并统计多个 Excel sheet"""
    results_all = []
    output_lines = []  # ⬅️ 新增：用于收集打印内容
    for i, sheet_name in enumerate(sheet_names, start=1):
        print(f"\n===== ({i}) 开始统计：{sheet_name} =====")
        try:
            # 加载 excel sheet。
            df = load_table(excel_path, sheet_name)
            # 获取表头字典的内容。
            col_map = get_columns_dict(df, i)
            # 分析统计。
            result = analyze_all(df, col_map, i)
            # 打印分析统计结果
            output_lines.append(print_all(sheet_name, result, i))


        except Exception as e:
            print(f"❌ 处理 {sheet_name} 时出错：{e}")
            # 拼接所有输出文本为字符串
    summary_text = "\n".join(output_lines)
    # ⬅️ 返回两种内容：打印汇总文本 + result 列表
    return summary_text
# 主入口函数
def main():
    excel_path = "../data/巡检报告数据集(1.0).xlsx"   # 固定输入路径
    # 获取 excel 的 sheet 更准确名称
    sheet_names = get_excel_sheets(excel_path)
    # 遍历 excel，对每个 sheet 进行 检查统计。
    scan_excel_sheets(excel_path, sheet_names)

# 程序入口
if __name__ == "__main__":
    main()
