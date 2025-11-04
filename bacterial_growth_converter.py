#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
细菌生长曲线数据转换工具 v1.0
===============================

功能：
- 将微孔板阅读器导出的Excel文件转换为标准CSV格式
- 支持48小时生长曲线数据
- 输出格式：Well, Time_s, Time_h, OD

输出格式说明：
- Well: 孔位编号 (A1, A2, ..., H12)
- Time_s: 时间（秒）
- Time_h: 时间（小时）
- OD: 光密度值

作者: AI Assistant
日期: 2025年11月4日
"""

import pandas as pd
import numpy as np
import sys
import os
import re
from pathlib import Path


def convert_bacterial_growth_data(excel_file, output_csv, verbose=True):
    """
    将细菌生长曲线Excel数据转换为CSV格式
    专门针对微孔板阅读器导出的格式
    
    参数:
        excel_file (str): 输入Excel文件路径
        output_csv (str): 输出CSV文件路径
        verbose (bool): 是否显示详细信息
    
    返回:
        bool: 转换成功返回True，失败返回False
    """
    
    try:
        if verbose:
            print(f"🔄 开始处理: {os.path.basename(excel_file)}")
        
        # 读取Excel文件
        df = pd.read_excel(excel_file, sheet_name=0, header=None)
        
        if verbose:
            print(f"📊 Excel数据维度: {df.shape[0]}行 × {df.shape[1]}列")
        
        # 固定的数据结构（根据Excel分析结果）
        header_row = 8      # 第9行（索引8）：孔位名称行
        data_start_row = 9  # 第10行（索引9）：数据开始行
        time_col = 1        # 第2列（索引1）：平均时间列
        data_col_start = 2  # 第3列（索引2）：第一个孔位的OD值
        
        if verbose:
            print(f"🎯 数据结构: 孔位行={header_row+1}, 时间列={time_col+1}, 数据列={data_col_start+1}")
        
        # 提取孔位名称
        well_names = []
        for col_idx in range(data_col_start, df.shape[1]):
            well = df.iloc[header_row, col_idx]
            if pd.notna(well):
                well_names.append(str(well).strip())
            else:
                break
        
        if verbose:
            print(f"📍 提取到{len(well_names)}个孔位")
        
        # 提取时间序列（单位：分钟）
        time_minutes = []
        for row_idx in range(data_start_row, df.shape[0]):
            time_val = df.iloc[row_idx, time_col]
            if pd.notna(time_val):
                try:
                    time_minutes.append(float(time_val))
                except:
                    break
            else:
                break
        
        if verbose:
            print(f"⏰ 提取到{len(time_minutes)}个时间点")
            if len(time_minutes) > 0:
                print(f"   时间范围: {min(time_minutes):.1f} - {max(time_minutes):.1f} 分钟")
                print(f"           {min(time_minutes)/60:.2f} - {max(time_minutes)/60:.2f} 小时")
        
        # 提取OD数据
        output_data = []
        
        for well_idx, well_name in enumerate(well_names):
            col_idx = data_col_start + well_idx
            
            if col_idx >= df.shape[1]:
                break
            
            # 提取这个孔位的所有OD值
            od_values = []
            for row_idx, time_min in enumerate(time_minutes):
                actual_row = data_start_row + row_idx
                if actual_row < df.shape[0]:
                    od_val = df.iloc[actual_row, col_idx]
                    if pd.notna(od_val):
                        try:
                            od_values.append(float(od_val))
                        except:
                            pass
            
            # 添加数据
            if len(od_values) == len(time_minutes):
                for time_min, od_val in zip(time_minutes, od_values):
                    time_s = time_min * 60
                    time_h = time_min / 60
                    output_data.append({
                        'Well': well_name,
                        'Time_s': round(time_s, 1),
                        'Time_h': round(time_h, 3),
                        'OD': round(float(od_val), 4)
                    })
                
                if verbose and well_idx < 5:
                    print(f"✅ 处理孔位 {well_name}: {len(od_values)}个数据点")
        
        if not output_data:
            raise ValueError("❌ 未能提取到有效的OD数据，请检查Excel文件格式")
        
        # 创建DataFrame
        output_df = pd.DataFrame(output_data)
        
        # 按孔位和时间排序
        output_df = output_df.sort_values(['Well', 'Time_s'])
        
        # 统计信息
        unique_wells = output_df['Well'].nunique()
        unique_times = output_df['Time_h'].nunique()
        total_points = len(output_df)
        time_range = f"{output_df['Time_h'].min():.2f} - {output_df['Time_h'].max():.2f}"
        od_range = f"{output_df['OD'].min():.4f} - {output_df['OD'].max():.4f}"
        
        if verbose:
            print(f"\n📈 转换结果统计:")
            print(f"   总数据点: {total_points:,}")
            print(f"   孔位数量: {unique_wells}")
            print(f"   时间点数: {unique_times}")
            print(f"   时间范围: {time_range} 小时")
            print(f"   OD值范围: {od_range}")
        
        # 保存到CSV
        output_df.to_csv(output_csv, index=False, encoding='utf-8')
        
        if verbose:
            print(f"💾 数据已保存到: {os.path.basename(output_csv)}")
            print(f"\n📋 数据预览 (前10行):")
            print(output_df.head(10).to_string(index=False))
        
        return True
        
    except Exception as e:
        print(f"❌ 转换失败: {str(e)}")
        if verbose:
            import traceback
            traceback.print_exc()
        return False


def validate_csv_output(csv_file):
    """
    验证输出的CSV文件格式
    """
    try:
        df = pd.read_csv(csv_file)
        
        # 检查必需的列
        required_columns = ['Well', 'Time_s', 'Time_h', 'OD']
        missing_columns = [col for col in required_columns if col not in df.columns]
        
        if missing_columns:
            print(f"❌ CSV文件缺少必需的列: {missing_columns}")
            return False
        
        # 检查数据类型
        if not pd.api.types.is_numeric_dtype(df['Time_s']):
            print("❌ Time_s列不是数值类型")
            return False
        
        if not pd.api.types.is_numeric_dtype(df['Time_h']):
            print("❌ Time_h列不是数值类型")
            return False
        
        if not pd.api.types.is_numeric_dtype(df['OD']):
            print("❌ OD列不是数值类型")
            return False
        
        print("✅ CSV文件格式验证通过")
        return True
        
    except Exception as e:
        print(f"❌ CSV验证失败: {e}")
        return False


def main():
    """主函数"""
    print("🦠 细菌生长曲线数据转换工具 v1.0")
    print("=" * 50)
    
    if len(sys.argv) == 3:
        excel_file = sys.argv[1]
        output_csv = sys.argv[2]
        
        # 检查输入文件
        if not os.path.exists(excel_file):
            print(f"❌ 错误: 输入文件不存在: {excel_file}")
            sys.exit(1)
        
        # 执行转换
        success = convert_bacterial_growth_data(excel_file, output_csv)
        
        if success:
            # 验证输出
            if validate_csv_output(output_csv):
                print(f"\n🎉 转换完成! 输出文件: {output_csv}")
            else:
                print(f"\n⚠️  转换完成但输出文件可能有问题")
        else:
            print(f"\n❌ 转换失败!")
            sys.exit(1)
            
    else:
        print("\n使用方法:")
        print("python bacterial_growth_converter.py <Excel文件> <输出CSV文件>")
        print("\n示例:")
        print('python bacterial_growth_converter.py "WZ 2025.11.3 OD600(1).xls" "growth_data.csv"')
        print("\n输出格式:")
        print("- Well: 孔位编号 (A1, A2, ..., H12)")
        print("- Time_s: 时间（秒）")
        print("- Time_h: 时间（小时）")
        print("- OD: 光密度值")
        sys.exit(1)


if __name__ == "__main__":
    main()