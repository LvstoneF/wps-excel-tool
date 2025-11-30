#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
根据映射文件中的规则，对源Excel文件进行数据处理和转换，并将最终处理结果输出到指定目录。

Usage:
    python process_task.py
"""

import os
import sys
from modules.file_reader import FileReader
from modules.vulnerability_extractor import VulnerabilityExtractor
from modules.ip_device_mapper import IPDeviceMapper
from modules.ip_replacer import IPReplacer
from modules.report_generator import ReportGenerator
from modules.host_detail_processor import HostDetailProcessor
from constants import VULN_SHEET_KEYWORDS, VULN_SHEET_PREFIX

def main():
    """主函数"""
    # 源文件路径
    source_file = r"C:\Users\Administrator\Desktop\11111\定制_xls\EXCEL\report_定制_20251117132848.xls"
    
    # 映射文件路径
    mapping_file = r"C:\Users\Administrator\Desktop\11111\62.docx"
    
    # 输出目录路径
    output_dir = r"C:\Users\Administrator\Desktop\tmp"
    
    # 检查文件是否存在
    if not os.path.exists(source_file):
        print(f"错误：源文件不存在: {source_file}")
        sys.exit(1)
    
    if not os.path.exists(mapping_file):
        print(f"错误：映射文件不存在: {mapping_file}")
        sys.exit(1)
    
    if not os.path.exists(output_dir):
        print(f"错误：输出目录不存在: {output_dir}")
        sys.exit(1)
    
    # 初始化日志记录函数
    def log(message):
        print(message)
    
    # 初始化各个模块
    file_reader = FileReader()
    vuln_extractor = VulnerabilityExtractor()
    ip_device_mapper = IPDeviceMapper(logger=log)
    ip_replacer = IPReplacer(logger=log)
    report_generator = ReportGenerator(logger=log)
    host_detail_processor = HostDetailProcessor(logger=log)
    
    try:
        # 1. 读取IP设备映射表
        log("1. 读取IP设备映射表...")
        ip_device_map = ip_device_mapper.read_ip_device_mapping(mapping_file)
        
        # 2. 获取源文件的工作表列表
        log("\n2. 获取源文件的工作表列表...")
        sheets = file_reader.get_sheets(source_file)
        log(f"   工作表列表: {sheets}")
        
        # 3. 处理每个工作表
        log("\n3. 处理每个工作表...")
        results = []
        vuln_sheets = []
        
        for sheet in sheets:
            log(f"   处理工作表: {sheet}")
            
            # 读取工作表数据
            rows = file_reader.read_file_rows(source_file, sheet)
            
            # 检查是否是漏洞详情工作表
            is_vuln_sheet = sheet in VULN_SHEET_KEYWORDS or sheet.startswith(VULN_SHEET_PREFIX)
            
            if is_vuln_sheet:
                # 提取漏洞信息
                vulnerabilities = vuln_extractor.extract_vulnerabilities(rows)
                # 转换为列表格式，过滤掉严重程度为"信息"的条目
                result = vuln_extractor.convert_vulnerabilities_to_list(vulnerabilities)
                results.append(result)
                vuln_sheets.append(sheet)
                log(f"   提取到 {len(result)} 条漏洞信息")
        
        # 4. 合并结果并保存
        log("\n4. 合并结果并保存...")
        merged_file = report_generator.merge_and_save_results(source_file, vuln_sheets, results, output_dir)
        
        # 5. 替换IP为设备名称
        log("\n5. 替换IP为设备名称...")
        replaced_file = os.path.join(output_dir, f"替换IP后_{os.path.basename(merged_file)}")
        ip_replacer.replace_ip_with_device(merged_file, replaced_file, ip_device_map)
        
        # 6. 生成主机漏洞统计报告
        log("\n6. 生成主机漏洞统计报告...")
        # 查找主机详情工作表
        host_sheet = None
        for sheet in sheets:
            if "主机" in sheet and "详细" in sheet:
                host_sheet = sheet
                break
        
        if host_sheet:
            # 处理主机详情工作表
            hosts = host_detail_processor.process_host_detail_sheet(source_file, host_sheet)
            # 统计漏洞数量
            vuln_counts = host_detail_processor.count_vulnerabilities_by_ip(results)
            # 生成统计报告
            report_generator.generate_host_vuln_stat_report(source_file, hosts, vuln_counts, ip_device_map, output_dir)
        else:
            log("   未找到主机详情工作表，跳过生成主机漏洞统计报告")
        
        log("\n处理完成！")
        log(f"   合并结果文件: {merged_file}")
        log(f"   替换IP后文件: {replaced_file}")
        if host_sheet:
            log(f"   主机漏洞统计报告已生成")
        
    except Exception as e:
        log(f"\n处理失败: {str(e)}")
        sys.exit(1)

if __name__ == "__main__":
    main()
