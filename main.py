import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import openpyxl
import os
import xlrd
from xlrd import xldate_as_tuple
import datetime
import re
from docx import Document

# 从constants.py导入所有常量
from constants import (
    IP_PATTERN,
    IP_DEVICE_PATTERNS,
    SEVERITY_MAP,
    VULN_SHEET_KEYWORDS,
    VULN_SHEET_PREFIX,
    IP_COLUMN_INDEX,
    HOST_DETAIL_SHEET_KEYWORDS,
    HOST_STAT_SHEET_NAME,
    SEVERITY_LEVELS
)

class FileReader:
    """文件读取器，负责读取不同格式的文件"""
    
    @staticmethod
    def get_sheets(file_path):
        """获取文件的工作表列表，支持xlsx、xls、docx和doc格式"""
        ext = os.path.splitext(file_path)[1].lower()
        sheets = []
        
        try:
            if ext == '.xlsx':
                # 使用openpyxl处理xlsx文件
                workbook = openpyxl.load_workbook(file_path)
                sheets = workbook.sheetnames
                workbook.close()
            elif ext == '.xls':
                # 使用xlrd处理xls文件
                workbook = xlrd.open_workbook(file_path)
                sheets = workbook.sheet_names()
            elif ext in ['.docx', '.doc']:
                # Word文件只有一个"工作表"（整个文档）
                sheets = ["文档内容"]
            else:
                raise Exception(f"不支持的文件格式: {ext}")
            
            return sheets
        except Exception as e:
            raise Exception(f"读取工作表失败: {str(e)}")
    
    @staticmethod
    def read_file_rows(file_path, sheet_name):
        """读取文件的行数据，支持xlsx、xls、docx和doc格式"""
        ext = os.path.splitext(file_path)[1].lower()
        rows = []
        
        try:
            # 检查文件是否存在
            if not os.path.exists(file_path):
                raise Exception(f"文件不存在: {file_path}")
            
            if ext == '.xlsx':
                try:
                    # 使用openpyxl处理xlsx文件
                    workbook = openpyxl.load_workbook(file_path)
                    sheet = workbook[sheet_name]
                    for row in sheet.iter_rows(min_row=1, values_only=True):
                        rows.append(row)
                    workbook.close()
                except Exception as e:
                    raise Exception(f"读取.xlsx文件失败: {str(e)}. 请检查文件是否损坏或格式不正确。")
            elif ext == '.xls':
                try:
                    # 使用xlrd处理xls文件
                    workbook = xlrd.open_workbook(file_path)
                    sheet = workbook.sheet_by_name(sheet_name)
                    for i in range(sheet.nrows):
                        row = []
                        for j in range(sheet.ncols):
                            cell_value = sheet.cell_value(i, j)
                            cell_type = sheet.cell_type(i, j)
                            
                            # 处理日期类型
                            if cell_type == xlrd.XL_CELL_DATE:
                                date_tuple = xldate_as_tuple(cell_value, workbook.datemode)
                                cell_value = datetime.datetime(*date_tuple).strftime('%Y-%m-%d')
                            # 处理数字类型，转换为字符串
                            elif cell_type == xlrd.XL_CELL_NUMBER:
                                # 如果是整数，转换为整数字符串，否则保留原格式
                                if cell_value == int(cell_value):
                                    cell_value = str(int(cell_value))
                                else:
                                    cell_value = str(cell_value)
                            # 处理空值
                            elif cell_type == xlrd.XL_CELL_EMPTY:
                                cell_value = ""
                            
                            row.append(cell_value)
                        rows.append(tuple(row))
                except Exception as e:
                    raise Exception(f"读取.xls文件失败: {str(e)}. 请检查文件是否损坏或格式不正确。")
            elif ext == '.docx':
                try:
                    # 使用python-docx处理docx文件
                    doc = Document(file_path)
                    # 读取文档内容，转换为类似Excel行的格式
                    for para in doc.paragraphs:
                        text = para.text.strip()
                        if text:
                            # 将每个段落作为一行，只有一列
                            rows.append((text,))
                    # 读取表格内容
                    for table in doc.tables:
                        for row in table.rows:
                            cells = [cell.text.strip() for cell in row.cells]
                            rows.append(tuple(cells))
                except Exception as e:
                    raise Exception(f"读取.docx文件失败: {str(e)}. 请检查文件是否损坏或格式不正确。")
            elif ext == '.doc':
                # 使用win32com.client处理doc文件（Windows系统）
                try:
                    import win32com.client
                    # 初始化Word应用程序
                    word = win32com.client.Dispatch('Word.Application')
                    word.Visible = False
                    # 打开doc文件
                    doc = word.Documents.Open(file_path)
                    # 读取文档内容
                    text = doc.Content.Text
                    # 关闭文档
                    doc.Close()
                    # 退出Word应用程序
                    word.Quit()
                    # 将文本按换行符分割成段落
                    paragraphs = text.split('\n')
                    for para in paragraphs:
                        text = para.strip()
                        if text:
                            # 将每个段落作为一行，只有一列
                            rows.append((text,))
                except ImportError:
                    raise Exception("处理.doc文件需要安装pywin32库，请使用pip install pywin32命令安装")
                except Exception as e:
                    raise Exception(f"读取.doc文件失败: {str(e)}. 请检查文件是否损坏或格式不正确。")
            else:
                raise Exception(f"不支持的文件格式: {ext}. 请选择.xlsx、.xls、.docx或.doc格式的文件。")
            
            return rows
        except Exception as e:
            raise Exception(f"读取文件失败: {str(e)}")

class VulnerabilityExtractor:
    """漏洞提取器，负责从文件中提取漏洞信息"""
    
    @staticmethod
    def extract_vulnerabilities(rows):
        """从行数据中提取漏洞信息"""
        vulnerabilities = []
        current_vuln = {}
        
        for row in rows:
            # 跳过空行
            if not any(row):
                continue
            
            # 检查是否是新漏洞标题行（以【数字】开头）
            title_cell = row[1] if len(row) > 1 else row[0]
            if title_cell and isinstance(title_cell, str) and title_cell.startswith("【") and "】" in title_cell:
                # 如果有当前漏洞，先保存
                if current_vuln:
                    vulnerabilities.append(current_vuln)
                # 开始新漏洞
                current_vuln = {
                    "序号": title_cell.split("】")[0][1:],
                    "安全漏洞名称": title_cell.split("】")[1].strip()
                }
            # 处理属性行
            else:
                VulnerabilityExtractor._process_vulnerability_attribute(row, current_vuln)
        
        # 保存最后一个漏洞
        if current_vuln:
            vulnerabilities.append(current_vuln)
        
        return vulnerabilities
    
    @staticmethod
    def _process_vulnerability_attribute(row, current_vuln):
        """处理漏洞属性行，提取属性名称和值"""
        # 检查是否是标准格式（B列是属性名，C列是属性值）
        if len(row) >= 3 and row[1] and isinstance(row[1], str) and row[2]:
            VulnerabilityExtractor._extract_attribute_from_columns(row, current_vuln, 1, 2)
        # 兼容旧格式：A列是属性名，B列是属性值
        elif row[0] and isinstance(row[0], str) and len(row) > 1 and row[1]:
            VulnerabilityExtractor._extract_attribute_from_columns(row, current_vuln, 0, 1)
        # 兼容DOCX格式：单行属性（如"危险级别：高"）
        elif len(row) >= 1 and row[0] and isinstance(row[0], str):
            VulnerabilityExtractor._extract_attribute_from_single_line(row[0], current_vuln)
    
    @staticmethod
    def _extract_attribute_from_columns(row, current_vuln, name_col, value_col):
        """从指定列提取属性名称和值"""
        attr_name = row[name_col].strip()
        attr_value = row[value_col].strip()
        VulnerabilityExtractor._update_vulnerability_attribute(current_vuln, attr_name, attr_value)
    
    @staticmethod
    def _extract_attribute_from_single_line(text, current_vuln):
        """从单行文本中提取属性名称和值"""
        text = text.strip()
        if ":" in text:
            attr_name, attr_value = text.split(":", 1)
            attr_name = attr_name.strip()
            attr_value = attr_value.strip()
            VulnerabilityExtractor._update_vulnerability_attribute(current_vuln, attr_name, attr_value)
    
    @staticmethod
    def _update_vulnerability_attribute(current_vuln, attr_name, attr_value):
        """更新漏洞属性，使用类常量作为危险级别映射"""
        if attr_name == "危险级别":
            # 映射危险级别到严重程度
            current_vuln["严重程度"] = SEVERITY_MAP.get(attr_value, attr_value)
        elif attr_name == "存在主机":
            current_vuln["关联资产/域名"] = attr_value
    
    @staticmethod
    def convert_vulnerabilities_to_list(vulnerabilities):
        """将漏洞字典列表转换为列表格式，过滤掉严重程度为"信息"的条目"""
        result_data = []
        for vuln in vulnerabilities:
            severity = vuln.get("严重程度", "")
            # 只保留严重程度不是"信息"的条目
            if severity != "信息":
                row_data = [
                    vuln.get("序号", ""),
                    vuln.get("安全漏洞名称", ""),
                    vuln.get("关联资产/域名", ""),
                    severity
                ]
                result_data.append(row_data)
        return result_data

class IPDeviceMapper:
    """IP设备映射器，负责从Word文件中读取IP设备映射"""
    
    def __init__(self, logger=None):
        self.logger = logger
    
    def log(self, message):
        """记录日志"""
        if self.logger:
            self.logger(message)
    
    def read_ip_device_mapping(self, doc_path):
        """读取Word文件中的IP设备名称映射表，支持docx和doc格式"""
        try:
            self.log(f"开始读取IP设备映射表: {doc_path}")
            
            # 检查文件是否存在
            if not os.path.exists(doc_path):
                raise Exception(f"文件不存在: {doc_path}")
            
            # 初始化映射字典
            ip_device_map = {}
            
            # 获取文件扩展名
            ext = os.path.splitext(doc_path)[1].lower()
            
            if ext == '.docx':
                ip_device_map = self._read_ip_device_mapping_from_docx(doc_path)
            elif ext == '.doc':
                ip_device_map = self._read_ip_device_mapping_from_doc(doc_path)
            else:
                raise Exception(f"不支持的文件格式: {ext}. 请选择.docx或.doc格式的文件。")
            
            self.log(f"\n读取完成，共找到 {len(ip_device_map)} 个IP设备映射")
            return ip_device_map
            
        except Exception as e:
            self.log(f"读取IP设备映射表失败: {str(e)}")
            raise Exception(f"读取IP设备映射表失败: {str(e)}")
    
    def _read_ip_device_mapping_from_docx(self, docx_path):
        """从docx文件读取IP设备映射"""
        try:
            doc = Document(docx_path)
            ip_device_map = {}
            
            # 遍历所有段落
            self.log("=== 遍历段落 ===")
            for para in doc.paragraphs:
                text = para.text.strip()
                if text:
                    self._match_ip_device_patterns(text, ip_device_map)
            
            # 遍历所有表格
            self.log("\n=== 遍历表格 ===")
            for table_idx, table in enumerate(doc.tables):
                self.log(f"表格 {table_idx+1}，共 {len(table.rows)} 行，{len(table.columns)} 列")
                
                # 遍历表格的每一行
                for row in table.rows:
                    cells = [cell.text.strip() for cell in row.cells]
                    self._extract_ip_device_from_table_row(cells, ip_device_map)
            
            return ip_device_map
        except Exception as e:
            raise Exception(f"读取.docx文件失败: {str(e)}. 请检查文件是否损坏或格式不正确。")
    
    def _read_ip_device_mapping_from_doc(self, doc_path):
        """从doc文件读取IP设备映射"""
        try:
            import win32com.client
            # 初始化Word应用程序
            word = win32com.client.Dispatch('Word.Application')
            word.Visible = False
            
            # 打开doc文件
            doc = word.Documents.Open(doc_path)
            text = doc.Content.Text
            doc.Close()
            word.Quit()
            
            # 将文本按换行符分割成段落
            paragraphs = text.split('\n')
            ip_device_map = {}
            
            # 遍历所有段落
            self.log("=== 遍历段落 ===")
            for para in paragraphs:
                text = para.strip()
                if text:
                    self._match_ip_device_patterns(text, ip_device_map)
            
            return ip_device_map
        except ImportError:
            raise Exception("处理.doc文件需要安装pywin32库，请使用pip install pywin32命令安装")
        except Exception as e:
            raise Exception(f"读取.doc文件失败: {str(e)}. 请检查文件是否损坏或格式不正确。")
    
    def _match_ip_device_patterns(self, text, ip_device_map):
        """匹配IP设备名称模式"""
        for pattern in IP_DEVICE_PATTERNS:
            match = re.match(pattern, text)
            if match:
                ip = match.group(1)
                device = match.group(2).strip()
                ip_device_map[ip] = device
                self.log(f"  匹配到: IP={ip}, 设备名称={device}")
                break
    
    def _extract_ip_device_from_table_row(self, cells, ip_device_map):
        """从表格行提取IP设备映射"""
        # 检查是否至少有2个单元格
        if len(cells) >= 2:
            # 首先检查从右到左：设备名称 -> IP
            for i in range(1, len(cells)):
                # 检查当前单元格是否是IP地址
                if re.match(r'^' + IP_PATTERN + r'$', cells[i]):
                    # 当前单元格是IP地址，前一个单元格是设备名称
                    device_candidate = cells[i-1]
                    if device_candidate and device_candidate != cells[i]:
                        ip = cells[i]
                        device = device_candidate
                        ip_device_map[ip] = device
                        self.log(f"  匹配到: IP={ip}, 设备名称={device}")
                        return
            
            # 如果从右到左没有匹配到，再尝试从左到右：IP -> 设备名称
            for i in range(len(cells) - 1):
                # 检查当前单元格是否是IP地址
                if re.match(r'^' + IP_PATTERN + r'$', cells[i]):
                    # 当前单元格是IP地址，下一个单元格是设备名称
                    device_candidate = cells[i+1]
                    if device_candidate and device_candidate != cells[i]:
                        ip = cells[i]
                        device = device_candidate
                        ip_device_map[ip] = device
                        self.log(f"  匹配到: IP={ip}, 设备名称={device}")
                        return

class IPReplacer:
    """IP替换器，负责替换文件中的IP地址为设备名称"""
    
    def __init__(self, logger=None):
        self.logger = logger
    
    def log(self, message):
        """记录日志"""
        if self.logger:
            self.logger(message)
    
    def replace_ip_with_device(self, input_file, output_file, ip_device_map):
        """替换Excel文件中的IP为设备名称，支持.xlsx和.xls格式"""
        try:
            self.log(f"开始替换IP为设备名称: {input_file}")
            
            # 获取文件扩展名
            ext = os.path.splitext(input_file)[1].lower()
            
            if ext == '.xlsx':
                # 使用openpyxl处理.xlsx文件
                workbook = openpyxl.load_workbook(input_file)
                sheet = workbook.active
                
                # 查找IP列（关联资产/域名列）
                ip_column = IP_COLUMN_INDEX  # 使用类常量
                
                # 遍历所有行，从第2行开始（跳过表头）
                replaced_count = 0
                for row in range(2, sheet.max_row + 1):
                    cell_value = sheet.cell(row=row, column=ip_column + 1).value  # Excel列从1开始
                    if cell_value and isinstance(cell_value, str):
                        # 检查是否是IP地址
                        ip_pattern = IP_PATTERN
                        # 查找所有IP地址
                        ips = re.findall(ip_pattern, cell_value)
                        if ips:
                            # 替换每个IP地址
                            modified_value = cell_value
                            for ip in ips:
                                if ip in ip_device_map:
                                    device_name = ip_device_map[ip]
                                    # 只在设备名称有效（非空且不是"/"）的情况下进行替换
                                    if device_name and device_name != "/" and device_name != ip:
                                        modified_value = modified_value.replace(ip, device_name)
                                        replaced_count += 1
                                        self.log(f"  替换IP {ip} 为 {device_name}")
                                    else:
                                        self.log(f"  跳过替换IP {ip}，设备名称无效: {device_name}")
                            # 如果有修改，更新单元格值
                            if modified_value != cell_value:
                                sheet.cell(row=row, column=ip_column + 1).value = modified_value
                
                # 保存输出文件
                workbook.save(output_file)
                workbook.close()
                
            elif ext == '.xls':
                # 使用xlrd读取.xls文件，然后使用openpyxl写入新的.xlsx文件
                self.log(f"检测到.xls文件，正在转换处理: {input_file}")
                
                # 读取.xls文件
                xlrd_workbook = xlrd.open_workbook(input_file)
                xlrd_sheet = xlrd_workbook.sheet_by_index(0)
                
                # 创建新的.xlsx文件
                openpyxl_workbook = openpyxl.Workbook()
                openpyxl_sheet = openpyxl_workbook.active
                
                # 复制表头
                header_row = xlrd_sheet.row_values(0)
                openpyxl_sheet.append(header_row)
                
                # 遍历所有行，从第2行开始（跳过表头）
                replaced_count = 0
                for row_idx in range(1, xlrd_sheet.nrows):
                    row_values = xlrd_sheet.row_values(row_idx)
                    
                    # 处理IP列（关联资产/域名列）
                    ip_column = IP_COLUMN_INDEX  # 使用类常量
                    if len(row_values) > ip_column:
                        cell_value = row_values[ip_column]
                        if cell_value and isinstance(cell_value, str):
                            # 检查是否是IP地址
                            ip_pattern = IP_PATTERN
                            # 查找所有IP地址
                            ips = re.findall(ip_pattern, cell_value)
                            if ips:
                                # 替换每个IP地址
                                modified_value = cell_value
                                for ip in ips:
                                    if ip in ip_device_map:
                                        device_name = ip_device_map[ip]
                                        # 只在设备名称有效（非空且不是"/"）的情况下进行替换
                                        if device_name and device_name != "/" and device_name != ip:
                                            modified_value = modified_value.replace(ip, device_name)
                                            replaced_count += 1
                                            self.log(f"  替换IP {ip} 为 {device_name}")
                                        else:
                                            self.log(f"  跳过替换IP {ip}，设备名称无效: {device_name}")
                                # 如果有修改，更新单元格值
                                row_values[ip_column] = modified_value
                    
                    # 将处理后的行添加到新工作表
                    openpyxl_sheet.append(row_values)
                
                # 保存输出文件
                openpyxl_workbook.save(output_file)
                openpyxl_workbook.close()
                
            else:
                raise Exception(f"不支持的文件格式: {ext}")
            
            self.log(f"IP替换完成！共替换 {replaced_count} 个IP地址")
            self.log(f"替换结果保存至: {output_file}")
            return output_file
        except Exception as e:
            self.log(f"替换IP失败: {str(e)}")
            raise Exception(f"替换IP失败: {str(e)}")
    
    def _replace_ips_in_text(self, text, ip_device_map):
        """替换文本中的IP为设备名称"""
        replaced_count = 0
        modified_value = text
        # 查找所有IP地址
        ips = re.findall(IP_PATTERN, text)
        if ips:
            # 替换每个IP地址
            for ip in ips:
                if ip in ip_device_map:
                    device_name = ip_device_map[ip]
                    # 只在设备名称有效（非空且不是"/"）的情况下进行替换
                    if device_name and device_name != "/" and device_name != ip:
                        modified_value = modified_value.replace(ip, device_name)
                        replaced_count += 1
                        self.log(f"  替换IP {ip} 为 {device_name}")
                    else:
                        self.log(f"  跳过替换IP {ip}，设备名称无效: {device_name}")
        return modified_value, replaced_count

class ReportGenerator:
    """报告生成器，负责生成漏洞统计报告"""
    
    def __init__(self, logger=None):
        self.logger = logger
    
    def log(self, message):
        """记录日志"""
        if self.logger:
            self.logger(message)
    
    def merge_and_save_results(self, file_path, sheet_names, results, output_path):
        """合并多个工作表的处理结果并保存"""
        try:
            # 创建新工作簿和工作表
            new_workbook = openpyxl.Workbook()
            new_sheet = new_workbook.active
            new_sheet.title = "合并漏洞详情处理结果"
            
            # 定义表头
            headers = ["序号", "安全漏洞名称", "关联资产/域名", "严重程度"]
            new_sheet.append(headers)
            
            # 设置列宽
            column_widths = [10, 50, 20, 10]
            for i, width in enumerate(column_widths):
                new_sheet.column_dimensions[openpyxl.utils.get_column_letter(i+1)].width = width
            
            # 合并所有结果
            total_rows = 0
            for i, (sheet_name, result_data) in enumerate(zip(sheet_names, results)):
                self.log(f"合并{sheet_name}的处理结果，共{len(result_data)}行")
                # 写入数据（跳过表头，因为已经添加过了）
                for row in result_data:
                    new_sheet.append(row)
                    total_rows += 1
            
            # 保存新文件，确保使用.xlsx扩展名
            base_name = os.path.basename(file_path)
            name_without_ext = os.path.splitext(base_name)[0]
            output_file = os.path.join(output_path, f"合并处理结果_{name_without_ext}.xlsx")
            new_workbook.save(output_file)
            
            self.log(f"合并完成！共处理{len(sheet_names)}个工作表，生成{total_rows}行数据")
            self.log(f"结果保存至: {output_file}")
            return output_file
        except Exception as e:
            raise Exception(f"合并结果失败: {str(e)}")
    
    def generate_host_vuln_stat_report(self, file_path, hosts, vuln_counts, ip_device_map, output_path):
        """生成主机漏洞统计报告"""
        try:
            self.log("开始生成主机漏洞统计报告")
            
            # 创建新工作簿和工作表
            new_workbook = openpyxl.Workbook()
            new_sheet = new_workbook.active
            new_sheet.title = HOST_STAT_SHEET_NAME
            
            # 定义表头
            headers = [
                "序号", 
                "设备名称或IP地址", 
                "系统及版本", 
                "安全漏洞数量", "", "", ""
            ]
            new_sheet.append(headers)
            
            # 定义二级表头
            sub_headers = [
                "", "", "", 
                "高", "中", "低", "小计"
            ]
            new_sheet.append(sub_headers)
            
            # 合并单元格
            new_sheet.merge_cells(start_row=1, start_column=4, end_row=1, end_column=7)
            
            # 设置列宽
            column_widths = [10, 30, 20, 10, 10, 10, 10]
            for i, width in enumerate(column_widths):
                new_sheet.column_dimensions[openpyxl.utils.get_column_letter(i+1)].width = width
            
            # 准备统计数据
            stat_data = []
            
            # 遍历主机列表，按照主机在列表中的顺序进行统计
            for host in hosts:
                # 查找IP地址
                ip = None
                for key, value in host.items():
                    if re.match(IP_PATTERN, value):
                        ip = value
                        break
                
                if not ip:
                    continue
                
                # 获取设备名称
                device_name = ip_device_map.get(ip, ip)
                
                # 获取系统及版本
                os_version = ""
                for key, value in host.items():
                    if "系统" in key or "版本" in key or "OS" in key:
                        os_version = value
                        break
                
                # 获取漏洞统计数据
                counts = vuln_counts.get(ip, {"高": 0, "中": 0, "低": 0})
                
                # 计算小计
                total = counts["高"] + counts["中"] + counts["低"]
                
                # 添加到统计数据
                stat_data.append({
                    "ip": ip,
                    "device_name": device_name,
                    "os_version": os_version,
                    "counts": counts,
                    "total": total
                })
            
            # 遍历映射表中未在主机列表中出现的设备
            for ip, device_name in ip_device_map.items():
                if ip not in [item["ip"] for item in stat_data]:
                    # 获取漏洞统计数据
                    counts = vuln_counts.get(ip, {"高": 0, "中": 0, "低": 0})
                    
                    # 计算小计
                    total = counts["高"] + counts["中"] + counts["低"]
                    
                    # 添加到统计数据
                    stat_data.append({
                        "ip": ip,
                        "device_name": device_name,
                        "os_version": "",
                        "counts": counts,
                        "total": total
                    })
            
            # 写入统计数据
            for i, data in enumerate(stat_data, start=1):
                row = [
                    i,
                    data["device_name"],
                    data["os_version"],
                    data["counts"]["高"],
                    data["counts"]["中"],
                    data["counts"]["低"],
                    data["total"]
                ]
                new_sheet.append(row)
            
            # 保存新文件，确保使用.xlsx扩展名
            base_name = os.path.basename(file_path)
            name_without_ext = os.path.splitext(base_name)[0]
            output_file = os.path.join(output_path, f"主机漏洞统计_{name_without_ext}.xlsx")
            new_workbook.save(output_file)
            
            self.log(f"主机漏洞统计报告生成完成！共统计 {len(stat_data)} 台设备")
            self.log(f"统计报告保存至: {output_file}")
            return output_file
        except Exception as e:
            self.log(f"生成主机漏洞统计报告失败: {str(e)}")
            raise Exception(f"生成主机漏洞统计报告失败: {str(e)}")

class HostDetailProcessor:
    """主机详情处理器，负责处理主机详情子表"""
    
    def __init__(self, logger=None):
        self.logger = logger
    
    def log(self, message):
        """记录日志"""
        if self.logger:
            self.logger(message)
    
    def process_host_detail_sheet(self, file_path):
        """处理主机详情子表，获取主机信息"""
        try:
            # 获取所有工作表
            sheets = FileReader.get_sheets(file_path)
            
            # 查找主机详情子表
            host_sheet = None
            for sheet in sheets:
                if any(keyword in sheet for keyword in HOST_DETAIL_SHEET_KEYWORDS):
                    host_sheet = sheet
                    break
            
            if not host_sheet:
                self.log("未找到主机详情子表")
                return []
            
            self.log(f"开始处理主机详情子表: {host_sheet}")
            
            # 读取主机详情子表数据
            rows = FileReader.read_file_rows(file_path, host_sheet)
            
            # 提取主机信息
            hosts = []
            header = None
            
            for i, row in enumerate(rows):
                if not any(row):
                    continue
                
                # 查找表头行
                if not header:
                    # 查找包含"IP地址"或"主机"的行作为表头
                    if any(cell and isinstance(cell, str) and ("IP地址" in cell or "主机" in cell or "设备" in cell) for cell in row):
                        header = [str(cell).strip() if cell else "" for cell in row]
                        self.log(f"找到主机详情表头: {header}")
                    continue
                
                # 提取主机信息
                host_info = {}
                for j, cell_value in enumerate(row):
                    if j < len(header):
                        host_info[header[j]] = str(cell_value).strip() if cell_value else ""
                
                hosts.append(host_info)
            
            self.log(f"成功提取 {len(hosts)} 台主机信息")
            return hosts
        except Exception as e:
            self.log(f"处理主机详情子表失败: {str(e)}")
            return []
    
    def count_vulnerabilities_by_ip(self, results):
        """按IP地址统计不同严重程度的漏洞数量"""
        # 初始化统计字典
        vuln_counts = {}
        
        # 遍历所有处理结果
        for result in results:
            for row in result:
                if len(row) < 4:
                    continue
                
                # 提取IP地址和严重程度
                ip_text = row[2].strip()
                severity = row[3].strip()
                
                # 只统计高、中、低三个级别
                if severity not in SEVERITY_LEVELS:
                    continue
                
                # 从文本中提取所有IP地址
                ips = re.findall(IP_PATTERN, ip_text)
                
                for ip in ips:
                    if ip:
                        # 初始化IP的统计数据
                        if ip not in vuln_counts:
                            vuln_counts[ip] = {
                                "高": 0,
                                "中": 0,
                                "低": 0
                            }
                        
                        # 更新统计数据
                        vuln_counts[ip][severity] += 1
        
        self.log(f"成功统计 {len(vuln_counts)} 个IP的漏洞数量")
        return vuln_counts

class WPSExcelTool:
    
    def __init__(self, root):
        self.root = root
        self.root.title("WPS Excel 处理工具")
        
        # 设置窗口为可调整大小
        self.root.resizable(True, True)
        
        # 设置窗口最小大小
        self.root.minsize(600, 400)
        
        # 初始化变量
        self.file_path = tk.StringVar()
        self.output_path = tk.StringVar()
        self.mapping_file_path = tk.StringVar()  # IP设备映射表文件路径
        
        # 初始化抽象类实例
        self.ip_device_mapper = IPDeviceMapper(logger=self.log)
        self.ip_replacer = IPReplacer(logger=self.log)
        self.report_generator = ReportGenerator(logger=self.log)
        self.host_detail_processor = HostDetailProcessor(logger=self.log)
        
        # 创建界面
        self.create_widgets()
        
        # 设置窗口居中显示
        self.center_window()
        
    def center_window(self):
        """将窗口居中显示"""
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f"{width}x{height}+{x}+{y}")
    
    def create_widgets(self):
        # 标题
        title_label = ttk.Label(self.root, text="WPS Excel 处理工具", font=("Arial", 16))
        title_label.pack(pady=20)
        
        # 文件选择
        file_frame = ttk.Frame(self.root)
        file_frame.pack(pady=5, padx=10, fill=tk.X)
        
        ttk.Label(file_frame, text="选择Excel文件:").pack(side=tk.LEFT, padx=5, anchor=tk.CENTER)
        # 让输入框能够扩展
        ttk.Entry(file_frame, textvariable=self.file_path).pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        ttk.Button(file_frame, text="浏览", command=self.browse_file).pack(side=tk.LEFT, padx=5)
        
        # 工作表选择 - 支持多选
        sheet_frame = ttk.LabelFrame(self.root, text="选择工作表（可多选）")
        sheet_frame.pack(pady=5, padx=10, fill=tk.BOTH, expand=True)
        
        # 刷新按钮
        refresh_button = ttk.Button(sheet_frame, text="刷新工作表", command=self.refresh_sheets)
        refresh_button.pack(anchor=tk.NE, padx=5, pady=5)
        
        # 列表框和滚动条
        listbox_frame = ttk.Frame(sheet_frame)
        listbox_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 垂直滚动条
        v_scrollbar = ttk.Scrollbar(listbox_frame, orient=tk.VERTICAL)
        v_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 水平滚动条
        h_scrollbar = ttk.Scrollbar(listbox_frame, orient=tk.HORIZONTAL)
        h_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)
        
        # 列表框，支持多选，去掉固定宽度，让它自动适应
        self.sheet_listbox = tk.Listbox(
            listbox_frame,
            selectmode=tk.MULTIPLE,
            yscrollcommand=v_scrollbar.set,
            xscrollcommand=h_scrollbar.set,
            height=8
        )
        self.sheet_listbox.pack(fill=tk.BOTH, expand=True)
        
        # 绑定滚动条
        v_scrollbar.config(command=self.sheet_listbox.yview)
        h_scrollbar.config(command=self.sheet_listbox.xview)
        
        # 全选和取消全选按钮
        select_frame = ttk.Frame(sheet_frame)
        select_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Button(select_frame, text="全选", command=self.select_all_sheets).pack(side=tk.LEFT, padx=5)
        ttk.Button(select_frame, text="取消全选", command=self.deselect_all_sheets).pack(side=tk.LEFT, padx=5)
        ttk.Button(select_frame, text="选择漏洞相关工作表", command=self.select_vuln_sheets).pack(side=tk.LEFT, padx=5)
        
        # 输出路径
        output_frame = ttk.Frame(self.root)
        output_frame.pack(pady=5, padx=10, fill=tk.X)
        
        ttk.Label(output_frame, text="输出路径:").pack(side=tk.LEFT, padx=5, anchor=tk.CENTER)
        # 让输入框能够扩展
        ttk.Entry(output_frame, textvariable=self.output_path).pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        ttk.Button(output_frame, text="浏览", command=self.browse_output).pack(side=tk.LEFT, padx=5)
        
        # IP设备映射表选择
        mapping_frame = ttk.Frame(self.root)
        mapping_frame.pack(pady=5, padx=10, fill=tk.X)
        
        ttk.Label(mapping_frame, text="IP设备映射表:").pack(side=tk.LEFT, padx=5, anchor=tk.CENTER)
        mapping_entry = ttk.Entry(mapping_frame, textvariable=self.mapping_file_path)
        mapping_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        ttk.Button(mapping_frame, text="浏览", command=self.browse_mapping_file).pack(side=tk.LEFT, padx=5)
        
        # 处理按钮
        button_frame = ttk.Frame(self.root)
        button_frame.pack(pady=20)
        
        ttk.Button(button_frame, text="处理文档", command=self.process_file).pack(side=tk.LEFT, padx=10)
        ttk.Button(button_frame, text="退出", command=self.root.quit).pack(side=tk.LEFT, padx=10)
        
        # 日志区域
        log_frame = ttk.Frame(self.root)
        log_frame.pack(pady=5, padx=10, fill=tk.BOTH, expand=True)
        
        ttk.Label(log_frame, text="处理日志:").pack(anchor=tk.W, padx=5)
        # 去掉固定的高度和宽度，让它自动适应
        self.log_text = tk.Text(log_frame, wrap=tk.WORD)
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 滚动条
        scrollbar = ttk.Scrollbar(self.log_text, orient=tk.VERTICAL, command=self.log_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.config(yscrollcommand=scrollbar.set)
    
    def browse_file(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel Files", "*.xlsx;*.xls"), ("Word Files", "*.docx;*.doc")]
        )
        if file_path:
            self.file_path.set(file_path)
            self.refresh_sheets()
            self.log(f"选择文件: {file_path}")
    
    def browse_output(self):
        output_path = filedialog.askdirectory()
        if output_path:
            self.output_path.set(output_path)
            self.log(f"选择输出路径: {output_path}")
    
    def browse_mapping_file(self):
        """选择IP设备映射表文件"""
        mapping_file = filedialog.askopenfilename(
            filetypes=[("Word Files", "*.docx;*.doc")]
        )
        if mapping_file:
            self.mapping_file_path.set(mapping_file)
            self.log(f"选择IP设备映射表: {mapping_file}")
    
    def get_sheets(self, file_path):
        """获取文件的工作表列表，支持xlsx、xls、docx和doc格式"""
        return FileReader.get_sheets(file_path)
    
    def read_ip_device_mapping(self, doc_path):
        """读取Word文件中的IP设备名称映射表，支持docx和doc格式"""
        return self.ip_device_mapper.read_ip_device_mapping(doc_path)
    
    def select_all_sheets(self):
        """全选所有工作表"""
        self.sheet_listbox.select_set(0, tk.END)
        self.log("已全选所有工作表")
    
    def deselect_all_sheets(self):
        """取消全选所有工作表"""
        self.sheet_listbox.selection_clear(0, tk.END)
        self.log("已取消全选所有工作表")
    
    def select_vuln_sheets(self):
        """选择所有漏洞相关工作表"""
        # 清空当前选择
        self.sheet_listbox.selection_clear(0, tk.END)
        
        # 获取所有工作表
        file_path = self.file_path.get()
        if not file_path:
            messagebox.showwarning("警告", "请先选择Excel文件")
            return
        
        try:
            sheets = self.get_sheets(file_path)
            
            # 选择漏洞相关工作表
            vuln_sheets = [sheet for sheet in sheets if sheet.startswith(VULN_SHEET_PREFIX) or sheet in VULN_SHEET_KEYWORDS]
            
            # 在列表框中选择对应的项
            for i, sheet in enumerate(sheets):
                if sheet in vuln_sheets:
                    self.sheet_listbox.select_set(i)
            
            self.log(f"已选择 {len(vuln_sheets)} 个漏洞相关工作表")
        except Exception as e:
            messagebox.showerror("错误", str(e))
            self.log(str(e))
    
    def refresh_sheets(self):
        """刷新工作表列表"""
        file_path = self.file_path.get()
        if not file_path:
            messagebox.showwarning("警告", "请先选择Excel文件")
            return
        
        try:
            sheets = self.get_sheets(file_path)
            
            # 清空列表框
            self.sheet_listbox.delete(0, tk.END)
            
            # 添加工作表到列表框
            for sheet in sheets:
                self.sheet_listbox.insert(tk.END, sheet)
            
            self.log(f"刷新工作表完成: {sheets}")
        except Exception as e:
            messagebox.showerror("错误", str(e))
            self.log(str(e))
    
    def read_file_rows(self, file_path, sheet_name):
        """读取文件的行数据，支持xlsx、xls、docx和doc格式"""
        ext = os.path.splitext(file_path)[1].lower()
        rows = []
        
        try:
            # 检查文件是否存在
            if not os.path.exists(file_path):
                raise Exception(f"文件不存在: {file_path}")
            
            if ext == '.xlsx':
                try:
                    # 使用openpyxl处理xlsx文件
                    workbook = openpyxl.load_workbook(file_path)
                    sheet = workbook[sheet_name]
                    for row in sheet.iter_rows(min_row=1, values_only=True):
                        rows.append(row)
                    workbook.close()
                except Exception as e:
                    raise Exception(f"读取.xlsx文件失败: {str(e)}. 请检查文件是否损坏或格式不正确。")
            elif ext == '.xls':
                try:
                    # 使用xlrd处理xls文件
                    workbook = xlrd.open_workbook(file_path)
                    sheet = workbook.sheet_by_name(sheet_name)
                    for i in range(sheet.nrows):
                        row = []
                        for j in range(sheet.ncols):
                            cell_value = sheet.cell_value(i, j)
                            cell_type = sheet.cell_type(i, j)
                            
                            # 处理日期类型
                            if cell_type == xlrd.XL_CELL_DATE:
                                date_tuple = xldate_as_tuple(cell_value, workbook.datemode)
                                cell_value = datetime.datetime(*date_tuple).strftime('%Y-%m-%d')
                            # 处理数字类型，转换为字符串
                            elif cell_type == xlrd.XL_CELL_NUMBER:
                                # 如果是整数，转换为整数字符串，否则保留原格式
                                if cell_value == int(cell_value):
                                    cell_value = str(int(cell_value))
                                else:
                                    cell_value = str(cell_value)
                            # 处理空值
                            elif cell_type == xlrd.XL_CELL_EMPTY:
                                cell_value = ""
                            
                            row.append(cell_value)
                        rows.append(tuple(row))
                except Exception as e:
                    raise Exception(f"读取.xls文件失败: {str(e)}. 请检查文件是否损坏或格式不正确。")
            elif ext == '.docx':
                try:
                    # 使用python-docx处理docx文件
                    doc = Document(file_path)
                    # 读取文档内容，转换为类似Excel行的格式
                    for para in doc.paragraphs:
                        text = para.text.strip()
                        if text:
                            # 将每个段落作为一行，只有一列
                            rows.append((text,))
                    # 读取表格内容
                    for table in doc.tables:
                        for row in table.rows:
                            cells = [cell.text.strip() for cell in row.cells]
                            rows.append(tuple(cells))
                except Exception as e:
                    raise Exception(f"读取.docx文件失败: {str(e)}. 请检查文件是否损坏或格式不正确。")
            elif ext == '.doc':
                # 使用win32com.client处理doc文件（Windows系统）
                try:
                    import win32com.client
                    # 初始化Word应用程序
                    word = win32com.client.Dispatch('Word.Application')
                    word.Visible = False
                    # 打开doc文件
                    doc = word.Documents.Open(file_path)
                    # 读取文档内容
                    text = doc.Content.Text
                    # 关闭文档
                    doc.Close()
                    # 退出Word应用程序
                    word.Quit()
                    # 将文本按换行符分割成段落
                    paragraphs = text.split('\n')
                    for para in paragraphs:
                        text = para.strip()
                        if text:
                            # 将每个段落作为一行，只有一列
                            rows.append((text,))
                except ImportError:
                    raise Exception("处理.doc文件需要安装pywin32库，请使用pip install pywin32命令安装")
                except Exception as e:
                    raise Exception(f"读取.doc文件失败: {str(e)}. 请检查文件是否损坏或格式不正确。")
            else:
                raise Exception(f"不支持的文件格式: {ext}. 请选择.xlsx、.xls、.docx或.doc格式的文件。")
            
            return rows
        except Exception as e:
            raise Exception(f"读取文件失败: {str(e)}")
    
    def process_single_sheet(self, file_path, sheet_name):
        """处理单个工作表或文档，返回处理结果数据，支持xlsx、xls、docx和doc格式"""
        try:
            # 读取文件行数据
            rows = FileReader.read_file_rows(file_path, sheet_name)
            
            # 检查文件格式
            ext = os.path.splitext(file_path)[1].lower()
            
            # 检查是否需要特殊处理
            is_vulnerability_sheet = (sheet_name in VULN_SHEET_KEYWORDS or sheet_name.startswith(VULN_SHEET_PREFIX))
            is_docx_file = (ext == '.docx')
            
            if is_vulnerability_sheet or is_docx_file:
                # 提取漏洞信息
                vulnerabilities = VulnerabilityExtractor.extract_vulnerabilities(rows)
                # 转换为列表格式，过滤掉严重程度为"信息"的条目
                result_data = VulnerabilityExtractor.convert_vulnerabilities_to_list(vulnerabilities)
                return result_data
            else:
                # 非漏洞详情工作表，返回原始数据
                return rows
        except Exception as e:
            raise Exception(f"处理工作表{sheet_name}失败: {str(e)}")
    
    def _extract_vulnerabilities(self, rows):
        """从行数据中提取漏洞信息"""
        vulnerabilities = []
        current_vuln = {}
        
        for row in rows:
            # 跳过空行
            if not any(row):
                continue
            
            # 检查是否是新漏洞标题行（以【数字】开头）
            title_cell = row[1] if len(row) > 1 else row[0]
            if title_cell and isinstance(title_cell, str) and title_cell.startswith("【") and "】" in title_cell:
                # 如果有当前漏洞，先保存
                if current_vuln:
                    vulnerabilities.append(current_vuln)
                # 开始新漏洞
                current_vuln = {
                    "序号": title_cell.split("】")[0][1:],
                    "安全漏洞名称": title_cell.split("】")[1].strip()
                }
            # 处理属性行
            else:
                self._process_vulnerability_attribute(row, current_vuln)
        
        # 保存最后一个漏洞
        if current_vuln:
            vulnerabilities.append(current_vuln)
        
        return vulnerabilities
    
    def _process_vulnerability_attribute(self, row, current_vuln):
        """处理漏洞属性行，提取属性名称和值"""
        # 检查是否是标准格式（B列是属性名，C列是属性值）
        if len(row) >= 3 and row[1] and isinstance(row[1], str) and row[2]:
            self._extract_attribute_from_columns(row, current_vuln, 1, 2)
        # 兼容旧格式：A列是属性名，B列是属性值
        elif row[0] and isinstance(row[0], str) and len(row) > 1 and row[1]:
            self._extract_attribute_from_columns(row, current_vuln, 0, 1)
        # 兼容DOCX格式：单行属性（如"危险级别：高"）
        elif len(row) >= 1 and row[0] and isinstance(row[0], str):
            self._extract_attribute_from_single_line(row[0], current_vuln)
    
    def _extract_attribute_from_columns(self, row, current_vuln, name_col, value_col):
        """从指定列提取属性名称和值"""
        attr_name = row[name_col].strip()
        attr_value = row[value_col].strip()
        self._update_vulnerability_attribute(current_vuln, attr_name, attr_value)
    
    def _extract_attribute_from_single_line(self, text, current_vuln):
        """从单行文本中提取属性名称和值"""
        text = text.strip()
        if ":" in text:
            attr_name, attr_value = text.split(":", 1)
            attr_name = attr_name.strip()
            attr_value = attr_value.strip()
            self._update_vulnerability_attribute(current_vuln, attr_name, attr_value)
    
    def _update_vulnerability_attribute(self, current_vuln, attr_name, attr_value):
        """更新漏洞属性，使用类常量作为危险级别映射"""
        if attr_name == "危险级别":
            # 映射危险级别到严重程度
            current_vuln["严重程度"] = SEVERITY_MAP.get(attr_value, attr_value)
        elif attr_name == "存在主机":
            current_vuln["关联资产/域名"] = attr_value
    
    def _convert_vulnerabilities_to_list(self, vulnerabilities):
        """将漏洞字典列表转换为列表格式，过滤掉严重程度为"信息"的条目"""
        result_data = []
        for vuln in vulnerabilities:
            severity = vuln.get("严重程度", "")
            # 只保留严重程度不是"信息"的条目
            if severity != "信息":
                row_data = [
                    vuln.get("序号", ""),
                    vuln.get("安全漏洞名称", ""),
                    vuln.get("关联资产/域名", ""),
                    severity
                ]
                result_data.append(row_data)
        return result_data
    
    def _process_host_detail_sheet(self, file_path):
        """处理主机详情子表，获取主机信息"""
        return self.host_detail_processor.process_host_detail_sheet(file_path)
    
    def _count_vulnerabilities_by_ip(self, results):
        """按IP地址统计不同严重程度的漏洞数量"""
        return self.host_detail_processor.count_vulnerabilities_by_ip(results)
    
    def _generate_host_vuln_stat_report(self, file_path, hosts, vuln_counts, ip_device_map, output_path):
        """生成主机漏洞统计报告"""
        return self.report_generator.generate_host_vuln_stat_report(file_path, hosts, vuln_counts, ip_device_map, output_path)
    
    def replace_ip_with_device(self, input_file, output_file, ip_device_map):
        """替换Excel文件中的IP为设备名称，支持.xlsx和.xls格式"""
        return self.ip_replacer.replace_ip_with_device(input_file, output_file, ip_device_map)
    
    def merge_and_save_results(self, file_path, sheet_names, results, output_path):
        """合并多个工作表的处理结果并保存"""
        return self.report_generator.merge_and_save_results(file_path, sheet_names, results, output_path)
    
    def process_file(self):
        file_path = self.file_path.get()
        output_path = self.output_path.get()
        mapping_file_path = self.mapping_file_path.get()
        
        if not file_path:
            messagebox.showwarning("警告", "请选择Excel文件")
            return
        
        if not output_path:
            messagebox.showwarning("警告", "请选择输出路径")
            return
        
        try:
            self.log("开始处理文件...")
            
            # 获取所有工作表
            all_sheets = self.get_sheets(file_path)
            
            # 获取用户选择的工作表索引
            selected_indices = self.sheet_listbox.curselection()
            
            if not selected_indices:
                messagebox.showwarning("警告", "请选择要处理的工作表")
                return
            
            # 获取选择的工作表名称
            selected_sheets = [all_sheets[i] for i in selected_indices]
            self.log(f"已选择 {len(selected_sheets)} 个工作表: {selected_sheets}")
            
            # 处理每个选择的工作表
            results = []
            for sheet in selected_sheets:
                self.log(f"开始处理{sheet}工作表...")
                result = self.process_single_sheet(file_path, sheet)
                results.append(result)
                self.log(f"成功处理{sheet}，生成{len(result)}行数据")
            
            # 合并结果并保存
            output_file = self.merge_and_save_results(file_path, selected_sheets, results, output_path)
            
            # 处理主机详情子表，生成漏洞统计报告
            hosts = self._process_host_detail_sheet(file_path)
            
            # 统计漏洞数量
            vuln_counts = self._count_vulnerabilities_by_ip(results)
            
            # 读取IP设备映射表（如果有）
            ip_device_map = {}
            if mapping_file_path:
                ip_device_map = self.read_ip_device_mapping(mapping_file_path)
            
            # 生成主机漏洞统计报告
            stat_output_file = self._generate_host_vuln_stat_report(file_path, hosts, vuln_counts, ip_device_map, output_path)
            
            # 检查是否需要替换IP为设备名称
            replaced_output_file = None
            if mapping_file_path and ip_device_map:
                # 生成替换后的输出文件名称
                replaced_output_file = os.path.join(output_path, f"替换IP后_{os.path.basename(output_file)}")
                # 执行IP替换
                replaced_output_file = self.replace_ip_with_device(output_file, replaced_output_file, ip_device_map)
            
            self.log(f"所有工作表处理完成！")
            
            # 显示成功信息
            success_msg = f"处理完成！\n合并结果保存至: {output_file}\n主机漏洞统计报告保存至: {stat_output_file}"
            if replaced_output_file:
                success_msg += f"\n替换IP后结果保存至: {replaced_output_file}"
            
            messagebox.showinfo("成功", success_msg)
        except Exception as e:
            messagebox.showerror("错误", f"处理失败: {str(e)}")
            self.log(f"处理失败: {str(e)}")
    
    def log(self, message):
        self.log_text.insert(tk.END, f"{message}\n")
        self.log_text.see(tk.END)

if __name__ == "__main__":
    root = tk.Tk()
    app = WPSExcelTool(root)
    root.mainloop()
