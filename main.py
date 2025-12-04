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

# 从modules目录导入所需组件
from modules.file_reader import FileReader
from modules.vulnerability_extractor import VulnerabilityExtractor
from modules.ip_device_mapper import IPDeviceMapper
from modules.ip_replacer import IPReplacer
from modules.report_generator import ReportGenerator
from modules.host_detail_processor import HostDetailProcessor
from modules.shengbang_processor import ShengBangProcessor
from modules.green_ally_processor import GreenAllyProcessor
from modules.green_ally_gui import GreenAllyGUI

class WPSExcelTool:
    """WPS Excel 处理工具主类，负责整个应用的GUI和业务逻辑协调"""
    
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
        self.shengbang_processor = ShengBangProcessor(logger=self.log)
        self.green_ally_processor = GreenAllyProcessor(logger=self.log)
        
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
        ttk.Button(button_frame, text="处理绿盟漏扫文件", command=self.open_green_ally_gui).pack(side=tk.LEFT, padx=10)
        ttk.Button(button_frame, text="批量处理绿盟漏扫报告", command=self.batch_process_green_ally_reports).pack(side=tk.LEFT, padx=10)
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
    
    def process_single_sheet(self, file_path, sheet_name):
        """处理单个工作表或文档，返回处理结果数据，支持xlsx、xls、docx和doc格式"""
        try:
            # 检查是否是盛邦漏洞扫描报告
            is_shengbang_report = self.shengbang_processor.is_shengbang_report(file_path)
            
            if is_shengbang_report:
                # 使用盛邦漏洞扫描报告处理器处理
                return self.shengbang_processor.process_report(file_path, sheet_name)
            
            # 检查是否是绿盟漏扫报告
            is_green_ally_report = self.green_ally_processor.is_green_ally_report(file_path)
            
            if is_green_ally_report:
                # 使用绿盟漏扫报告处理器处理
                return self.green_ally_processor.process_green_ally_report(file_path)
            
            # 使用现有逻辑处理
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
    
    def batch_process_green_ally_reports(self):
        """批量处理绿盟漏扫报告"""
        # 选择文件夹
        folder_path = filedialog.askdirectory()
        if not folder_path:
            self.log("未选择文件夹")
            return
        
        # 选择输出路径
        output_path = filedialog.askdirectory()
        if not output_path:
            self.log("未选择输出路径")
            return
        
        try:
            self.log(f"开始批量处理绿盟漏扫报告，文件夹: {folder_path}")
            
            # 调用绿盟漏扫报告处理器的批量处理方法
            result = self.green_ally_processor.batch_process_folder(
                folder_path,
                output_path,
                self.mapping_file_path.get() if self.mapping_file_path.get() else None
            )
            
            # 显示处理结果
            self.log(f"批量处理完成")
            self.log(f"成功处理: {len(result['success'])} 个文件")
            for file in result['success']:
                self.log(f"  成功: {file}")
            
            if result['failed']:
                self.log(f"处理失败: {len(result['failed'])} 个文件")
                for item in result['failed']:
                    self.log(f"  失败: {item['file']} - {item['reason']}")
            
            messagebox.showinfo("成功", f"批量处理完成\n成功: {len(result['success'])} 个文件\n失败: {len(result['failed'])} 个文件")
        except Exception as e:
            self.log(f"批量处理失败: {str(e)}")
            messagebox.showerror("错误", f"批量处理失败: {str(e)}")
    
    def log(self, message):
        self.log_text.insert(tk.END, f"{message}\n")
        self.log_text.see(tk.END)
    
    def open_green_ally_gui(self):
        """打开绿盟漏扫文件处理GUI"""
        try:
            # 创建一个新窗口来显示绿盟漏扫GUI
            green_ally_window = tk.Toplevel(self.root)
            green_ally_window.title("绿盟漏扫文件处理")
            green_ally_window.geometry("800x600")
            
            # 初始化绿盟漏扫GUI
            self.green_ally_gui = GreenAllyGUI(green_ally_window, self.green_ally_processor, self.log)
            
            self.log("打开绿盟漏扫文件处理GUI成功")
        except Exception as e:
            self.log(f"打开绿盟漏扫文件处理GUI失败: {str(e)}")
            messagebox.showerror("错误", f"打开绿盟漏扫文件处理GUI失败: {str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = WPSExcelTool(root)
    root.mainloop()