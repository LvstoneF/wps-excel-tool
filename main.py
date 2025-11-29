import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import openpyxl
import os

class WPSExcelTool:
    def __init__(self, root):
        self.root = root
        self.root.title("WPS Excel 处理工具")
        self.root.geometry("600x400")
        
        # 初始化变量
        self.file_path = tk.StringVar()
        self.sheet_name = tk.StringVar()
        self.output_path = tk.StringVar()
        
        # 创建界面
        self.create_widgets()
    
    def create_widgets(self):
        # 标题
        title_label = ttk.Label(self.root, text="WPS Excel 处理工具", font=("Arial", 16))
        title_label.pack(pady=20)
        
        # 文件选择
        file_frame = ttk.Frame(self.root)
        file_frame.pack(pady=10, padx=20, fill=tk.X)
        
        ttk.Label(file_frame, text="选择Excel文件:").pack(side=tk.LEFT, padx=5)
        ttk.Entry(file_frame, textvariable=self.file_path, width=40).pack(side=tk.LEFT, padx=5)
        ttk.Button(file_frame, text="浏览", command=self.browse_file).pack(side=tk.LEFT, padx=5)
        
        # 工作表选择
        sheet_frame = ttk.Frame(self.root)
        sheet_frame.pack(pady=10, padx=20, fill=tk.X)
        
        ttk.Label(sheet_frame, text="工作表:").pack(side=tk.LEFT, padx=5)
        ttk.Combobox(sheet_frame, textvariable=self.sheet_name, width=38).pack(side=tk.LEFT, padx=5)
        ttk.Button(sheet_frame, text="刷新", command=self.refresh_sheets).pack(side=tk.LEFT, padx=5)
        
        # 输出路径
        output_frame = ttk.Frame(self.root)
        output_frame.pack(pady=10, padx=20, fill=tk.X)
        
        ttk.Label(output_frame, text="输出路径:").pack(side=tk.LEFT, padx=5)
        ttk.Entry(output_frame, textvariable=self.output_path, width=40).pack(side=tk.LEFT, padx=5)
        ttk.Button(output_frame, text="浏览", command=self.browse_output).pack(side=tk.LEFT, padx=5)
        
        # 处理按钮
        button_frame = ttk.Frame(self.root)
        button_frame.pack(pady=20)
        
        ttk.Button(button_frame, text="处理文档", command=self.process_file).pack(side=tk.LEFT, padx=10)
        ttk.Button(button_frame, text="退出", command=self.root.quit).pack(side=tk.LEFT, padx=10)
        
        # 日志区域
        log_frame = ttk.Frame(self.root)
        log_frame.pack(pady=10, padx=20, fill=tk.BOTH, expand=True)
        
        ttk.Label(log_frame, text="处理日志:").pack(anchor=tk.W, padx=5)
        self.log_text = tk.Text(log_frame, height=10, width=70)
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 滚动条
        scrollbar = ttk.Scrollbar(self.log_text, orient=tk.VERTICAL, command=self.log_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.config(yscrollcommand=scrollbar.set)
    
    def browse_file(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel Files", "*.xlsx;*.xls")]
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
    
    def refresh_sheets(self):
        file_path = self.file_path.get()
        if not file_path:
            messagebox.showwarning("警告", "请先选择Excel文件")
            return
        
        try:
            workbook = openpyxl.load_workbook(file_path)
            sheets = workbook.sheetnames
            workbook.close()
            
            # 更新下拉框
            combobox = self.root.nametowidget('.!frame2.!combobox')
            combobox['values'] = sheets
            if sheets:
                self.sheet_name.set(sheets[0])
            self.log(f"刷新工作表: {sheets}")
        except Exception as e:
            messagebox.showerror("错误", f"读取工作表失败: {str(e)}")
            self.log(f"读取工作表失败: {str(e)}")
    
    def process_file(self):
        file_path = self.file_path.get()
        sheet_name = self.sheet_name.get()
        output_path = self.output_path.get()
        
        if not file_path:
            messagebox.showwarning("警告", "请选择Excel文件")
            return
        
        if not sheet_name:
            messagebox.showwarning("警告", "请选择工作表")
            return
        
        if not output_path:
            messagebox.showwarning("警告", "请选择输出路径")
            return
        
        try:
            self.log("开始处理文件...")
            
            # 打开工作簿
            workbook = openpyxl.load_workbook(file_path)
            sheet = workbook[sheet_name]
            
            # 创建新工作簿和工作表
            new_workbook = openpyxl.Workbook()
            new_sheet = new_workbook.active
            new_sheet.title = "漏洞详情处理结果"
            
            # 如果是漏洞相关工作表，进行特殊处理
            if sheet_name in ["漏洞详情", "Sheet1"]:
                self.log(f"开始处理{sheet_name}工作表...")
                
                # 定义表头
                headers = ["序号", "漏洞标题", "漏洞编号", "漏洞类型", "危险级别", "影响平台", "CVSS分值", 
                          "bugtraq编号", "CVE编号", "CNCVE编号", "国家漏洞库编号", "CNNVD编号", 
                          "CNVD编号", "漏洞可利用性", "存在主机", "简单描述", "详细描述", "修补建议", "参考网址", "漏洞安全性"]
                new_sheet.append(headers)
                
                # 设置列宽
                for col in range(1, len(headers) + 1):
                    new_sheet.column_dimensions[openpyxl.utils.get_column_letter(col)].width = 20
                
                # 遍历原始工作表，提取漏洞信息
                vulnerabilities = []
                current_vuln = {}
                vuln_index = 0
                
                for row in sheet.iter_rows(min_row=1, values_only=True):
                    # 跳过空行
                    if not any(row):
                        continue
                    
                    # 检查是否是新漏洞标题行（以【数字】开头，在B列或A列）
                    title_cell = row[1] if len(row) > 1 else row[0]
                    if title_cell and isinstance(title_cell, str) and title_cell.startswith("【") and "】" in title_cell:
                        # 如果有当前漏洞，先保存
                        if current_vuln:
                            vulnerabilities.append(current_vuln)
                        # 开始新漏洞
                        current_vuln = {
                            "序号": title_cell.split("】")[0][1:],
                            "漏洞标题": title_cell.split("】")[1].strip()
                        }
                        vuln_index += 1
                    # 检查是否是属性行（B列或A列有属性名称）
                    elif len(row) >= 3 and row[1] and isinstance(row[1], str) and row[2]:
                        # 提取属性名称和值（B列是属性名，C列是属性值）
                        attr_name = row[1].strip()
                        attr_value = row[2].strip()
                        
                        # 映射属性名称到表头
                        attr_map = {
                            "漏洞编号": "漏洞编号",
                            "漏洞类型": "漏洞类型",
                            "危险级别": "危险级别",
                            "影响平台": "影响平台",
                            "CVSS分值": "CVSS分值",
                            "bugtraq编号": "bugtraq编号",
                            "CVE编号": "CVE编号",
                            "CNCVE编号": "CNCVE编号",
                            "国家漏洞库编号": "国家漏洞库编号",
                            "CNNVD编号": "CNNVD编号",
                            "CNVD编号": "CNVD编号",
                            "漏洞可利用性": "漏洞可利用性",
                            "存在主机": "存在主机",
                            "简单描述": "简单描述",
                            "详细描述": "详细描述",
                            "修补建议": "修补建议",
                            "参考网址": "参考网址",
                            "漏洞安全性": "漏洞安全性"
                        }
                        
                        if attr_name in attr_map:
                            current_vuln[attr_map[attr_name]] = attr_value
                    # 兼容旧格式：A列是属性名，B列是属性值
                    elif row[0] and isinstance(row[0], str) and len(row) > 1 and row[1]:
                        attr_name = row[0].strip()
                        attr_value = row[1].strip()
                        
                        attr_map = {
                            "漏洞编号": "漏洞编号",
                            "漏洞类型": "漏洞类型",
                            "危险级别": "危险级别",
                            "影响平台": "影响平台",
                            "CVSS分值": "CVSS分值",
                            "bugtraq编号": "bugtraq编号",
                            "CVE编号": "CVE编号",
                            "CNCVE编号": "CNCVE编号",
                            "国家漏洞库编号": "国家漏洞库编号",
                            "CNNVD编号": "CNNVD编号",
                            "CNVD编号": "CNVD编号",
                            "漏洞可利用性": "漏洞可利用性",
                            "存在主机": "存在主机",
                            "简单描述": "简单描述",
                            "详细描述": "详细描述",
                            "修补建议": "修补建议",
                            "参考网址": "参考网址",
                            "漏洞安全性": "漏洞安全性"
                        }
                        
                        if attr_name in attr_map:
                            current_vuln[attr_map[attr_name]] = attr_value
                
                # 保存最后一个漏洞
                if current_vuln:
                    vulnerabilities.append(current_vuln)
                
                # 将提取的漏洞信息写入新工作表
                for vuln in vulnerabilities:
                    # 按照表头顺序提取值
                    row_data = [vuln.get(header, "") for header in headers]
                    new_sheet.append(row_data)
                
                self.log(f"成功提取 {len(vulnerabilities)} 个漏洞信息")
            else:
                # 非漏洞详情工作表，执行默认处理（复制所有行）
                for row in sheet.iter_rows(values_only=True):
                    new_sheet.append(row)
                self.log("执行默认复制处理")
            
            # 保存新文件
            output_file = os.path.join(output_path, f"处理结果_{os.path.basename(file_path)}")
            new_workbook.save(output_file)
            
            workbook.close()
            new_workbook.close()
            
            self.log(f"处理完成！结果保存至: {output_file}")
            messagebox.showinfo("成功", f"处理完成！结果保存至: {output_file}")
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
