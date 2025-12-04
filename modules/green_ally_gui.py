import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import queue

class GreenAllyGUI:
    """绿盟漏扫文件处理GUI，提供可交互的用户界面"""
    
    def __init__(self, root, green_ally_processor, logger=None):
        """初始化绿盟漏扫文件处理GUI
        
        Args:
            root (tk.Tk): 主窗口
            green_ally_processor (GreenAllyProcessor): 绿盟漏扫文件处理器
            logger (callable, optional): 日志记录函数，默认None
        """
        self.root = root
        self.green_ally_processor = green_ally_processor
        self.logger = logger
        
        # 创建一个队列用于线程间通信
        self.queue = queue.Queue()
        
        # 初始化变量
        self.selected_files = []
        self.processed_vulnerabilities = []
        self.current_vulnerability_index = -1
        self.mapping_file = ""
        self.output_path = ""
        
        # 创建主页面
        self.create_main_page()
    
    def log(self, message):
        """记录日志
        
        Args:
            message (str): 日志消息
        """
        if self.logger:
            self.logger(message)
    
    def create_main_page(self):
        """创建主页面"""
        # 清空现有组件
        for widget in self.root.winfo_children():
            widget.destroy()
        
        # 设置窗口标题
        self.root.title("绿盟漏扫文件处理")
        
        # 创建主框架
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 创建标题
        title_label = ttk.Label(main_frame, text="绿盟漏扫文件处理系统", font=("Arial", 18, "bold"))
        title_label.pack(pady=20)
        
        # 创建功能按钮框架
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(pady=20, fill=tk.X)
        
        # 添加文件按钮
        self.add_file_button = ttk.Button(button_frame, text="添加漏扫文件", command=self.add_files)
        self.add_file_button.pack(side=tk.LEFT, padx=10)
        
        # 移除文件按钮
        self.remove_file_button = ttk.Button(button_frame, text="移除选中文件", command=self.remove_selected_file)
        self.remove_file_button.pack(side=tk.LEFT, padx=10)
        
        # 清空文件按钮
        self.clear_files_button = ttk.Button(button_frame, text="清空文件列表", command=self.clear_files)
        self.clear_files_button.pack(side=tk.LEFT, padx=10)
        
        # 处理文件按钮
        self.process_button = ttk.Button(button_frame, text="处理选中文件", command=self.process_files, state=tk.DISABLED)
        self.process_button.pack(side=tk.LEFT, padx=10)
        
        # 返回主界面按钮
        self.back_button = ttk.Button(button_frame, text="返回主界面", command=self.back_to_main)
        self.back_button.pack(side=tk.RIGHT, padx=10)
        
        # 创建文件列表框架
        file_list_frame = ttk.LabelFrame(main_frame, text="已选择的漏扫文件", padding="10")
        file_list_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        # 文件列表
        self.file_list = tk.Listbox(file_list_frame, selectmode=tk.SINGLE, height=10)
        self.file_list.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5)
        
        # 文件列表滚动条
        file_scrollbar = ttk.Scrollbar(file_list_frame, orient=tk.VERTICAL, command=self.file_list.yview)
        file_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.file_list.config(yscrollcommand=file_scrollbar.set)
        
        # 创建配置框架
        config_frame = ttk.LabelFrame(main_frame, text="处理配置", padding="10")
        config_frame.pack(fill=tk.X, pady=10)
        
        # 映射文件选择
        mapping_frame = ttk.Frame(config_frame)
        mapping_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(mapping_frame, text="映射文件:", width=10).pack(side=tk.LEFT, padx=5)
        self.mapping_file_var = tk.StringVar()
        self.mapping_file_var.set("未选择映射文件")
        ttk.Label(mapping_frame, textvariable=self.mapping_file_var, width=40, anchor=tk.W).pack(side=tk.LEFT, padx=5)
        self.select_mapping_button = ttk.Button(mapping_frame, text="选择映射文件", command=self.select_mapping_file, state=tk.DISABLED)
        self.select_mapping_button.pack(side=tk.RIGHT, padx=5)
        
        # 输出路径配置
        output_frame = ttk.Frame(config_frame)
        output_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(output_frame, text="输出路径:", width=10).pack(side=tk.LEFT, padx=5)
        self.output_path_var = tk.StringVar()
        self.output_path_var.set("未选择输出路径")
        ttk.Label(output_frame, textvariable=self.output_path_var, width=40, anchor=tk.W).pack(side=tk.LEFT, padx=5)
        self.select_output_button = ttk.Button(output_frame, text="选择输出路径", command=self.select_output_path, state=tk.DISABLED)
        self.select_output_button.pack(side=tk.RIGHT, padx=5)
        
        # 创建处理状态框架
        status_frame = ttk.LabelFrame(main_frame, text="处理状态", padding="10")
        status_frame.pack(fill=tk.X, pady=10)
        
        # 状态文本
        self.status_text = tk.Text(status_frame, height=5, wrap=tk.WORD, state=tk.DISABLED)
        self.status_text.pack(fill=tk.X, padx=5)
        
        # 状态滚动条
        status_scrollbar = ttk.Scrollbar(status_frame, orient=tk.VERTICAL, command=self.status_text.yview)
        status_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.status_text.config(yscrollcommand=status_scrollbar.set)
        
        # 创建处理结果框架
        result_frame = ttk.LabelFrame(main_frame, text="处理结果", padding="10")
        result_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        # 结果按钮框架
        result_button_frame = ttk.Frame(result_frame)
        result_button_frame.pack(fill=tk.X, pady=10)
        
        # 查看漏洞列表按钮
        self.view_vuln_button = ttk.Button(result_button_frame, text="查看漏洞列表", command=self.view_vulnerabilities, state=tk.DISABLED)
        self.view_vuln_button.pack(side=tk.LEFT, padx=10)
        
        # 导出结果按钮
        self.export_button = ttk.Button(result_button_frame, text="导出结果", command=self.export_results, state=tk.DISABLED)
        self.export_button.pack(side=tk.LEFT, padx=10)
        
        # 导出统计按钮
        self.export_stat_button = ttk.Button(result_button_frame, text="导出统计报告", command=self.export_statistics, state=tk.DISABLED)
        self.export_stat_button.pack(side=tk.LEFT, padx=10)
        
        # 处理结果文本
        self.result_text = tk.Text(result_frame, height=10, wrap=tk.WORD, state=tk.DISABLED)
        self.result_text.pack(fill=tk.BOTH, expand=True, padx=5)
        
        # 结果滚动条
        result_scrollbar = ttk.Scrollbar(result_frame, orient=tk.VERTICAL, command=self.result_text.yview)
        result_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.result_text.config(yscrollcommand=result_scrollbar.set)
        
        # 启动队列处理
        self.process_queue()
        
        # 更新按钮状态
        self.update_button_states()
    
    def add_files(self):
        """添加漏扫文件"""
        files = filedialog.askopenfilenames(
            filetypes=[("Excel Files", "*.xls")],
            title="选择绿盟漏扫文件"
        )
        
        if files:
            # 添加新文件到列表
            for file in files:
                if file not in self.selected_files:
                    self.selected_files.append(file)
                    self.file_list.insert(tk.END, os.path.basename(file))
            
            self.log(f"添加了 {len(files)} 个文件")
            self.update_status(f"已添加 {len(files)} 个文件到列表")
            
            # 更新按钮状态，启用映射文件选择
            self.update_button_states()
            
            # 自动弹出映射文件选择对话框
            self.select_mapping_file()
    
    def update_button_states(self):
        """根据当前状态更新按钮状态"""
        # 只有当有文件被选择时，才能选择映射文件
        if self.selected_files:
            self.select_mapping_button.config(state=tk.NORMAL)
        else:
            self.select_mapping_button.config(state=tk.DISABLED)
            self.select_output_button.config(state=tk.DISABLED)
            self.process_button.config(state=tk.DISABLED)
        
        # 只有当选择了映射文件后，才能选择输出路径
        if self.mapping_file:
            self.select_output_button.config(state=tk.NORMAL)
        else:
            self.select_output_button.config(state=tk.DISABLED)
            self.process_button.config(state=tk.DISABLED)
        
        # 只有当选择了输出路径后，才能处理文件
        if self.output_path:
            self.process_button.config(state=tk.NORMAL)
        else:
            self.process_button.config(state=tk.DISABLED)
    
    def select_mapping_file(self):
        """选择映射文件"""
        mapping_file = filedialog.askopenfilename(
            filetypes=[("Word Files", "*.docx *.doc")],
            title="选择映射文件"
        )
        
        if mapping_file:
            self.mapping_file = mapping_file
            self.mapping_file_var.set(os.path.basename(mapping_file))
            self.update_status(f"已选择映射文件: {os.path.basename(mapping_file)}")
            
            # 更新按钮状态，启用输出路径选择
            self.update_button_states()
            
            # 自动弹出输出路径选择对话框
            self.select_output_path()
    
    def select_output_path(self):
        """选择输出路径"""
        output_path = filedialog.askdirectory(title="选择输出路径")
        
        if output_path:
            self.output_path = output_path
            self.output_path_var.set(output_path)
            self.update_status(f"已选择输出路径: {output_path}")
            
            # 更新按钮状态，启用处理按钮
            self.update_button_states()
    
    def remove_selected_file(self):
        """移除选中的文件"""
        selected_index = self.file_list.curselection()
        if selected_index:
            index = selected_index[0]
            self.file_list.delete(index)
            del self.selected_files[index]
            self.update_status(f"已移除选中文件")
            
            # 如果文件列表为空，重置映射文件和输出路径
            if not self.selected_files:
                self.mapping_file = ""
                self.mapping_file_var.set("未选择映射文件")
                self.output_path = ""
                self.output_path_var.set("未选择输出路径")
                self.update_status("已重置映射文件和输出路径设置")
            
            # 更新按钮状态
            self.update_button_states()
    
    def clear_files(self):
        """清空文件列表"""
        self.file_list.delete(0, tk.END)
        self.selected_files.clear()
        self.update_status("已清空文件列表")
        
        # 重置映射文件和输出路径
        self.mapping_file = ""
        self.mapping_file_var.set("未选择映射文件")
        self.output_path = ""
        self.output_path_var.set("未选择输出路径")
        self.update_status("已重置映射文件和输出路径设置")
        
        # 更新按钮状态
        self.update_button_states()
    
    def process_files(self):
        """处理选中的文件"""
        if not self.selected_files:
            messagebox.showwarning("警告", "请先添加漏扫文件")
            return
        
        if not self.mapping_file:
            messagebox.showwarning("警告", "请先选择映射文件")
            return
        
        if not self.output_path:
            messagebox.showwarning("警告", "请先选择输出路径")
            return
        
        # 禁用按钮
        self.add_file_button.config(state=tk.DISABLED)
        self.remove_file_button.config(state=tk.DISABLED)
        self.clear_files_button.config(state=tk.DISABLED)
        self.process_button.config(state=tk.DISABLED)
        self.select_mapping_button.config(state=tk.DISABLED)
        self.select_output_button.config(state=tk.DISABLED)
        
        # 更新状态
        self.update_status("开始处理文件...")
        
        # 在新线程中处理文件，避免GUI卡顿
        threading.Thread(target=self._process_files_thread, daemon=True).start()
    
    def _process_files_thread(self):
        """处理文件的线程函数"""
        try:
            # 重置结果
            self.processed_vulnerabilities.clear()
            self.current_vulnerability_index = -1
            
            # 处理每个文件
            for i, file_path in enumerate(self.selected_files):
                self.queue.put(f"正在处理文件 {i+1}/{len(self.selected_files)}: {os.path.basename(file_path)}")
                
                # 处理文件
                vulnerabilities = self.green_ally_processor.process_green_ally_report(file_path)
                self.processed_vulnerabilities.extend(vulnerabilities)
                
                self.queue.put(f"成功处理文件: {os.path.basename(file_path)}，提取 {len(vulnerabilities)} 个漏洞")
            
            # 生成合并结果和统计报告
            self.queue.put("正在生成合并结果和统计报告...")
            
            # 使用当前时间作为文件名的一部分
            import datetime
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            base_name = f"green_ally_result_{timestamp}"
            
            # 生成合并结果文件
            merged_file = self.green_ally_processor._generate_merged_result(
                self.processed_vulnerabilities,
                self.output_path,
                base_name
            )
            
            # 生成统计报告文件
            ip_device_map = None
            if self.mapping_file:
                # 如果有映射表，先读取它
                from .ip_device_mapper import IPDeviceMapper
                ip_mapper = IPDeviceMapper(logger=self.log)
                ip_device_map = ip_mapper.read_ip_device_mapping(self.mapping_file)
            
            # 生成统计报告，传递IP映射表
            stat_file = self.green_ally_processor._generate_vuln_stat_report(
                self.processed_vulnerabilities,
                self.output_path,
                base_name,
                ip_device_map
            )
            
            # 生成按漏洞名称合并的结果文件
            self.queue.put("正在生成按漏洞名称合并的结果...")
            vuln_merged_file = self.green_ally_processor._generate_vulnerability_merged_result(
                self.processed_vulnerabilities,
                self.output_path,
                base_name
            )
            
            # 生成替换IP后的结果（如果提供了映射表）
            replaced_file = None
            replaced_stat_file = None
            replaced_vuln_merged_file = None
            if self.mapping_file:
                self.queue.put("正在生成替换IP后的结果...")
                
                from .ip_device_mapper import IPDeviceMapper
                from .ip_replacer import IPReplacer
                
                # 创建IP设备映射器
                ip_mapper = IPDeviceMapper(logger=self.log)
                # 读取IP设备映射表
                ip_device_map = ip_mapper.read_ip_device_mapping(self.mapping_file)
                
                # 创建IP替换器
                ip_replacer = IPReplacer(logger=self.log)
                
                # 替换合并结果文件中的IP地址（IP_COLUMN_INDEX=2）
                replaced_file = os.path.join(self.output_path, f"替换IP后_合并处理结果_{base_name}.xlsx")
                ip_replacer.replace_ip_with_device(merged_file, replaced_file, ip_device_map)
                
                # 替换统计报告文件中的IP地址（设备名称或IP地址列，索引为1）
                replaced_stat_file = os.path.join(self.output_path, f"替换IP后_主机漏洞统计_{base_name}.xlsx")
                ip_replacer.replace_ip_with_device(stat_file, replaced_stat_file, ip_device_map, ip_column_index=1)
                
                # 替换按漏洞名称合并结果文件中的IP地址（IP_COLUMN_INDEX=2）
                replaced_vuln_merged_file = os.path.join(self.output_path, f"替换IP后_按漏洞名称合并结果_{base_name}.xlsx")
                ip_replacer.replace_ip_with_device(vuln_merged_file, replaced_vuln_merged_file, ip_device_map)
            
            # 更新结果
            self.queue.put(f"所有文件处理完成，共提取 {len(self.processed_vulnerabilities)} 个漏洞")
            self.queue.put(f"合并结果已保存到: {merged_file}")
            self.queue.put(f"统计报告已保存到: {stat_file}")
            self.queue.put(f"按漏洞名称合并结果已保存到: {vuln_merged_file}")
            if replaced_file:
                self.queue.put(f"替换IP后的结果已保存到: {replaced_file}")
            if replaced_stat_file:
                self.queue.put(f"替换IP后的统计报告已保存到: {replaced_stat_file}")
            if replaced_vuln_merged_file:
                self.queue.put(f"替换IP后的按漏洞名称合并结果已保存到: {replaced_vuln_merged_file}")
            self.queue.put("update_result_buttons")
            
            # 显示处理结果
            result_message = f"文件处理完成！\n合并结果: {os.path.basename(merged_file)}\n统计报告: {os.path.basename(stat_file)}\n按漏洞名称合并结果: {os.path.basename(vuln_merged_file)}"
            if replaced_file:
                result_message += f"\n替换IP后的结果: {os.path.basename(replaced_file)}"
            if replaced_stat_file:
                result_message += f"\n替换IP后的统计报告: {os.path.basename(replaced_stat_file)}"
            if replaced_vuln_merged_file:
                result_message += f"\n替换IP后的按漏洞名称合并结果: {os.path.basename(replaced_vuln_merged_file)}"
            messagebox.showinfo("成功", result_message)
        except Exception as e:
            self.queue.put(f"处理文件失败: {str(e)}")
            messagebox.showerror("错误", f"处理文件失败: {str(e)}")
        finally:
            # 启用按钮
            self.queue.put("enable_buttons")
    
    def view_vulnerabilities(self):
        """查看漏洞列表"""
        if not self.processed_vulnerabilities:
            messagebox.showwarning("警告", "请先处理漏扫文件")
            return
        
        # 创建漏洞列表页面
        self.create_vulnerability_list_page()
    
    def create_vulnerability_list_page(self):
        """创建漏洞列表页面"""
        # 清空现有组件
        for widget in self.root.winfo_children():
            widget.destroy()
        
        # 设置窗口标题
        self.root.title("漏洞列表")
        
        # 创建主框架
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 创建标题
        title_label = ttk.Label(main_frame, text="漏洞列表", font=("Arial", 16, "bold"))
        title_label.pack(pady=10)
        
        # 创建返回按钮
        back_button = ttk.Button(main_frame, text="返回主页面", command=self.create_main_page)
        back_button.pack(anchor=tk.NE, pady=10)
        
        # 创建漏洞统计信息
        stat_frame = ttk.Frame(main_frame)
        stat_frame.pack(fill=tk.X, pady=10)
        
        total_vuln_label = ttk.Label(stat_frame, text=f"共发现 {len(self.processed_vulnerabilities)} 个漏洞")
        total_vuln_label.pack(side=tk.LEFT, padx=10)
        
        # 统计不同严重程度的漏洞数量
        severity_counts = self._count_vulnerabilities_by_severity()
        severity_text = f"高: {severity_counts['高']} 个，中: {severity_counts['中']} 个，低: {severity_counts['低']} 个"
        severity_label = ttk.Label(stat_frame, text=severity_text)
        severity_label.pack(side=tk.LEFT, padx=10)
        
        # 创建漏洞列表框架
        vuln_list_frame = ttk.LabelFrame(main_frame, text="漏洞列表", padding="10")
        vuln_list_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        # 创建漏洞列表
        self.vuln_listbox = tk.Listbox(
            vuln_list_frame,
            selectmode=tk.SINGLE,
            height=15
        )
        self.vuln_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5)
        
        # 绑定列表选择事件
        self.vuln_listbox.bind("<<ListboxSelect>>", self.on_vuln_select)
        
        # 添加滚动条
        vuln_scrollbar = ttk.Scrollbar(vuln_list_frame, orient=tk.VERTICAL, command=self.vuln_listbox.yview)
        vuln_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.vuln_listbox.config(yscrollcommand=vuln_scrollbar.set)
        
        # 填充漏洞列表
        for i, vuln in enumerate(self.processed_vulnerabilities):
            # 格式：序号 - 漏洞名称 (严重程度) - 关联资产
            vuln_text = f"{i+1} - {vuln[1]} ({vuln[3]}) - {vuln[2]}"
            self.vuln_listbox.insert(tk.END, vuln_text)
        
        # 创建漏洞详情框架
        vuln_detail_frame = ttk.LabelFrame(main_frame, text="漏洞详情", padding="10")
        vuln_detail_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        # 漏洞详情文本
        self.vuln_detail_text = tk.Text(vuln_detail_frame, height=10, wrap=tk.WORD, state=tk.DISABLED)
        self.vuln_detail_text.pack(fill=tk.BOTH, expand=True, padx=5)
        
        # 详情滚动条
        detail_scrollbar = ttk.Scrollbar(vuln_detail_frame, orient=tk.VERTICAL, command=self.vuln_detail_text.yview)
        detail_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.vuln_detail_text.config(yscrollcommand=detail_scrollbar.set)
    
    def on_vuln_select(self, event):
        """当选择漏洞列表项时触发"""
        selected_index = self.vuln_listbox.curselection()
        if selected_index:
            index = selected_index[0]
            vuln = self.processed_vulnerabilities[index]
            self.display_vuln_detail(vuln)
    
    def display_vuln_detail(self, vuln):
        """显示漏洞详情
        
        Args:
            vuln (list): 漏洞数据
        """
        # 启用文本控件
        self.vuln_detail_text.config(state=tk.NORMAL)
        self.vuln_detail_text.delete(1.0, tk.END)
        
        # 显示漏洞详情
        detail_text = f"序号: {vuln[0]}\n"
        detail_text += f"安全漏洞名称: {vuln[1]}\n"
        detail_text += f"关联资产/域名: {vuln[2]}\n"
        detail_text += f"严重程度: {vuln[3]}\n"
        
        self.vuln_detail_text.insert(tk.END, detail_text)
        self.vuln_detail_text.config(state=tk.DISABLED)
    
    def _count_vulnerabilities_by_severity(self):
        """按严重程度统计漏洞数量
        
        Returns:
            dict: 漏洞统计字典
        """
        counts = {"高": 0, "中": 0, "低": 0}
        
        for vuln in self.processed_vulnerabilities:
            severity = vuln[3]
            if severity in counts:
                counts[severity] += 1
        
        return counts
    
    def export_results(self):
        """导出结果"""
        if not self.processed_vulnerabilities:
            messagebox.showwarning("警告", "请先处理漏扫文件")
            return
        
        # 选择输出文件路径
        output_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")],
            title="选择导出文件路径"
        )
        
        if output_path:
            try:
                # 生成合并处理结果
                base_name = os.path.basename(output_path)
                name_without_ext = os.path.splitext(base_name)[0]
                output_dir = os.path.dirname(output_path)
                
                merged_file = self.green_ally_processor._generate_merged_result(
                    self.processed_vulnerabilities,
                    output_dir,
                    name_without_ext
                )
                
                messagebox.showinfo("成功", f"结果已导出到: {merged_file}")
                self.update_status(f"结果已导出到: {merged_file}")
            except Exception as e:
                messagebox.showerror("错误", f"导出结果失败: {str(e)}")
                self.update_status(f"导出结果失败: {str(e)}")
    
    def export_statistics(self):
        """导出统计报告"""
        if not self.processed_vulnerabilities:
            messagebox.showwarning("警告", "请先处理漏扫文件")
            return
        
        # 选择输出文件路径
        output_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")],
            title="选择导出统计报告路径"
        )
        
        if output_path:
            try:
                # 生成主机漏洞统计报告
                base_name = os.path.basename(output_path)
                name_without_ext = os.path.splitext(base_name)[0]
                output_dir = os.path.dirname(output_path)
                
                stat_file = self.green_ally_processor._generate_vuln_stat_report(
                    self.processed_vulnerabilities,
                    output_dir,
                    name_without_ext
                )
                
                messagebox.showinfo("成功", f"统计报告已导出到: {stat_file}")
                self.update_status(f"统计报告已导出到: {stat_file}")
            except Exception as e:
                messagebox.showerror("错误", f"导出统计报告失败: {str(e)}")
                self.update_status(f"导出统计报告失败: {str(e)}")
    
    def update_status(self, message):
        """更新状态文本
        
        Args:
            message (str): 状态消息
        """
        self.queue.put(message)
    
    def process_queue(self):
        """处理队列消息，更新GUI"""
        try:
            while not self.queue.empty():
                message = self.queue.get_nowait()
                
                if message == "enable_buttons":
                    # 启用按钮
                    self.add_file_button.config(state=tk.NORMAL)
                    self.remove_file_button.config(state=tk.NORMAL)
                    self.clear_files_button.config(state=tk.NORMAL)
                    self.process_button.config(state=tk.NORMAL)
                elif message == "update_result_buttons":
                    # 更新结果按钮
                    self.view_vuln_button.config(state=tk.NORMAL)
                    self.export_button.config(state=tk.NORMAL)
                    self.export_stat_button.config(state=tk.NORMAL)
                else:
                    # 更新状态文本
                    self.status_text.config(state=tk.NORMAL)
                    self.status_text.insert(tk.END, f"{message}\n")
                    self.status_text.see(tk.END)
                    self.status_text.config(state=tk.DISABLED)
                    
                    # 更新结果文本
                    self.result_text.config(state=tk.NORMAL)
                    self.result_text.insert(tk.END, f"{message}\n")
                    self.result_text.see(tk.END)
                    self.result_text.config(state=tk.DISABLED)
        except queue.Empty:
            pass
        finally:
            # 继续处理队列
            self.root.after(100, self.process_queue)
    
    def back_to_main(self):
        """返回主界面"""
        # 这里需要实现返回主界面的逻辑
        # 由于我们没有主界面的引用，暂时只关闭窗口
        self.root.destroy()

# 添加os导入，因为在代码中使用了os.path
import os