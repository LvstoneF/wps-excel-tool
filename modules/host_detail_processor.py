import re
from constants import IP_PATTERN, HOST_DETAIL_SHEET_KEYWORDS, VULN_SHEET_PREFIX
from .file_reader import FileReader

class HostDetailProcessor:
    """主机详情处理器，负责处理主机详情子表"""
    
    def __init__(self, logger=None):
        """初始化主机详情处理器
        
        Args:
            logger (callable, optional): 日志记录函数，默认None
        """
        self.logger = logger
    
    def log(self, message):
        """记录日志
        
        Args:
            message (str): 日志消息
        """
        if self.logger:
            self.logger(message)
    
    def process_host_detail_sheet(self, file_path, sheet_name):
        """处理主机详情子表，返回处理结果数据
        
        Args:
            file_path (str): 文件路径
            sheet_name (str): 工作表名称
        
        Returns:
            list: 处理结果数据
        
        Raises:
            Exception: 处理工作表失败时抛出异常
        """
        try:
            # 读取主机详情子表数据
            rows = FileReader.read_file_rows(file_path, sheet_name)
            
            # 提取主机信息
            hosts = self._extract_hosts_from_rows(rows)
            
            return hosts
        except Exception as e:
            raise Exception(f"处理工作表{sheet_name}失败: {str(e)}")
    
    def _extract_hosts_from_rows(self, rows):
        """从行数据中提取主机信息
        
        Args:
            rows (list): 行数据列表
        
        Returns:
            list: 主机信息列表
        """
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
                    continue
            
            # 提取主机信息
            if header:
                host = {}
                for j, cell in enumerate(row):
                    if j < len(header):
                        host[header[j]] = str(cell).strip() if cell else ""
                hosts.append(host)
        
        return hosts
    
    def count_vulnerabilities_by_ip(self, results):
        """按IP地址统计不同严重程度的漏洞数量
        
        Args:
            results (list): 处理结果数据列表
        
        Returns:
            dict: IP漏洞统计字典
        """
        vuln_counts = {}
        
        for result in results:
            for row in result:
                if len(row) >= 4:
                    ip_text = row[2]
                    severity = row[3]
                    if ip_text:
                        # 从文本中提取所有IP地址
                        import re
                        from constants import IP_PATTERN
                        ips = re.findall(IP_PATTERN, ip_text)
                        
                        for ip in ips:
                            if ip:
                                if ip not in vuln_counts:
                                    vuln_counts[ip] = {
                                        "高": 0,
                                        "中": 0,
                                        "低": 0
                                    }
                                # 只统计高、中、低三个级别
                                if severity in ["高", "中", "低"]:
                                    vuln_counts[ip][severity] += 1
        
        return vuln_counts
    
    def is_vulnerability_sheet(self, sheet_name):
        """判断是否是漏洞详情工作表
        
        Args:
            sheet_name (str): 工作表名称
        
        Returns:
            bool: 是否是漏洞详情工作表
        """
        from constants import VULN_SHEET_KEYWORDS, VULN_SHEET_PREFIX
        return sheet_name in VULN_SHEET_KEYWORDS or sheet_name.startswith(VULN_SHEET_PREFIX)
    
    def get_sheets(self, file_path):
        """获取文件的工作表列表
        
        Args:
            file_path (str): 文件路径
        
        Returns:
            list: 工作表名称列表
        
        Raises:
            Exception: 读取工作表失败时抛出异常
        """
        return FileReader.get_sheets(file_path)
    
    def process_file(self, file_path):
        """处理文件，返回处理结果数据
        
        Args:
            file_path (str): 文件路径
        
        Returns:
            list: 处理结果数据
        
        Raises:
            Exception: 处理文件失败时抛出异常
        """
        sheets = self.get_sheets(file_path)
        results = []
        for sheet in sheets:
            if self.is_vulnerability_sheet(sheet):
                result = self.process_single_sheet(file_path, sheet)
                results.append(result)
        return results
