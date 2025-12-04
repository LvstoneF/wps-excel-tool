from constants import SEVERITY_MAP

class ShengBangProcessor:
    """盛邦漏洞扫描报告处理器，负责盛邦漏洞扫描报告的解析、存储和展示"""
    
    def __init__(self, logger=None):
        """初始化盛邦漏洞扫描报告处理器
        
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
    
    def process_report(self, file_path, sheet_name):
        """处理盛邦漏洞扫描报告
        
        Args:
            file_path (str): 文件路径
            sheet_name (str): 工作表名称
        
        Returns:
            list: 处理后的漏洞数据列表
        
        Raises:
            Exception: 处理报告失败时抛出异常
        """
        try:
            self.log(f"开始处理盛邦漏洞扫描报告: {file_path}")
            
            # 这里可以根据盛邦漏洞扫描报告的具体格式进行解析
            # 目前先使用现有的VulnerabilityExtractor来提取漏洞信息
            from .file_reader import FileReader
            from .vulnerability_extractor import VulnerabilityExtractor
            
            # 读取文件行数据
            rows = FileReader.read_file_rows(file_path, sheet_name)
            
            # 提取漏洞信息
            vulnerabilities = VulnerabilityExtractor.extract_vulnerabilities(rows)
            
            # 转换为列表格式，过滤掉严重程度为"信息"的条目
            result_data = VulnerabilityExtractor.convert_vulnerabilities_to_list(vulnerabilities)
            
            self.log(f"成功处理盛邦漏洞扫描报告，提取{len(result_data)}个漏洞")
            return result_data
        except Exception as e:
            self.log(f"处理盛邦漏洞扫描报告失败: {str(e)}")
            raise Exception(f"处理盛邦漏洞扫描报告失败: {str(e)}")
    
    def is_shengbang_report(self, file_path):
        """判断是否是盛邦漏洞扫描报告
        
        Args:
            file_path (str): 文件路径
        
        Returns:
            bool: 是否是盛邦漏洞扫描报告
        """
        try:
            from .file_reader import FileReader
            
            # 获取文件的工作表
            sheets = FileReader.get_sheets(file_path)
            
            # 盛邦漏洞扫描报告通常会有特定的工作表名称
            # 这里可以根据实际情况调整判断条件
            shengbang_sheet_keywords = ['漏洞', 'vulnerability', 'security', '安全']
            
            for sheet in sheets:
                if any(keyword in sheet.lower() for keyword in shengbang_sheet_keywords):
                    return True
            
            return False
        except Exception as e:
            self.log(f"判断盛邦漏洞扫描报告失败: {str(e)}")
            return False
    
    def extract_specific_info(self, rows):
        """从盛邦漏洞扫描报告中提取特定信息
        
        Args:
            rows (list): 行数据列表
        
        Returns:
            dict: 特定信息字典
        """
        # 根据盛邦漏洞扫描报告的具体格式提取特定信息
        # 这里可以根据实际情况进行扩展
        specific_info = {
            'report_type': 'shengbang',
            'vulnerability_count': 0,
            'scan_date': None
        }
        
        return specific_info
    
    def generate_summary(self, vulnerabilities):
        """生成盛邦漏洞扫描报告摘要
        
        Args:
            vulnerabilities (list): 漏洞字典列表
        
        Returns:
            dict: 报告摘要字典
        """
        summary = {
            'total_vulnerabilities': len(vulnerabilities),
            'severity_counts': {
                '高': 0,
                '中': 0,
                '低': 0,
                '信息': 0
            },
            'by_vulnerability_type': {}
        }
        
        for vuln in vulnerabilities:
            severity = vuln.get('严重程度', '信息')
            if severity in summary['severity_counts']:
                summary['severity_counts'][severity] += 1
            
            vuln_type = vuln.get('安全漏洞名称', '未知')
            if vuln_type not in summary['by_vulnerability_type']:
                summary['by_vulnerability_type'][vuln_type] = 0
            summary['by_vulnerability_type'][vuln_type] += 1
        
        return summary