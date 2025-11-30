# 常量定义
IP_PATTERN = r'\d+\.\d+\.\d+\.\d+'
IP_DEVICE_PATTERNS = [
    r'(\d+\.\d+\.\d+\.\d+)\s+(.*)',  # IP 设备名称
    r'(\d+\.\d+\.\d+\.\d+)-(.*)',   # IP-设备名称
    r'(\d+\.\d+\.\d+\.\d+):(.*)',   # IP:设备名称
    r'(\d+\.\d+\.\d+\.\d+)\s*->\s*(.*)'  # IP -> 设备名称
]
SEVERITY_MAP = {
    "高危险": "高",
    "中危险": "中",
    "低危险": "低",
    "高危": "高",
    "中危": "中",
    "低危": "低",
    "信息": "信息",
    "信息级": "信息"
}
VULN_SHEET_KEYWORDS = ["漏洞详情", "Sheet1"]
VULN_SHEET_PREFIX = "漏洞详细"
IP_COLUMN_INDEX = 2  # 关联资产/域名列，第3列，索引为2
HOST_DETAIL_SHEET_KEYWORDS = ["主机详情"]  # 主机详情子表关键字
HOST_STAT_SHEET_NAME = "主机漏洞统计"  # 主机漏洞统计工作表名称
SEVERITY_LEVELS = ["高", "中", "低"]  # 统计的严重程度级别
