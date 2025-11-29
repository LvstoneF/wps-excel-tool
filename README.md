# WPS Excel 处理工具

一个基于Python的GUI工具，用于处理WPS Excel文档。

## 功能特点

- 支持选择Excel文件（.xlsx, .xls）
- 自动读取并显示工作表列表
- 支持选择输出路径
- 实时日志记录
- 简单易用的GUI界面

## 安装依赖

```bash
pip install -r requirements.txt
```

## 运行程序

```bash
python main.py
```

## 使用说明

1. 点击"浏览"按钮选择要处理的Excel文件
2. 从下拉列表中选择要处理的工作表
3. 点击"浏览"按钮选择输出路径
4. 点击"处理文档"按钮开始处理
5. 查看日志区域的处理结果

## 开发环境

- Python 3.7+
- tkinter (Python标准库)
- openpyxl

## 项目结构

```
.
├── main.py          # 主程序文件
├── requirements.txt # 依赖列表
└── README.md        # 项目说明
```

## 许可证

MIT
