# 智能文档生成工具

## 功能说明
本工具可根据用户提供的Word模板文件（包含目录结构和大纲标题）及指定主题，自动生成完整文档并保存至results目录。

## 环境要求
- Python 3.8+
- 有效的豆包API密钥

## 快速开始
1. 安装依赖：
```bash
pip install -r requirements.txt
```
2. 设置环境变量（Windows）：
```cmd
setx ARK_API_KEY "your-api-key"
```
3. 运行程序：
```bash
python main.py --template templates/sample.docx --theme "人工智能发展报告" --output result.docx
```

## 参数说明
- --template: 模板文件路径（需预先创建templates目录并放置模板文件）
- --theme: 文档主题内容
- --output: 输出文件名（默认：output.docx）