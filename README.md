# Table-cleaning-
A simple Python tool for cleaning Excel and CSV files ,removing duplicates,handing empty rows and columns,trimming whitespace,and exporting clean data with a user-friendly interface.一款基于Python的表格清洗工具，支持删除重复数据、清理空白行列、适用于工作中的表格清理
- [app.py]Streamlit 主界面，负责文件导入、工作表/分隔符选择、清洗配置、预览、统计和下载。
- [cleaner.py]：核心逻辑，包括 CSV/Excel 读取、编码识别、清洗、公式保护及导出。
- [requirements.txt]：所需依赖及兼容版本范围。
- [run.bat]：Windows 双击启动脚本，自动创建或复用 .venv、检查依赖并启动应用。
- [README.md]：完整中文安装、使用、隐私、安全和兼容性说明。
- [.gitignore]：排除虚拟环境、缓存、日志和用户表格文件。
- [.streamlit/config.toml]：关闭使用统计，只允许本机访问，限制上传大小。
- [samples/sample.csv]：包含重复行、空白行、空白列和多余空格的可测试示例。
- [tests/test_cleaning.py]：16 项核心与导入导出测试。
如何启动
直接双击 [run.bat]。浏览器通常会自动打开；
有问题联系zchen8194@Gmail.com
