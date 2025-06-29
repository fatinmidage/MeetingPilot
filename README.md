# 会议侠 - AI会议纪要任务提取工具

[![Python Version](https://img.shields.io/badge/python-3.10+-blue.svg)](https://python.org)
[![License](https://img.shields.io/badge/license-MIT-green.svg)](LICENSE)
[![Platform](https://img.shields.io/badge/platform-Windows%20%7C%20macOS%20%7C%20Linux-lightgrey.svg)](https://github.com/meeting-extractor)

AI驱动的智能会议记录分析工具，自动从会议纪要中提取任务信息并生成结构化Excel表格。基于火山引擎方舟大模型，支持中文会议记录的智能分析。

## ✨ 功能特性

- 🤖 **智能AI分析**: 基于火山引擎方舟大模型，准确理解会议内容语义
- 📋 **结构化提取**: 自动识别并分类信息项和行动项
- 📊 **Excel导出**: 生成格式化的Excel任务清单，即开即用
- 🖥️ **跨平台支持**: 原生支持Windows、macOS和Linux
- 🚀 **简单易用**: 命令行一键操作，无需复杂配置
- 📦 **独立部署**: 支持打包为独立可执行文件

## 🎯 适用场景

- **企业会议**: 团队会议、项目评审、工作汇报
- **学术会议**: 研讨会、论文讨论、学术交流
- **客户沟通**: 需求讨论、方案评审、进度同步
- **个人管理**: 待办整理、任务规划、工作安排

## 🚀 快速开始

### 环境要求

- Python 3.10 或更高版本
- 火山引擎方舟平台API访问权限
- 稳定的网络连接

### 1. 项目安装

```bash
# 克隆项目
git clone https://github.com/meeting-extractor/meeting-extractor.git
cd meeting-extractor

# 安装依赖（推荐使用uv）
uv sync

# 或使用pip
pip install -r requirements.txt
```

### 2. 配置API密钥

```bash
# 复制环境变量模板
cp env_template .env

# 编辑.env文件，填入您的API密钥
ARK_API_KEY=your_actual_api_key_here
MODEL_ID=doubao-seed-1.6-250615
BASE_URL=https://ark.cn-beijing.volces.com/api/v3
```

### 3. 准备会议记录

创建会议记录文本文件（支持UTF-8编码）：

```text
# meeting_notes.txt 示例
项目进度汇报会议

参会人员：张三、李四、王五

讨论内容：
1. 前端开发进度延期，需要张三在本周五前完成登录页面
2. 数据库设计需要优化，李四负责下周一提交优化方案
3. 测试环境搭建，王五在本周三前完成
4. 下次会议定在下周四下午2点

决定事项：
- 增加前端开发人手
- 延长项目时间线一周
```

### 4. 运行提取

```bash
# 基本用法
python meeting_extractor.py meeting_notes.txt

# 指定输出文件
python meeting_extractor.py meeting_notes.txt -o tasks.xlsx

# 显示详细帮助
python meeting_extractor.py --help
```

## 📊 输出示例

工具将生成包含以下字段的Excel表格：

| 任务类型 | 任务描述 | 负责人 | 纳期 | 备注 |
|---------|---------|-------|------|------|
| 行动 | 完成登录页面开发 | 张三 | 2024-01-12 | 前端开发延期 |
| 行动 | 提交数据库优化方案 | 李四 | 2024-01-15 | 性能优化 |
| 行动 | 完成测试环境搭建 | 王五 | 2024-01-10 | 基础设施 |
| 信息 | 下次会议安排 | 全员 | 2024-01-18 | 下周四下午2点 |

## 🔧 高级配置

### 自定义模型参数

在`.env`文件中配置：

```bash
# 模型配置
MODEL_ID=doubao-seed-1.6-250615
BASE_URL=https://ark.cn-beijing.volces.com/api/v3

# 输出配置
MAX_TOKENS=2000
TEMPERATURE=0.1
```

### 批量处理

```bash
# 处理多个文件
for file in *.txt; do
    python meeting_extractor.py "$file" -o "${file%.txt}_tasks.xlsx"
done
```

## 📦 构建部署

### 构建可执行文件

```bash
# 安装构建依赖
uv add pyinstaller

# 构建当前平台的可执行文件
python build.py

# 查看构建选项
python build.py --help

# 清理构建文件
python build.py clean
```

### 部署到目标机器

1. 将生成的可执行文件复制到目标机器
2. 创建`.env`文件配置API密钥
3. 准备会议记录文本文件
4. 运行可执行文件

## 🛠️ 开发指南

### 项目结构

```
会议侠/
├── meeting_extractor.py    # 主程序文件
├── build.py               # 构建打包脚本
├── pyproject.toml         # 项目配置
├── env_template           # 环境变量模板
├── README.md              # 项目文档
├── .gitignore            # Git忽略文件
└── 参考代码/              # 参考资料
    └── 大模型结构化输出.md
```

### 本地开发

```bash
# 安装开发依赖
uv sync --group dev

# 代码格式化
black meeting_extractor.py build.py

# 代码排序
isort meeting_extractor.py build.py

# 代码检查
flake8 meeting_extractor.py build.py

# 运行测试
pytest
```

### 贡献代码

1. Fork 项目仓库
2. 创建功能分支 (`git checkout -b feature/amazing-feature`)
3. 提交更改 (`git commit -m 'Add amazing feature'`)
4. 推送到分支 (`git push origin feature/amazing-feature`)
5. 创建 Pull Request

## 🐛 故障排除

### 常见问题

**Q: API调用失败怎么办？**
A: 检查网络连接和API密钥配置，确保`.env`文件中的`ARK_API_KEY`正确。

**Q: 中文编码乱码？**
A: 确保会议记录文件保存为UTF-8编码格式。

**Q: 提取结果不准确？**
A: 尝试调整会议记录格式，使用更清晰的结构化描述。

**Q: 构建失败？**
A: 确保已安装PyInstaller：`uv add pyinstaller`

### 获取帮助

- 📖 查看项目文档
- 🐛 [提交Issue](https://github.com/meeting-extractor/meeting-extractor/issues)
- 💬 参与社区讨论

## 📄 许可证

本项目采用 [MIT许可证](LICENSE)。

## 🙏 致谢

- [火山引擎方舟平台](https://www.volcengine.com/product/ark) - 提供强大的AI能力
- [Pydantic](https://docs.pydantic.dev/) - 数据验证和设置管理
- [OpenPyXL](https://openpyxl.readthedocs.io/) - Excel文件处理
- [PyInstaller](https://pyinstaller.org/) - Python应用打包

---

<div align="center">
    <p>如果这个项目对你有帮助，请给我们一个 ⭐ Star！</p>
    <p>让AI为你的会议管理增效！</p>
</div>
