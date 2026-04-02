
# AutoAssignment

批量处理 Word 格式前端作业、JS 实验报告，自动填充个人信息 + Kimi AI 生成合规答案 / 报告内容，模拟人工手写格式。

---

## 目录

1. [项目简介](#项目简介)
2. [项目结构](#项目结构)
3. [环境要求](#环境要求)
4. [安装步骤](#安装步骤)
5. [配置说明](#配置说明)
6. [使用指南](#使用指南)
7. [核心特性](#核心特性)
8. [注意事项](#注意事项)
9. [常见问题](#常见问题)

---

## 项目简介

AutoAssignment 是一个自动化文档处理工具，专为高校学生设计，用于快速完成前端课程作业和 JavaScript 实验报告。

### 主要功能

- **普通作业处理**：自动识别选择题、简答题，生成专业答案
- **实验报告处理**：自动生成实验结果（Result）和结论（Conclusion）
- **个人信息填充**：自动填充姓名、学号、专业班级等信息
- **批量处理**：支持同时处理多个文档

---

## 项目结构

```
AutoAssignment/
├── assignment_auto/          # 普通作业处理模块
│   └── assignment_auto.py    # 主程序
├── experiment_auto/          # 实验报告处理模块
│   └── experiment_auto.py    # 主程序
├── config.py                 # 配置文件（API密钥、个人信息）
├── README.md                 # 项目说明文档
├── .gitignore               # Git忽略配置
├── input/                   # 输入文件夹（需手动创建）
└── output/                  # 输出文件夹（自动创建）
```

---

## 环境要求

- **Python**: 3.8 或更高版本
- **操作系统**: Windows / macOS / Linux
- **网络**: 需要连接互联网（调用 Kimi API）

### 依赖包

| 包名 | 版本要求 | 用途 |
|------|----------|------|
| python-docx | >=0.8.11 | 读写 Word 文档 |
| requests | >=2.25.0 | HTTP 请求 |

---

## 安装步骤

### 1. 克隆或下载项目

```bash
git clone <repository-url>
cd AutoAssignment
```

### 2. 安装依赖

```bash
pip install -r requirements.txt
```

或者手动安装：

```bash
pip install python-docx requests
```

### 3. 配置项目

1. 复制配置文件模板：
   ```bash
   cp config.example.py config.py
   ```
   （Windows 使用：`copy config.example.py config.py`）

2. 编辑 `config.py`，填写你的信息：
   - **KIMI_API_TOKEN**: 从 [月之暗面开放平台](https://platform.moonshot.cn/) 获取
   - **PERSONAL_INFO**: 填写你的姓名、学号、专业班级等

### 4. 验证安装

```bash
python -c "from assignment_auto.assignment_auto import batch_process_all; print('安装成功')"
```

---

## 配置说明

配置文件位于 `config.py`，包含以下设置项：

### 3.1 API 配置

```python
KIMI_API_TOKEN = "your-api-token-here"
```

**获取方式**：
1. 访问 [月之暗面开放平台](https://platform.moonshot.cn/)
2. 注册账号并创建 API Key
3. 将获取的 Token 填入配置文件

**安全提示**：建议使用环境变量存储 API Token，避免硬编码：

```python
import os
KIMI_API_TOKEN = os.getenv("KIMI_API_TOKEN", "your-default-token")
```

### 3.2 个人信息配置

```python
PERSONAL_INFO = {
    "date": "2026年4月2日",           # 日期格式
    "major_class": "软件工程01",      # 专业班级
    "name": "张三",                   # 姓名
    "student_id": "20230001"          # 学号
}
```

### 3.3 文件夹配置

```python
INPUT_FOLDER = "input"     # 输入文件夹名称
OUTPUT_FOLDER = "output"   # 输出文件夹名称
```

---

## 使用指南

### 4.1 准备文档

1. 在项目根目录创建 `input` 文件夹
2. 将需要处理的 Word 文档（`.docx` 格式）放入该文件夹
3. 确保文档格式符合要求（见下方说明）

### 4.2 运行程序

**重要提示**：必须在项目根目录（`AutoAssignment/` 文件夹）运行脚本。

#### 处理普通作业

```bash
# Windows
python assignment_auto/assignment_auto.py

# macOS / Linux
python3 assignment_auto/assignment_auto.py
```

**适用场景**：
- 选择题
- 简答题
- 概念解释题

#### 处理实验报告

```bash
# Windows
python experiment_auto/experiment_auto.py

# macOS / Linux
python3 experiment_auto/experiment_auto.py
```

**适用场景**：
- JavaScript 实验报告
- 需要生成 Result 和 Conclusion 的文档

### 4.3 查看结果

处理完成后，生成的文档保存在 `output` 文件夹中，文件名前缀为 `finished_`。

---

## 核心特性

| 特性 | 说明 |
|------|------|
| **配置统一管理** | 仅需修改 `config.py` 一处，两个脚本同步生效 |
| **智能内容识别** | 自动识别题目和报告填充位置，精准匹配格式 |
| **专业内容生成** | 答案和报告内容符合大学作业要求 |
| **字数智能控制** | 实验报告 Conclusion 严格控制在 80 词以内 |
| **图片自动跳过** | 图片题、无效内容自动留空 |
| **批量处理能力** | 支持同时处理多个文档，运行过程不中断 |
| **格式完整保留** | 最大程度保留原文档的字体、行距、表头等格式 |

---

## 注意事项

### 6.1 安全事项

- ⚠️ **请勿将 `config.py` 提交到公共仓库**，其中包含 API Token 和个人隐私信息
- ⚠️ **建议将 `config.py` 添加到 `.gitignore`**
- ⚠️ **API Token 请妥善保管**，泄露可能导致账号被盗用

### 6.2 使用限制

- 仅支持 `.docx` 格式，不支持 `.doc`、`.pdf` 等格式
- 需要稳定的网络连接
- API 调用可能产生费用，请注意使用量

### 6.3 文档格式要求

- 作业文档需包含标准题号格式（如 1.、2.、(1)、(2) 等）
- 实验报告需包含 "2. Requirements"、"3. Result"、"4. Conclusion" 等标准章节
- 个人信息区域需包含特定关键词（如 Name、Student ID、Major 等）

### 6.4 免责声明

本工具仅供学习辅助使用，请勿用于：
- 违反学校学术诚信规定的行为
- 商业用途
- 其他违法违规用途

---

## 常见问题

### Q1: 运行时报错 `ModuleNotFoundError: No module named 'config'`

**原因**：Python 找不到配置文件

**解决方法**：
- 确保在项目根目录运行脚本
- 或修改脚本中的路径配置

### Q2: API 请求失败怎么办？

**可能原因**：
- 网络连接问题
- API Token 无效或过期
- 请求频率限制

**解决方法**：
- 检查网络连接
- 验证 API Token 有效性
- 稍后重试

### Q3: 生成的内容不符合要求？

**建议**：
- 检查原始文档格式是否符合要求
- 手动微调生成结果
- 调整 prompt 模板（高级用户）

### Q4: 如何自定义输出格式？

**方法**：
- 修改 `config.py` 中的 `OUTPUT_FOLDER`
- 修改源代码中的格式模板（需要编程基础）

---

## 更新日志

### v1.0.0 (2026-04-02)

- 初始版本发布
- 支持普通作业和实验报告处理
- 实现个人信息自动填充
- 支持批量处理功能

---

## 技术支持

如有问题或建议，欢迎提交 Issue 或联系开发者。

---

**免责声明**：本工具仅供学习交流使用，使用者需自行承担使用风险。
