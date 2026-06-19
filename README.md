# Paper Formatter

论文格式修改 Claude Code Skill — 通过自然语言或参考模板，一键将论文排版为目标格式。

## 功能

- 📄 **页面布局**：纸张大小、页边距、纵向/横向
- 🔤 **字体样式**：中英文字体、字号、加粗、斜体（正确处理 CJK 加粗）
- 📝 **段落格式**：行间距、段间距、对齐方式、首行缩进
- 📑 **标题层级**：Heading 1-6 样式自定义
- 📋 **页眉页脚**：页眉文字、页码位置
- 🔍 **格式检测**：自动分析文档当前格式参数
- 🎯 **三种模式**：预设格式 / JSON 配置 / 模板学习

## 安装

```bash
# 1. 克隆到 Claude Code skills 目录
git clone <repo-url> ~/.claude/skills/paper-formatter

# 2. 安装依赖
pip install python-docx
# 可选：pypandoc_binary（PDF/Markdown 转换）

# 3. 重启 Claude Code 会话即可生效
```

## 快速开始

### 方式 1：预设格式（最简单）

```bash
python scripts/format_docx.py 论文.docx 输出.docx --preset gb-thesis
```

可用预设：
| 预设 | 说明 |
|------|------|
| `gb-thesis` | GB/T 学位论文标准（A4, 宋体小四, 黑体标题, 1.5倍行距） |
| `ieee` | IEEE 会议/期刊论文 |
| `apa` | APA 7th 论文格式 |

### 方式 2：从模板学习

```bash
python scripts/format_docx.py 论文.docx 输出.docx --template 参考论文.docx
```

### 方式 3：格式检测

```bash
python scripts/detect_format.py 论文.docx
```

### 方式 4：对话式（推荐）

直接在 Claude Code 中说：

> "帮我把 D:\论文\毕业论文.docx 按 GB/T 学位论文格式排版"

## 预设格式详情

### `gb-thesis` — GB/T 学位论文

| 属性 | 值 |
|------|-----|
| 纸张 | A4 |
| 页边距 | 上3.0 下3.0 左3.0 右2.5 cm |
| 正文 | 宋体 / Times New Roman 小四 (12pt) |
| 一级标题 | 黑体 三号 (16pt) 加粗 |
| 二级标题 | 黑体 四号 (14pt) 加粗 |
| 三级标题 | 黑体 小四 (12pt) 加粗 |
| 行距 | 1.5 倍 |
| 对齐 | 两端对齐 |
| 首行缩进 | 2 字符 |

## 项目结构

```
paper-formatter/
├── SKILL.md                    # Skill 入口（触发条件 + AI 执行指南）
├── README.md
├── references/
│   ├── gb-layout.md            # GB/T 学位论文格式规范
│   └── workflow.md             # 详细执行流程
├── scripts/
│   ├── install_deps.py         # 依赖检查安装
│   ├── detect_format.py        # 文档格式检测
│   └── format_docx.py          # 核心格式修改引擎
└── assets/templates/           # 模板文件
```

## 支持的格式

**当前 (Phase 1)**：
- [x] .docx（完整支持）

**计划中**：
- [ ] Phase 2: 引用格式转换 (APA/GB/IEEE/MLA 互转)
- [ ] Phase 3: LaTeX (.tex) 支持
- [ ] Phase 4: Markdown + PDF 管道

## License

MIT
