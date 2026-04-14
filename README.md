# PaperFormatting

一个面向`.docx`论文文档的格式标准化工具。

本项目基于Python和`python-docx`实现，核心目标是把已经写好的论文Word文档自动整理为统一、规范的排版格式。工具能够识别论文中的常见结构元素，并自动调整字体、字号、行距、缩进、页边距、图表标题、参考文献、页眉页脚和页码等格式；同时提供图形界面和可编辑配置文件，方便按学校、课程或模板要求进行微调。

## 适用场景

- 课程论文、结课论文、毕业论文初稿的格式整理
- 已有Word文档，需要快速统一排版规范
- 同一套模板需要反复应用到多篇论文
- 希望通过配置文件调整格式规则，而不是手动逐段修改

## 主要功能

- 自动处理`.docx`论文文档
- 识别并格式化常见论文结构
- 支持中英文标题、摘要和关键词
- 支持正文、参考文献、图表标题、来源注释和公式段落格式化
- 统一设置页面尺寸、页边距、字体、字号、段前段后间距和行距
- 自动写入页眉标题和页脚页码
- 使用JSONC配置文件管理格式规则，支持注释
- 提供Tkinter图形界面，便于直接选择文档和编辑配置

## 当前项目结构

```text
PaperFormatting/
├── docx_formatter_gui.py        # 图形界面入口
├── standardize_docx_paper.py    # 文档格式化核心逻辑
├── formatter_config.jsonc       # 默认格式配置
└── output/                      # 默认输出目录
```

## 工作方式

工具会读取输入的`.docx`文件，识别文档中的典型论文结构，并按配置规则统一排版。当前已覆盖的主要内容包括：

- 中文标题、英文标题
- 中文摘要、英文摘要
- 中文关键词、英文关键词
- 一级到多级标题
- 正文段落
- 图表标题
- 资料来源/数据来源/注释
- 独立公式段落
- 参考文献标题与条目
- 表格单元格文本

## 安装依赖

项目当前没有单独的依赖清单文件，最少需要安装：

```bash
python3 -m pip install python-docx
```

说明：

- `tkinter`通常随Python一起提供，用于图形界面
- 建议使用Python 3.10及以上版本运行

## 快速开始

启动图形界面：

```bash
python3 docx_formatter_gui.py
```

打开界面后可以：

1. 选择要处理的`.docx`论文文档
2. 选择导出目录
3. 查看或修改`formatter_config.jsonc`
4. 勾选是否允许覆盖输出文件、是否写入页眉页脚
5. 点击“开始格式化”

## 配置说明

格式规则通过[`formatter_config.jsonc`](formatter_config.jsonc)管理，主要包含以下几类参数：

- `general`：输出后缀、默认导出目录、是否覆盖、是否添加页眉页脚
- `fonts`：中文与西文字体
- `sizes_pt`：标题、摘要、正文、小五号等字号
- `indents_cm`：段落缩进
- `page`：纸张尺寸、页边距、页眉页脚距离
- `spacing_pt`：标题、图表、公式、参考文献等段前段后间距
- `labels`：摘要、关键词、参考文献等固定标签文本


## 作为模块调用

如果你想在别的Python脚本里复用格式化能力，可以直接调用核心模块：

```python
from pathlib import Path
import standardize_docx_paper as formatter

produced, errors = formatter.run_batch(
    [Path("example.docx")],
    output_dir=Path("output/doc"),
    overwrite=False,
    add_header_footer=True,
)

print("生成文件：", produced)
print("处理问题：", errors)
```
