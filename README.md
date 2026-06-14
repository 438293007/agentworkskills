# 论文排版助手 V2 — 开发文档

> 对应实现目录: `paper_formatter/` (本仓库根目录)

## 一、项目目标

打造一个基于 **Python + Microsoft Word COM** 的智能论文排版系统,实现:

- 一键完成毕业论文格式化
- 自动生成目录
- 自动设置页码
- 自动设置页眉页脚
- 自动插入分节符
- 自动检查格式错误
- 支持多学校模板
- 支持 JSON 模板驱动

**目标效果**

```
原始论文 Word
   ↓
AI 论文排版助手
   ↓
符合学校要求的最终论文
```

---

## 二、项目目录结构

```
paper_formatter/
├── formatter.py              # 主入口,13 步流程编排
├── config.json               # 配置文件
│
├── templates/                # 学校模板(JSON)
│   ├── xaut.json             # 西安理工大学高科学院
│   ├── xjtu.json             # 西安交通大学(待扩展)
│   ├── nwpu.json             # 西北工业大学(待扩展)
│   └── ...                   # 其他学校
│
├── skills/                   # 所有 skill 模块
│   ├── __init__.py
│   ├── word_com.py           # COM 枚举/工具封装
│   ├── template_skill.py     # 1  应用学校模板
│   ├── cover_skill.py        # 2  封面/扉页/声明/授权书
│   ├── heading_skill.py      # 3  标题层级识别
│   ├── toc_skill.py          # 4  目录
│   ├── section_break_skill.py# 5  奇数页分节符
│   ├── page_number_skill.py  # 6  页码
│   ├── header_footer_skill.py# 7  页眉页脚
│   ├── figure_table_skill.py # 8  图表
│   ├── reference_skill.py    # 9  参考文献
│   └── check_skill.py        # 10 格式检查
│
└── docs/
    └── 开发文档.md
```

---

## 三、主流程

[`formatter.py`](../formatter.py) 的执行顺序:

```
1.  加载配置
2.  连接 Word
3.  加载模板
4.  应用模板
5.  识别标题
6.  重建目录
7.  插入分节符
8.  设置页码
9.  设置页眉页脚
10. 图表处理
11. 参考文献处理
12. 格式检查
13. 保存文档
```

**流程图**

```
Word 文档
   ↓
Template
   ↓
Heading
   ↓
TOC
   ↓
Section Break
   ↓
Page Number
   ↓
Header Footer
   ↓
Figure / Table
   ↓
Reference
   ↓
Check
   ↓
Save
```

---

## 四、Skill 设计

### 1. `template_skill.py`

- **职责**: 应用学校模板
- **内容**:
  - `load_template()`
  - `apply_template()`
- **功能**: 页边距 / 默认字体 / 默认字号 / 正文格式 / Heading 样式
- **输入**: `doc`, `template`
- **输出**: 格式化后的文档

### 2. `cover_skill.py`

- **职责**: 处理封面、扉页、原创性声明、版权授权书
- **识别**:
  - 毕业设计（论文）
  - 学位论文原创性声明
  - 学位论文版权使用授权书
- **应用**:
  - `cover`         校徽 + "毕业设计(论文)"
  - `cover_detail`  扉页:学校代码/学号/密级/分类号
  - `declaration`   原创性声明
  - `copyright_authorization` 版权授权书

### 3. `heading_skill.py`

- **职责**: 自动识别标题层级
- **支持**:
  ```
  1   绪论              → Heading 1
  1.1 FPGA 设计         → Heading 2
  1.1.1 时序分析        → Heading 3
  1.1.1.1 初始化        → Heading 4
  ```
- **Word 内建样式**:
  - `-2 Heading 1`
  - `-3 Heading 2`
  - `-4 Heading 3`
  - `-5 Heading 4`

### 4. `toc_skill.py`

- **职责**: 自动生成目录
- **流程**: 删除旧目录 → 寻找目录锚点 → 插入目录字段 → 更新目录 → 格式化目录 → 目录末尾插入分节符
- **支持**:
  - 一级目录: 黑体 小二 加粗
  - 二三级目录: 宋体 小四
  - 固定 23 磅行距
- **接口**: `rebuild_toc(doc, levels=3, template=None)`

### 5. `section_break_skill.py`

- **职责**: 章节分页
- **学校要求**: 每章从奇数页开始
- **处理对象**:
  - `1 绪论`
  - `2 FPGA 设计`
  - `3 系统实现`
  - `结论`
  - `参考文献`
  - `附录`
- **插入**: `wdSectionBreakOddPage`
- **常量**: `WD_SECTION_BREAK_ODD_PAGE = 4`
- **接口**: `insert_odd_section_breaks(doc)`

### 6. `page_number_skill.py`

- **职责**: 设置页码
- **前置部分**: 致谢 / 摘要 / ABSTRACT / 目录
  - `I / II / III / IV`
- **正文**: 1 / 2 / 3 / 4
  - `RestartNumberingAtSection = True`
  - `StartingNumber = 1`
- **接口**: `setup_page_numbers(doc, template)`

### 7. `header_footer_skill.py`

- **职责**: 设置页眉页脚
- **前置部分**: 致谢 / 摘要 / ABSTRACT / 目录
  - 页眉显示章节名(用 `STYLEREF` 引用 Heading 1,或固定文字)
- **正文**:
  - 奇数页: `2026届XXXX专业毕业设计（论文）`
  - 偶数页: `学生姓名：课题名称`
- **支持**: `OddAndEvenPagesHeaderFooter`
- **接口**: `setup_headers(doc, template)`

### 8. `figure_table_skill.py`

- **职责**: 图表规范
- **图**: `图1.1` / `图1.2` / `图2.1`
- **表**: `表1.1` / `表1.2` / `表2.1`
- **自动检查**:
  - 标题是否存在
  - 编号是否连续
  - 图表是否跨页
- **Word 属性**: `KeepWithNext`, `KeepTogether`

### 9. `reference_skill.py`

- **职责**: 参考文献检查
- **支持类型**: `[J] [M] [D] [P] [S] [R] [N] [C] [A] [EB/OL] [DB/OL]`
- **格式**: 宋体小四,首行缩进 2 字符,固定 23 磅行距
- **检查**:
  - 编号连续
  - 类型齐全
  - 格式规范
- **接口**: `check_references(doc)`, `process_references(doc, template)`

### 10. `check_skill.py`

- **职责**: 最终检查
- **检查项**: 页边距 / 字体 / 字号 / 标题层级 / 目录 / 页码 / 页眉 / 图表 / 参考文献
- **输出示例**:
  ```
  ✔ 一级标题
  ✔ 页边距
  ✘ 图 2.3 跨页
  ✘ 参考文献第 8 条缺少类型
  ```
- **接口**: `check_document(doc, template) -> str`

---

## 五、模板规范

### 文件

`templates/xaut.json`(每所学校一份)。

### 模板组成

```json
{
  "page": {},
  "header": {},
  "footer": {},
  "cover": {},
  "toc_styles": {},
  "heading1": {},
  "heading2": {},
  "heading3": {},
  "heading4": {},
  "body": {},
  "figure": {},
  "table": {},
  "reference": {},
  "section_breaks": {},
  "page_number": {}
}
```

### 新增学校

1. 复制 `xaut.json` 为 `templates/<school>.json`
2. 修改参数(字体、字号、行距、页边距、模板文字等)
3. 在 `config.json` 设置 `default_template` 指向新模板
4. 即可支持新学校

### 字段说明

| 字段 | 含义 | 示例 |
|------|------|------|
| `page.top/bottom/left/right` | 页边距(厘米) | `2.54` / `3.18` |
| `body.size` | 正文字号(pt) | `12` (小四) |
| `body.line_spacing` | 正文行距(磅) | `23` |
| `heading1.size` | 一级标题字号(pt) | `18` (小二) |
| `heading1.line_spacing` | 标题行距(磅) | `32` |
| `cover.title.line` | 封面标题行距(磅) | `32` |
| `figure.title_size` | 图标题字号(pt) | `12` (小四) |
| `reference.first_line_indent` | 首行缩进(字符) | `2` |

---

## 六、V3 规划

### AI 识别论文结构

- 自动识别: 摘要 / ABSTRACT / 目录 / 绪论 / 结论 / 参考文献 / 附录
- 无需人工标注。

### AI 纠错

- 自动发现: 标题格式错误 / 页码错误 / 目录错误 / 参考文献错误
- 并自动修复。

### Claude Code 集成

- 增加 Agent 模式,自动:
  - 读取 Word
  - 分析格式
  - 生成修复方案
  - 执行修复

---

## 七、最终目标

```
Word 论文
   ↓
一键运行
   ↓
符合学校要求
   ↓
自动检查
   ↓
输出最终版论文
```

**目标**: 排版时间 3 小时 → 30 秒,达到毕业论文自动排版工具的可商用水平。

---

## 附录 A: 快速开始

```bash
# 安装依赖(仅 Windows,需安装 Office)
pip install pywin32

# 准备输入: 将待排版的 Word 放到 input/thesis.docx
mkdir -p input
cp your_thesis.docx input/thesis.docx

# 一键运行
python formatter.py
```

排版完成后,结果保存为 `output/thesis.docx`(路径由 `config.json.output_doc` 控制)。

## 附录 B: 常见问题

| 现象 | 可能原因 | 解决 |
|------|----------|------|
| 提示 `未安装 pywin32` | Linux/macOS 环境 | 仅 Windows + Office 支持 |
| 目录未更新 | Word 域未刷新 | 在 Word 中按 `Ctrl+A` → `F9` 刷新所有域 |
| 页眉中文乱码 | 字体未安装 | 安装 宋体 / 黑体 / Times New Roman |
| 图/表编号未识别 | 标题前有多余空格 | 手工清空前缀空格后重跑 |
