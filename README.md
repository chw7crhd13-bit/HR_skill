# HR 薪酬 Excel Skill

这是一个用于 OpenClaw 的 HR 薪酬数据处理 skill，主要帮助处理薪酬、工资、社保、公积金、奖金、考勤、员工花名册等 Excel 数据。

它的目标是把重复性的薪酬数据工作整理成一套可复用流程：先保留原始数据，再做字段标准化、公式校验、异常标记和汇总报表。

## 主要功能

- 清洗和标准化 HR 薪酬数据。
- 检查关键字段是否缺失，例如员工 ID、月份、部门、职级、状态等。
- 检查同一员工同一月份是否重复。
- 检查员工状态、部门、职级是否符合字典表规则。
- 检查基本工资是否超出职级薪酬区间。
- 使用 Excel 公式完成核对和统计，例如 `VLOOKUP`、`COUNTIFS`、`SUMIFS`、`AVERAGEIFS`、`IF`。
- 按部门、月份、职级或城市生成汇总报表。
- 生成一个可直接打开的 Excel 示例工作簿。
- 提醒薪酬数据的保密处理方式，例如隐藏姓名、员工 ID 和明细薪酬。

## 文件结构

```text
.
├── SKILL.md
├── assets/
│   └── hr_compensation_demo.xlsx
├── references/
│   └── usage-guide.md
└── scripts/
    └── build_compensation_demo.py
```

## 生成 Excel 示例文件

仓库里已经包含一个示例文件：

```text
assets/hr_compensation_demo.xlsx
```

如果需要重新生成，可以运行：

```bash
python3 scripts/build_compensation_demo.py --output assets/hr_compensation_demo.xlsx
```

如果缺少 `openpyxl`，先安装依赖：

```bash
python3 -m pip install openpyxl
```

生成的工作簿包含：

- `raw_payroll_data`：原始薪酬数据。
- `lookup_tables`：部门、职级、状态等字典表。
- `validation_checks`：公式校验和异常标记。
- `summary_report`：部门级汇总报表。
- `process_notes`：流程说明页。

## 安装到 OpenClaw

克隆仓库：

```bash
git clone https://github.com/chw7crhd13-bit/HR_skill.git
```

复制到 OpenClaw skills 目录：

```bash
mkdir -p ~/.openclaw/workspace/skills
cp -R HR_skill ~/.openclaw/workspace/skills/hr-compensation-excel
```

检查 OpenClaw 是否识别：

```bash
openclaw skills list
openclaw skills info hr-compensation-excel
openclaw skills check
```

## 使用方式

可以让 OpenClaw 使用这个 skill 处理薪酬 Excel 数据，例如：

```bash
openclaw agent --local --message "使用 hr-compensation-excel，帮我检查这个薪酬表的数据完整性、重复记录、职级薪酬区间异常，并生成汇总说明。文件路径：/path/to/payroll.xlsx" --json
```

也可以让它生成一个标准化处理流程：

```bash
openclaw agent --local --message "使用 hr-compensation-excel，帮我设计一个月度薪酬数据处理流程，包括原始数据保留、字段标准化、公式校验、异常清单和部门汇总报表。" --json
```

## 适用场景

- 月度薪酬数据整理。
- 工资、奖金、社保、公积金数据核对。
- 员工花名册与薪酬表匹配。
- 部门薪酬成本统计。
- Excel 公式核对模板搭建。
- 薪酬数据脱敏和共享前检查。
