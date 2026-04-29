---
name: hr-compensation-excel
description: "OpenClaw skill for HR compensation work: clean payroll and salary data, validate employee compensation records, build Excel formula checks, prepare pivot-style reports, and document confidentiality controls for sensitive HR data."
version: "1.0.0"
metadata:
  {"openclaw":{"emoji":"💼","requires":{"bins":["python3"]},"tags":["hr","compensation","excel","payroll","reporting"]}}
---

# HR Compensation Excel

## Purpose

Use this skill when the user is preparing for or performing HR compensation data work, especially:

- 整理薪酬、工资、社保、公积金、奖金、考勤或员工花名册数据。
- 用 Excel 公式检查准确性、完整性、及时性。
- 用 VLOOKUP/XLOOKUP、COUNTIFS、SUMIFS、AVERAGEIFS、IF 等公式做核对。
- 生成按部门、月份、职级、城市的薪酬统计报表。
- 搭建可复用的薪酬 Excel 数据处理流程。

## Core Workflow

1. 明确交付物：清洗后的数据、异常清单、公式核对表、透视汇总表、示例文件或流程说明。
2. 保留原始数据：把源数据放在 `raw_payroll_data`，不要直接覆盖。
3. 标准化字段：月份、员工 ID、姓名、部门、城市、职级、在职状态、基本工资、补贴、奖金、社保、公积金、个税。
4. 建字典表：部门负责人/成本中心、职级薪酬区间、城市规则、员工状态。
5. 做数据校验：
   - 员工 ID、部门、月份、职级、状态等关键字段是否缺失。
   - 同一个员工同一个月份是否重复。
   - 部门、状态、职级是否在字典表中。
   - 基本工资是否超出职级薪酬带宽。
   - 实发工资、扣款、税费是否出现明显异常。
6. 做统计报表：按部门、月份、职级、城市汇总人数、应发、实发、平均工资、福利成本、异常数量。
7. 做保密处理：对外沟通时隐藏姓名、员工 ID 和明细薪酬，只保留必要统计口径。
8. 输出结论：说明发现了哪些异常、哪些数据已确认、还有哪些需要 HR 或财务复核。

## Excel Workbook Structure

When creating an Excel workbook, prefer these sheets:

- `raw_payroll_data`: 原始或轻度标准化数据。
- `lookup_tables`: 部门、职级、城市、状态等规则表。
- `validation_checks`: 行级校验公式和异常标签。
- `summary_report`: 按部门/月度/职级的汇总报表。
- `process_notes`: 流程、公式和保密处理说明。

Recommended Excel features:

- Freeze panes and filters for all data sheets.
- Conditional formatting for `Review` or abnormal rows.
- `VLOOKUP`/`XLOOKUP` for dictionary matching.
- `COUNTIFS` for duplicate and exception counts.
- `SUMIFS` and `AVERAGEIFS` for department or grade summaries.
- Pivot-table-style summary tables for management review.

## Demo Script

Run the bundled script to generate a sample workbook:

```bash
python3 scripts/build_compensation_demo.py --output assets/hr_compensation_demo.xlsx
```

If `python3` does not have `openpyxl`, install it or use the runtime Python that already includes `openpyxl`.

The generated workbook demonstrates:

- raw salary data
- lookup tables
- formula-based validation
- duplicate checks
- salary-band checks
- department summary report
- process notes

## Usage Notes

When the user needs implementation guidance, read `references/usage-guide.md` and help them explain:

- Which Excel formulas solve which business problems.
- How to structure raw data, validation sheets, and summary reports.
- How to handle salary-data confidentiality.
