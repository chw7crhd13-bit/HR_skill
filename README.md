# HR Compensation Excel Skill

OpenClaw skill for HR compensation internship work. It helps prepare and explain a repeatable Excel workflow for payroll and salary data cleaning, validation, reporting, and confidentiality handling.

## What It Covers

- Clean and standardize HR compensation data.
- Check required fields, duplicate employee-month rows, invalid statuses, and salary-band exceptions.
- Use Excel formulas such as `VLOOKUP`, `COUNTIFS`, `SUMIFS`, `AVERAGEIFS`, and `IF`.
- Build department-level summary reports.
- Prepare interview talking points for HR compensation internship roles.
- Generate a demo Excel workbook that can be opened and shown during interview preparation.

## Files

```text
.
├── SKILL.md
├── assets/
│   └── hr_compensation_demo.xlsx
├── references/
│   └── interview-playbook.md
└── scripts/
    └── build_compensation_demo.py
```

## Generate The Demo Workbook

```bash
python3 scripts/build_compensation_demo.py --output assets/hr_compensation_demo.xlsx
```

If `openpyxl` is missing:

```bash
python3 -m pip install openpyxl
```

## Install Into OpenClaw

Clone the repository, then copy it into the OpenClaw skills directory:

```bash
git clone https://github.com/chw7crhd13-bit/HR_skill.git
mkdir -p ~/.openclaw/workspace/skills
cp -R HR_skill ~/.openclaw/workspace/skills/hr-compensation-excel
```

Check whether OpenClaw can see it:

```bash
openclaw skills list
openclaw skills info hr-compensation-excel
openclaw skills check
```

## Interview Pitch

Use this line as the short version:

> I prepared a small HR compensation Excel workflow: raw data stays separate, lookup tables define department and salary-band rules, formulas flag duplicates or abnormal records, and the summary sheet gives department-level payroll totals and average pay. I also considered confidentiality by masking sensitive employee data when sharing reports.

For a fuller explanation, see `references/interview-playbook.md`.
