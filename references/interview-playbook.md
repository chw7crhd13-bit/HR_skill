# HR Compensation Excel Interview Playbook

## JD Mapping

- Data collection, organization, and checking: explain raw-data preservation, field standardization, duplicate checks, missing-value checks, and abnormal salary checks.
- Compensation statistical reports: explain `SUMIFS`, `COUNTIFS`, `AVERAGEIFS`, pivot tables, department/month/grade summaries, and chart-ready tables.
- Document classification and confidentiality: explain source/report separation, least-necessary sharing, masking employee identifiers, and permission-controlled files.
- Daily operational support: explain repeatable templates, checklists, and clear handoff notes.

## Formula Talking Points

- `VLOOKUP` or `XLOOKUP`: match employee IDs to department, grade, city, or salary-band tables.
- `COUNTIFS`: find duplicate employee-month records and count validation errors by department.
- `SUMIFS`: calculate department salary totals, benefit totals, and monthly payroll cost.
- `AVERAGEIFS`: compare average pay by department, grade, city, or status.
- `IF` plus `AND` or `OR`: produce row-level check results such as `OK`, `Missing field`, `Out of band`, or `Invalid status`.
- Pivot tables: summarize headcount, gross pay, net pay, and error counts for managers.

## Suggested Interview Pitch

"I saw this role cares a lot about accuracy, completeness, and repeatable salary-data processing. So I prepared a small HR compensation Excel workflow: raw data stays separate, lookup tables define department and salary-band rules, formulas flag duplicates or abnormal records, and the summary sheet gives department-level payroll totals and average pay. I also added confidentiality handling, such as masking sensitive fields when sharing reports. This is close to the daily work I expect in the internship: careful data collection, Excel formula checks, pivot-style reporting, and clear handoff."

## Mini Case to Describe

1. Receive a monthly salary file from HR or finance.
2. Check whether every active employee has one and only one record for the month.
3. Use lookup tables to verify department, job grade, and salary band.
4. Calculate gross pay and net pay.
5. Produce a department summary for analysis.
6. List exceptions for the compensation team to review.

## Confidentiality Notes

- Do not send raw salary files in chat tools or public channels.
- Share only the minimum columns needed for review.
- Mask names or employee IDs in demo reports.
- Keep raw data, working files, and final reports in separate folders.
- Delete temporary exports after confirmation.
