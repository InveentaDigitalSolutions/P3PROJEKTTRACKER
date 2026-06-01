# Data drop zone

Put your Excel/CSV files here and tell Claude the filenames. **Don't reformat your
existing files** — Claude will inspect the columns and map them. This doc just shows
what the target Dataverse tables need, so you can check nothing critical is missing.

---

## 1. Projects  → `pth_projects`

| Field | Required | Allowed values / notes |
|---|---|---|
| Project ID (external) | ✅ | Your own code, e.g. `PRJ-001`. Used to link tasks. |
| Project Name | ✅ | text |
| Category | ✅ | `ECR` · `Path Forward` · `CIP` · `New Programs` |
| Site | ✅ | `TCA` · `SLP` |
| Objective | ✅ | text |
| Sponsor Executive | ✅ | name |
| Project Manager | ✅ | name |
| Client Code | – | text |
| Budget Allocated | – | number |
| KPIs Impacted | – | text |
| Monitoring Criteria | – | text |
| Planned Start Date | – | YYYY-MM-DD |
| Planned End Date | – | YYYY-MM-DD |

> Don't have a value for a required field? Leave it blank — Claude will flag it and
> we'll decide a default together.

---

## 2. Task template standard  → drives auto-created Milestones + Activities

This is the "every project of category X gets these tasks" rule. Fill in
`task-template.csv` (or hand Claude your own version). One row per task.

Columns:

| Column | Meaning |
|---|---|
| `category` | Which project category this task belongs to |
| `milestone` | Milestone/phase this task rolls up to (groups tasks) |
| `task_name` | The activity name |
| `owner_role` | Role responsible (e.g. Quality Engineer) — or leave blank to assign later |
| `start_offset_days` | Days from project start when the task begins (can be negative) |
| `duration_days` | How many days the task takes |
| `critical_path` | `yes` / `no` |

See `task-template.csv` for example rows you can edit or replace.
