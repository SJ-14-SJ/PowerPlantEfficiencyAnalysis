# Energy operations: SQL case study

Business question: which parameters were closest to their design values, how did readings change month to month, and where is the dataset incomplete?

## Data model

`equipment` has one row per equipment category. `parameters` belongs to equipment and stores the design baseline. `measurements` has a composite primary key of parameter and month, preventing duplicate observations. Foreign keys enforce ownership. Ingestion uses parameterized SQL and rebuilds the owned snapshot in a transaction.

## Run

From the repository root, run `python sql_pipeline.py`. This creates `outputs/energy.sqlite` and four CSV reports. SQLite is included with Python; no database server is needed.

| Query | Technique | Interpretation |
| --- | --- | --- |
| `monthly_deviations.sql` | Joins and NULLIF | Relative deviation for each parameter/month; a zero baseline stays undefined. |
| `best_month.sql` | CTE and ROW_NUMBER | Closest observed month per parameter, resolving ties by earliest date. |
| `month_over_month.sql` | CTE and LAG | Change from the previous calendar row; missing values remain unknown. |
| `data_quality.sql` | LEFT JOIN, grouping and conditional counts | Coverage and invalid baselines per parameter. |

## Scope

The bundled data produces 72 measurement rows across 12 parameters, two equipment categories and six months. These counts were verified locally. Monthly observations use the first day as a period label, not a claim about measurement time. Six observations cannot establish causal relationships or quantify operational savings. The workbook's precomputed Performance field is not substituted for monthly measurements.

## Design decisions

- Keep missing month rows so LAG does not silently jump to a more distant observation.
- Avoid ranking undefined deviations or dividing by zero.
- Use a transaction for repeatable refreshes; the pipeline only writes its own generated database.
- Separate SQL from Python orchestration so the analytical logic is easy to review.
