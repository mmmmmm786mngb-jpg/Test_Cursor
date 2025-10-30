### DU post-load processing speed analysis

This guide explains how to prepare data and run charts showing average speed and acceleration versus the baseline scenario for the period in your screenshot.

- Input file: `data/du_tasks_times.csv` (UTF-8, semicolon-separated)
- Output charts: `Документация/Reports/du_speed_analysis/figures/`
- Runner: `Скрипты/analytics/run_du_speed_analysis.ps1`

#### CSV format (semicolon-separated)

Columns:
- `date` — ISO date `YYYY-MM-DD`
- `scenario` — one of: `Типовой`, `Без дублей обменов`, `Без дублей обменов + Параллельные портфели`
- `minutes` — total minutes from the bar on that date

Example rows:

```
date;scenario;minutes
2025-07-31;Типовой;182
2025-10-08;Без дублей обменов;120
2025-10-22;Без дублей обменов + Параллельные портфели;35
```

#### Produced diagrams

- Daily bars by scenario (overlay)
- 7-day rolling average per scenario
- Acceleration vs baseline `Типовой` by date: `acceleration = baseline_minutes / scenario_minutes`
- Average acceleration by calendar week (box/mean)

#### How to run

Option A — extract from HTML automatically:

```
pwsh Скрипты/analytics/run_extract_du_from_html.ps1
```

This will read `Документы/ИзменениеСкоростиОбработки.htm` and write `data/du_tasks_times.csv`.

Option B — provide CSV manually:

1) Place your `du_tasks_times.csv` to `data/`.
2) Run the analysis script:

```
pwsh Скрипты/analytics/run_du_speed_analysis.ps1
```

Charts will appear in `figures/` and a brief run log in the console.



