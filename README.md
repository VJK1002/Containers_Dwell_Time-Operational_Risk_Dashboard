# Container Dwell Time & Operational Risk Dashboard

An Excel dashboard for tracking container dwell time, demurrage exposure, and operational risk across terminals and customers.

## Overview

This workbook ingests raw container movement data (discharge date, free time, gate-out date) and turns it into:
- Dwell time and risk classification (HIGH / MEDIUM / LOW) per container
- Estimated demurrage cost exposure
- Terminal-level and customer-level performance summaries
- A visual dashboard with pivot tables, charts, and slicers

## File Structure

| Sheet | Purpose |
|---|---|
| `Raw Data` | Source data — one row per container (~1,500 records), plus calculated columns for dwell time, risk level, and demurrage $ |
| `Terminal Performance` | Per-terminal rollups: container count, average dwell time, high-risk count, demurrage exposure |
| `KPI` | Reference sheet explaining each KPI and why it matters |
| `Pivot Table` | Pivot tables summarizing dwell time, demurrage, and risk by terminal/customer |
| `Dashboard` | Visual summary — 6 charts + slicers (Risk Level, Terminal, Cleared/Still in Yard) |

## Key Metrics (KPIs)

- **Average Dwell Time** — how long containers sit before leaving the terminal
- **% Containers Over Free Time** — headline risk indicator
- **Risk Category counts** — HIGH (10+ days over) / MEDIUM (1–9 days over) / LOW (0 days over)
- **Total Estimated Demurrage Exposure** — dollar cost across all containers
- **Avg Demurrage Cost per At-Risk Container**
- **Containers Still in Yard**
- **Avg Dwell by Terminal** / **Demurrage Exposure by Customer**
- **Highest Single Dwell Time** — outlier check

## Data Dictionary (Raw Data columns)

| Column | Field | Source or Formula |
|---|---|---|
| A | S.no | Manual / static |
| B | Container ID | Source data |
| C | Terminal | Source data |
| D | Vessel | Source data |
| E | Customer | Source data |
| F | Type | Source data |
| G | Discharge Date | Source data |
| H | Free Time (days) | Source data |
| I | Gate Out Date | Source data — blank = still in yard |
| J | Free Time Finish Day | `=G5+H5` |
| K | Demurrage Rate/day | Source data |
| L | Dwell Time | `=MAX(0, IF(I5="", $C$2-J5, I5-J5))` |
| M | Risk Level | `=IFS(L5>=10,"HIGH", L5=0,"LOW", L5<=9,"MEDIUM")` |
| N | Demurrage Calculation$ | `=IFS(M5="HIGH",K5*L5, M5="MEDIUM",K5*L5, M5="LOW","NA")` |
| **O** | **Containers still in Yard** *(helper column)* | `=IF(I5="","still in yard","cleared")` |

## 🔧 Helper Column — `Containers still in Yard` (Column O)

This column was added on top of the original source data specifically to support filtering. It doesn't come from the source feed — it's a **derived flag** that reduces the raw "Gate Out Date is blank or not" logic into a plain-text label:

```
=IF(I5="","still in yard","cleared")
```

**Why it exists:** slicers and pivot tables can't filter cleanly on "is this cell blank," but they filter beautifully on a text value. This column feeds the **"Containers still in..."** slicer on the Dashboard tab, letting you isolate containers that have already left vs. ones still sitting in the yard, without touching a single pivot field setting.

If you extend the raw data with more rows, make sure this formula (and the other formula columns J, L, M, N) get copied down to match — Excel Tables (`Ctrl+T`) will do this automatically; a plain range will not.

## Formulas Used — Summary

| Formula | Where | Purpose |
|---|---|---|
| `=G5+H5` | Raw Data, col J | Calculates the free-time deadline from discharge date + allowed free days |
| `=MAX(0, IF(I5="", $C$2-J5, I5-J5))` | Raw Data, col L | Days overdue past the deadline — floors at 0 so early/on-time pickups never register as risk |
| `=IFS(...)` | Raw Data, col M | Buckets each container into HIGH / MEDIUM / LOW risk based on Dwell Time |
| `=IFS(...)` | Raw Data, col N | Calculates $ demurrage exposure — Rate × Dwell Time for HIGH/MEDIUM, "NA" for LOW |
| `=IF(I5="","still in yard","cleared")` | Raw Data, col O | Helper column — plain-text yard status, used to drive the Dashboard slicer |
| `=COUNTIFS('Raw Data'!$C$5:$C$1504, B3)` | Terminal Performance, col C | Container count per terminal (fully locked range) |
| `=AVERAGEIF('Raw Data'!$C$5:$C$1504, B3, 'Raw Data'!$L$5:$L$1504)` | Terminal Performance, col D | Average dwell time per terminal |
| `=COUNTIFS('Raw Data'!$C$5:$C$1504, B3, 'Raw Data'!$M$5:$M$1504, "HIGH")` | Terminal Performance, col E | High-risk container count per terminal |
| `=SUMIFS('Raw Data'!$N$5:$N$1504, 'Raw Data'!$C$5:$C$1504, B3, 'Raw Data'!$M$5:$M$1504, "HIGH")` | Terminal Performance, col F | Total demurrage $ for high-risk containers, per terminal |
| `=E3/C3` | Terminal Performance, col G | % of containers at high risk, per terminal |

## How to Use

1. Open the file in Excel (desktop recommended — slicers and pivot charts render best there).
2. Update `Raw Data` with new records as containers move. Copy formula columns J, L, M, N, and the **O helper column** down to match new rows.
3. On the `Pivot Table` sheet, right-click any pivot → **Refresh** (or **Data → Refresh All**) to pull in new rows.
4. Use the slicers on the `Dashboard` tab to filter by Risk Level, Terminal, or Yard status (powered by the O helper column).


