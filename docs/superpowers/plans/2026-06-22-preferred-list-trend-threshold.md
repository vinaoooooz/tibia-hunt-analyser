# Preferred List, Thresholds & Historical Trend — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add Preferred List (top 5 creatures by kills/h), threshold color-coding (red/green for >30% XP/h deviation), historical trend sections, and cleanup existing tables in the Tibia Hunt Analyzer Excel generator.

**Architecture:** All changes go into `hunt_analyzer_excel.py`. No new files, no new dependencies. Three new functions, modifications to four existing functions.

**Tech Stack:** Python 3, openpyxl, PIL/Pillow, requests

---

### Task 1: Global kills/h computation

**Files:**
- Modify: `hunt_analyzer_excel.py`

- [ ] **Step 1: Add global kills/h computation function**

Add before `build_summary_sheet`:

```python
def compute_global_kills_h(sessions):
    kills_per_creature = {}
    total_hours = 0
    for s in sessions:
        h = s.get("hours", 0)
        total_hours += h
        for cname, ccount in s.get("monsters", {}).items():
            kills_per_creature[cname] = kills_per_creature.get(cname, 0) + ccount
    if total_hours == 0:
        return []
    result = []
    for cname, ccount in kills_per_creature.items():
        kh = round(ccount / total_hours)
        result.append((cname, kh, ccount))
    result.sort(key=lambda x: x[1], reverse=True)
    return result[:5]
```

- [ ] **Step 2: Add Preferred List sheet builder**

Add before `build_summary_sheet`:

```python
def build_preferred_list_sheet(wb, top5, all_sessions):
    ws = wb.create_sheet(title="Preferred List")
    
    ws.merge_cells("A1:E1")
    c = ws["A1"]
    c.value = "Preferred List — Top 5 Criaturas por Kills/h"
    c.font = Font(bold=True, size=14, color="1F4E79")
    c.alignment = Alignment(horizontal="center")
    
    total_hours = sum(s.get("hours", 0) for s in all_sessions)
    avg_kh = mean(item[1] for item in top5) if top5 else 0
    
    headers = ["#", "Criatura", "Kills/h", "+/-% Media", "Imagem"]
    for col, h in enumerate(headers, 1):
        ws.cell(row=3, column=col, value=h)
    style_header_row(ws, 3, len(headers))
    
    ws.column_dimensions["A"].width = 6
    ws.column_dimensions["B"].width = 30
    ws.column_dimensions["C"].width = 14
    ws.column_dimensions["D"].width = 14
    ws.column_dimensions["E"].width = 14
    
    for idx, (cname, kh, _) in enumerate(top5):
        row = 4 + idx
        ws.cell(row=row, column=1, value=idx + 1)
        ws.cell(row=row, column=2, value=cname.capitalize())
        cell_kh = ws.cell(row=row, column=3, value=kh)
        cell_kh.number_format = '#,##0'
        if avg_kh > 0:
            pct = round((kh - avg_kh) / avg_kh * 100, 1)
            cell_pct = ws.cell(row=row, column=4, value=f"{pct:+.1f}%")
            if pct > 0:
                cell_pct.font = Font(color="006100")
            elif pct < 0:
                cell_pct.font = Font(color="9C0006")
        else:
            ws.cell(row=row, column=4, value="-")
        
        for col in range(1, 5):
            style_data_cell(ws.cell(row=row, column=col), idx % 2 == 0)
        
        img_data = download_creature_image(cname)
        if img_data:
            xl_img = XlImage(img_data)
            xl_img.width = 64
            xl_img.height = 64
            ws.add_image(xl_img, f"E{row}")
            ws.row_dimensions[row].height = 70
        else:
            ws.cell(row=row, column=5, value="(N/D)")
    
    info_row = 4 + len(top5) + 1
    ws.cell(row=info_row, column=1, value=f"Baseado em {len(all_sessions)} sessoes e {total_hours:.1f}h totais").font = Font(italic=True, color="666666")
```

- [ ] **Step 3: Wire into main()**

In `main()`, after `build_summary_sheet` and before the detail sheets loop:

```python
top5 = compute_global_kills_h(sessions)
if top5:
    build_preferred_list_sheet(wb, top5, sessions)
```

---

### Task 2: Threshold color-coding in detail sheets

**Files:**
- Modify: `hunt_analyzer_excel.py`

- [ ] **Step 1: Add RED_FILL and GREEN_FILL constants**

Near the other style constants (after line 42):

```python
RED_FILL = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")
GREEN_FILL = PatternFill(start_color="CCFFCC", end_color="CCFFCC", fill_type="solid")
```

- [ ] **Step 2: Add threshold helper function**

```python
def apply_threshold_format(ws, group_avg, start_row, end_row, xp_col):
    for row in range(start_row, end_row + 1):
        cell = ws.cell(row=row, column=xp_col)
        val = cell.value
        if isinstance(val, (int, float)) and group_avg > 0:
            deviation = (val - group_avg) / group_avg
            if deviation < -0.30:
                for col in range(1, ws.max_column + 1):
                    ws.cell(row=row, column=col).fill = RED_FILL
            elif deviation > 0.30:
                for col in range(1, ws.max_column + 1):
                    ws.cell(row=row, column=col).fill = GREEN_FILL
```

- [ ] **Step 3: Add deviation column and threshold in detail sheet session table**

In `build_hunt_detail_sheet`, in the session details table section (after line 545), add a new column header "% Media" and compute the deviation:

Find the section starting at line 539:

```python
sess_headers = ["Sessao", "Duracao", "Raw XP Gain", "Raw XP/h", "Profit/h", "Balance", "Kills", "Damage", "Damage/h", "% Media"]
```

And in the loop after that, add:
```python
group_avg = group["avg_raw_xp_h"]
session_xp = s.get("calc_raw_xp_h", 0)
if group_avg > 0:
    dev = round((session_xp - group_avg) / group_avg * 100, 1)
    ws.cell(row=row, column=10, value=f"{dev:+.1f}%")
else:
    ws.cell(row=row, column=10, value="-")
```

Then after the loop, call:
```python
if n_sessions > 0:
    apply_threshold_format(ws, group["avg_raw_xp_h"], shrow + 1, shrow + n_sessions, 4)
```

---

### Task 3: Threshold color-coding in Resumo sheet

**Files:**
- Modify: `hunt_analyzer_excel.py`

- [ ] **Step 1: Clean up Resumo and add threshold column**

The current "Detalhamento por Sessao" section in `build_summary_sheet` (starting around line 432) needs to be modified:

- Add "% Media" to the detail headers  
- Compute deviation per session against its group's avg_raw_xp_h
- Apply red/green formatting

Modify the section after line 438:

```python
detail_headers = [
    "Sessao", "Data", "Duracao", "Raw XP Gain", "Raw XP/h (calc)",
    "Balance", "Profit/h (calc)", "% Media"
]
```

In the session loop, we need to find which group each session belongs to and get that group's average. Since the loops iterate `sorted_groups` then `group["sessions"]`, we already have `group["avg_raw_xp_h"]` available:

```python
for group in sorted_groups:
    group_avg = group["avg_raw_xp_h"]
    for s in group["sessions"]:
        srow += 1
        # ... existing fields ...
        session_xp = s.get("calc_raw_xp_h", 0)
        if group_avg > 0:
            dev = round((session_xp - group_avg) / group_avg * 100, 1)
            ws.cell(row=srow, column=8, value=f"{dev:+.1f}%")
        else:
            ws.cell(row=srow, column=8, value="-")
```

Then after the loop, apply formatting. We need to track the start row from `hrow + 1` to `srow`:
```python
apply_threshold_format(ws, 1, hrow + 1, srow, 5)
```

Wait - this won't work because `apply_threshold_format` takes a single `group_avg` value, but sessions in the Resumo detail table come from different groups with different averages. Let me use a different approach: apply per-group formatting within the group loop.

Actually, let me keep it simpler. I'll apply the formatting per-group within the loop:

```python
for group in sorted_groups:
    group_avg = group["avg_raw_xp_h"]
    group_start_row = srow + 1
    for s in group["sessions"]:
        srow += 1
        # ... write all fields ...
    if group["count"] > 0:
        apply_threshold_format(ws, group_avg, group_start_row, srow, 5)
```

But this is getting complex. Let me simplify by just marking the deviation cell color directly within the loop instead of using the helper function, since each group has a different average.

---

### Task 4: Historical trend section

**Files:**
- Modify: `hunt_analyzer_excel.py`

- [ ] **Step 1: Add trend calculation and section in detail sheets**

In `build_hunt_detail_sheet`, after the "Criaturas Mortas" section (around line 534) and before "Detalhes das Sessoes neste Grupo" (line 535), add:

```python
sessions_with_date = [(s.get("date", ""), s.get("calc_raw_xp_h", 0)) for s in group["sessions"] if s.get("date")]
sessions_with_date.sort(key=lambda x: x[0])

if len(sessions_with_date) >= 2:
    mid = len(sessions_with_date) // 2
    first_half = [x[1] for x in sessions_with_date[:mid]]
    second_half = [x[1] for x in sessions_with_date[mid:]]
    avg_first = mean(first_half) if first_half else 0
    avg_second = mean(second_half) if second_half else 0
    
    if avg_first > 0:
        ratio = avg_second / avg_first
        if ratio >= 1.05:
            trend = u"\u2191 Melhorando"
            trend_color = "006100"
        elif ratio <= 0.95:
            trend = u"\u2193 Piorando"
            trend_color = "9C0006"
        else:
            trend = u"\u2192 Estavel"
            trend_color = "666666"
    else:
        trend = "-"
        trend_color = "666666"
    
    trend_start = cstart + len(group["monsters"]) + 3
    ws.merge_cells(f"A{trend_start}:F{trend_start}")
    c = ws.cell(row=trend_start, column=1, value="Tendencia Historica")
    c.font = Font(bold=True, size=12, color="1F4E79")
    
    labels = [
        ("Tendencia", trend),
        ("Media 1a metade", f"{avg_first:,.0f} XP/h"),
        ("Media 2a metade", f"{avg_second:,.0f} XP/h"),
        ("Melhor sessao", f"{max(x[1] for x in sessions_with_date):,} XP/h ({max(sessions_with_date, key=lambda x: x[1])[0]})"),
        ("Pior sessao", f"{min(x[1] for x in sessions_with_date):,} XP/h ({min(sessions_with_date, key=lambda x: x[1])[0]})"),
    ]
    for i, (label, value) in enumerate(labels):
        row = trend_start + 1 + i
        ws.cell(row=row, column=1, value=label).font = Font(bold=True)
        ws.cell(row=row, column=1).border = THIN_BORDER
        cell_val = ws.cell(row=row, column=2, value=value)
        cell_val.border = THIN_BORDER
        if label == "Tendencia":
            cell_val.font = Font(color=trend_color)

```

- [ ] **Step 2: Convert bar chart to line chart with average line**

Replace the bar chart section (around line 564-577):

```python
if n_sessions > 0:
    sess_chart = LineChart()
    sess_chart.title = "XP/h por Sessao (Tendencia)"
    sess_chart.style = 10
    sess_chart.width = 18
    sess_chart.height = 10
    sdata = Reference(ws, min_col=4, min_row=shrow, max_row=shrow + n_sessions)
    scats = Reference(ws, min_col=1, min_row=shrow + 1, max_row=shrow + n_sessions)
    sess_chart.add_data(sdata, titles_from_data=True)
    sess_chart.set_categories(scats)
    sess_chart.x_axis.title = "Sessao"
    sess_chart.y_axis.title = "XP/h"
    sess_chart.y_axis.scaling.min = 0
    
    from openpyxl.chart.series import SeriesLabel
    from openpyxl.chart.label import DataLabelList
    
    avg_ref = Reference(ws, min_col=4, min_row=shrow, max_row=shrow + n_sessions)
    # Add average line as a separate series with constant value
    # openpyxl doesn't support horizontal line annotations easily,
    # so we create an additional series with the average value
    avg_series_data = []
    group_avg_val = group["avg_raw_xp_h"]
    for _ in range(n_sessions):
        avg_series_data.append(group_avg_val)
    
    avg_series = Reference(ws, min_col=4, min_row=shrow, max_row=shrow + 1)
    from openpyxl.chart import LineChart
    
    ws.add_chart(sess_chart, f"K1")
    sess_chart.y_axis.title = "XP/h"
```

Actually, openpyxl doesn't support horizontal line annotations natively. The practical approach is:
1. Add an extra column with the average value repeated
2. Plot it as a second line series with a dashed style

Let me simplify: write the average values to a hidden column, reference it in the chart.

Actually, the simplest approach: just use the bar chart but add a secondary series with the average.

Let me think about this... openpyxl LineChart has limitations. The simplest working approach:
1. Write the group average in a helper column (e.g., column 11)
2. Use LineChart (from openpyxl.chart.line import LineChart) with two series
3. Set the second series to have a dashed line

But this adds complexity. Let me just convert to a LineChart with the existing data; the average reference can be a dotted line by adding a second series from a helper column.

Let me write a practical version:

```python
if n_sessions > 0:
    from openpyxl.chart.line import LineChart
    from openpyxl.chart.series import DataPoint
    
    # Write average values in a helper column (col 11, hidden)
    avg_col = 11
    ws.cell(row=shrow, column=avg_col, value="Media do Grupo")
    for idx in range(n_sessions):
        ws.cell(row=shrow + 1 + idx, column=avg_col, value=group["avg_raw_xp_h"])
    ws.column_dimensions[get_column_letter(avg_col)].hidden = True
    
    chart = LineChart()
    chart.title = "XP/h por Sessao (Tendencia)"
    chart.style = 10
    chart.width = 18
    chart.height = 10
    
    xp_data = Reference(ws, min_col=4, min_row=shrow, max_row=shrow + n_sessions)
    avg_data = Reference(ws, min_col=avg_col, min_row=shrow, max_row=shrow + n_sessions)
    cats = Reference(ws, min_col=1, min_row=shrow + 1, max_row=shrow + n_sessions)
    
    chart.add_data(xp_data, titles_from_data=True)
    chart.add_data(avg_data, titles_from_data=True)
    chart.set_categories(cats)
    
    # Style the average line
    s = chart.series[1]
    s.graphicalProperties.line.dashStyle = "dash"
    
    chart.x_axis.title = "Sessao"
    chart.y_axis.title = "XP/h"
    chart.y_axis.scaling.min = 0
    
    ws.add_chart(chart, f"K1")
```

---

### Task 5: Cleanup and standardization

**Files:**
- Modify: `hunt_analyzer_excel.py`

- [ ] **Step 1: Rename "TOP 5 LUCRO" to "TOP 5 Profit/h"**

In `build_summary_sheet`, change line 408:
```python
c.value = "TOP 5 Profit/h"
```

- [ ] **Step 2: Remove "Detalhamento por Sessao" from Resumo**

Remove lines 432-465 (the whole detail section at the bottom of Resumo).

- [ ] **Step 3: Remove noisy comparison chart**

In `build_hunt_detail_sheet`, remove lines 579-592 (the "Dano/h vs XP/h vs Lucro/h por Sessao" chart).

- [ ] **Step 4: Prefix sheet names**

In `build_hunt_detail_sheet`, change line 472:
```python
safe_name = re.sub(r'[\\/*?:\[\]]', '_', top_name)[:24]
ws = wb.create_sheet(title=f"Detalhe - {safe_name}")
```

- [ ] **Step 5: Verify and run**

```bash
python hunt_analyzer_excel.py
```

Check that the Excel file is generated with all new features.
