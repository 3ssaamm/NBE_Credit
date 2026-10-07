---
name: spreadsheet-layout-fitting
description: Provides definitive standards, algorithms, and best practices for spreadsheet visual layout, column fitting, row height budgeting, and banner ergonomics across Google Sheets and Excel. Enforces proportional viewport budgeting, prevents the merged-banner autoResize pitfall where Column 1 balloons to 400+ px, and guarantees clean, consistent table typography.
---

# Spreadsheet Layout, Column Fitting & Visual Ergonomics

Spreadsheets must feel like state-of-the-art web dashboards. Unconstrained auto-resizing, disproportionate row heights, and ballooned header columns degrade usability, force horizontal scrolling, and create visual clutter.

This skill governs all layout sizing, column budgeting, and row ergonomics for automated and formula-based spreadsheets.

---

## 1. The Core Architectural Pitfall: The Merged-Banner AutoResize Bug

### Root Cause
In Google Sheets (and Excel automation APIs), calling:
```javascript
// ANTIPATTERN: Never run autoResize on Column 1 if Row 1 has a merged title banner!
sheet.autoResizeColumns(1, numCols);
```
evaluates the content of cell `A1`. When `A1:K1` is merged with a title banner (e.g., `📅 NBE Credit Card — Monthly Overview (Upcoming Forecasts & Previous History)`), the engine measures the **entire merged text string (70–120 characters)** as if it resided inside Column 1 alone.

### The Consequences
1. **Ballooned Column A**: Column 1 expands to **400–600 pixels wide**, leaving enormous, ugly whitespace next to short labels (`Family Member`, `#`, `Date`).
2. **Off-Screen Compression**: Vital data columns (amounts, statuses, upcoming months) are pushed off the viewport, forcing users to scroll horizontally.
3. **Disproportionate Row Heights**: Merged title cells with 14pt+ typography and default auto-height balloon Row 1 to awkward heights (45–60px).

---

## 2. Standard Proportional Sizing System

Every dashboard and data sheet must follow this strict vertical and horizontal sizing scale:

### Vertical Sizing (Row Heights)
| Element Type | Height (px) | Typography / Styling |
| :--- | :--- | :--- |
| **Primary Title Banner (Row 1)** | `30px` (clamped) | 13–14pt Bold, White text, Solid Brand Background |
| **Subtitle / Description (Row 2)** | `20px` (clamped) | 9pt Italic, Muted Gray text |
| **Spacing / Gap Rows** | `12–16px` | Empty spacer rows |
| **Section Header Banners** | `26px` | 11–12pt Bold, Primary / Secondary Accent |
| **Table Column Headers** | `24–26px` | 10pt Bold, Neutral Light Gray / Soft Tint |
| **Data Rows** | `21–23px` | 10pt Regular, Clean padding |
| **Total / Summary Rows** | `24–26px` | 10pt Bold, Soft Accent Highlight |

### Horizontal Sizing (Column Widths)
| Column Type | Width (px) | Guidelines |
| :--- | :--- | :--- |
| **Index / Row `#`** | `45px` | Strict width; numbers 1–999 fit cleanly |
| **Short Date (`yyyy-MM-dd` or `MM/DD`)**| `95–105px` | Fits formatted dates without clipping |
| **Long Date (`EEEE, MMMM d, yyyy`)** | `130–150px` | Full weekday & month display |
| **Currency Amount (`#,##0.00`)** | `110–120px` | Aligned right; accommodates up to 8 digits |
| **Percentage (`0.0%`)** | `80–90px` | Aligned right |
| **Short Label / Person / Category** | `110–140px` | Clean left alignment |
| **Matrix Primary Entity (`Family Member`)** | `160–170px` | Wide enough for longest name + total label |
| **Timeline Matrix Month (`Month yyyy`)** | `105–125px` | Clamped to ensure 8–10 months fit on screen |
| **Description / Merchant / Note** | `220–260px` | Left aligned with clean truncation or wrap |
| **Status Badge (`✅ Paid`, `🕒 Forecast`)** | `120–140px` | Centered or left with emoji indicator |

---

## 3. Production Pattern: The Universal Layout Fitting Helper

Always implement and invoke `applySheetLayoutFitting` after rendering or clearing any sheet:

```javascript
/**
 * Applies crisp visual ergonomics, row height constraints, and safe column widths.
 * Prevents the merged-banner autoResize pitfall where Column 1 expands to 400-600px.
 * 
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet 
 * @param {Object} options Sizing configuration
 */
function applySheetLayoutFitting(sheet, options) {
  if (!sheet) return;
  const opts = options || {};

  // 1. Explicit Row Heights for Banners & Subtitles
  const bRow = opts.bannerRow || 1;
  const sRow = opts.subtitleRow || 2;
  try { sheet.setRowHeight(bRow, opts.bannerHeight || 30); } catch (e) { }
  try { sheet.setRowHeight(sRow, opts.subtitleHeight || 20); } catch (e) { }

  if (opts.customRowHeights) {
    Object.keys(opts.customRowHeights).forEach(r => {
      try { sheet.setRowHeight(parseInt(r, 10), opts.customRowHeights[r]); } catch (e) { }
    });
  }

  // 2. Safe AutoResize (ALWAYS skip Column 1 if it starts a merged title banner!)
  if (opts.autoResizeStartCol && opts.autoResizeColCount && opts.autoResizeColCount > 0) {
    try {
      sheet.autoResizeColumns(opts.autoResizeStartCol, opts.autoResizeColCount);
    } catch (e) { }
  }

  // 3. Apply Explicit Column Width Overrides
  if (opts.explicitColWidths) {
    Object.keys(opts.explicitColWidths).forEach(c => {
      try {
        sheet.setColumnWidth(parseInt(c, 10), opts.explicitColWidths[c]);
      } catch (e) { }
    });
  }

  // 4. Clamp Column Widths within [min, max] Bounds
  if (opts.colWidthBounds) {
    const bounds = opts.colWidthBounds;
    const startCol = opts.autoResizeStartCol || 2;
    const endCol = startCol + (opts.autoResizeColCount || 0) - 1;
    for (let c = startCol; c <= endCol; c++) {
      const bound = bounds[c] || bounds.default;
      if (bound) {
        try {
          const currentWidth = sheet.getColumnWidth(c);
          if (bound.min && currentWidth < bound.min) {
            sheet.setColumnWidth(c, bound.min);
          } else if (bound.max && currentWidth > bound.max) {
            sheet.setColumnWidth(c, bound.max);
          }
        } catch (e) { }
      }
    }
  }
}
```

---

## 4. Screen Budgeting & Zero Horizontal Scroll Rule

When designing multi-month comparison matrixes (such as timeline dashboards):
- **Desktop Viewport Budget**: Standard 1080p browser viewport without sidebar is **~1350–1450 pixels** of table area.
- **Formula Budget**:
  $$\text{Column 1} (165\text{px}) + N \times \text{Month Column} (115\text{px}) \le 1350\text{px}$$
- At $N = 10$ months: $165 + (10 \times 115) = 1315\text{px}$ $\rightarrow$ **Zero horizontal scroll required!**
- Clamping month columns between `105px` and `125px` guarantees all data is visible immediately upon sheet switch.

---

## 5. Multi-Table Column Alignment & Text Fitting Ergonomics

When a single sheet contains vertically stacked tables (e.g. Summary KPI card, Table 1 Family Totals, and Table 2 Itemized Charges):

1. **Strict Column Semantic Alignment**:
   - Never insert an index `#` column in Table 2 if Table 1 starts with `Family Member`. This shifts all columns by one, creates conflicting column widths, and forces Column 1 to serve both a 45px `#` and a 200px Total string.
   - Column A must consistently represent `Family Member` across all stacked tables.
2. **Label Budgeting in KPI Cards**:
   - Keep metric labels under 25 characters (e.g. `Bank Statement Closing Bill:`, `Family Shares Total:`, `Recon Status:`, `One-Time Purchases Total:`). Long labels like `Sum of Family Shares Accounted For:` (36 chars) clip into `Sum of Family Sha`.
   - Keep dates formatted without weekday redundancy (`September 25, 2026` instead of `Sunday, September 25, 2026`).
3. **Merge Multi-Word Status Badges**:
   - Always merge status badges across available right-side columns (e.g. `F5:H5`) so status badges like `✅ Perfect Match (0.00 EGP difference)` have 400px+ of room with zero truncation.
4. **Wrap Long Notes**:
   - Always apply `.setWrap(true)` on description and source/audit note columns (Column H) to prevent text bleeding across empty columns.
5. **Always Set `autoResizeStartCol` and `colWidthBounds`**:
   - Never hardcode arbitrary static widths without safe bounds and auto-resizing. Always specify bounds (e.g. `colWidthBounds: { 2: { min: 155, max: 175 }, 3: { min: 210, max: 260 }, ... }`).

---

## 6. Checklist for Every Sheet Creation / Update

- [ ] Row 1 banner height is explicitly set to `30px`.
- [ ] Row 2 subtitle height is explicitly set to `20px`.
- [ ] Column 1 is NEVER auto-resized across a merged cell range.
- [ ] Column 1 width is explicitly assigned (e.g. `45px` for index `#`, `165px` for family entity names).
- [ ] Columns 2..N are auto-resized starting at column 2 (`autoResizeColumns(2, N - 1)`).
- [ ] Matrix columns have upper bounds clamped (`max: 125px`) to prevent off-screen blowout.
- [ ] Section headers have explicit row heights (`26px`).
