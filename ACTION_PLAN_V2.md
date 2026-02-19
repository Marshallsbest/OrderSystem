# ACTION PLAN - v1.4.18+ (UX & Logic Hardening)

Current Phase: **Finalizing Web App Polish**

## 🎯 Primary Objectives
1.  **Correct Category Ordering in Web App**: Ensure the web app strictly follows the numeric `Display Order` from the `SETTINGS` sheet, matching the spreadsheet dashboard perfectly.
2.  **Fix "SO" Product Master Case Logic**: Resolve the issue where certain products calculate as master cases when they shouldn't, and ensure the appropriate "Unit" vs "Case" fields are visible and calculating correctly.
3.  **Clean UI Maintenance**: Preserve the removal of unnecessary header labels and footers to keep the interface focused and "pesky-free".

## 🛠️ Implementation Steps

### 1. Category Order Synchronization
- **Problem**: Web app categories are appearing alphabetically instead of using the numeric sort.
- **Root Cause Analysis**: Potential mismatch between raw product category strings and the normalized keys in `window.categorySettings`, or the grouping logic in `js.html` is using unmapped names.
- **Fix**: Update the rendering logic in `js.html` to map every product category to its official "Settings Display Name" during the grouping phase to ensure the sort keys match the settings exactly.

### 2. Master Case & "SO" Product Calculation Fix
- **Problem**: Calculation or Field visibility issues for products with Master Case (units per case) settings.
- **Fix Calculation**: Verify the `calculateTotal` logic in `js.html` correctly handles `units * price` vs `cases * price`.
- **Fix Field Visibility**: Ensure that if a product has a master case capacity, BOTH "Unit Qty" and "MC Qty" fields are available if applicable, or that "Unit Qty" doesn't accidentally calculate as "Case Qty" for "SO" (Special Order?) items.

### 3. Verification & Deployment
- [ ] Push changes via `clasp`.
- [ ] Verify category order matches the `SETTINGS` sheet.
- [ ] Verify "Unit Qty" vs "MC Qty" inputs for various product types.
- [ ] Confirm totals are accurate (Price * Quantity).

---
*Last Updated: 2026-01-19*
