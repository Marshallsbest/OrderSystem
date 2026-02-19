# 📐 Master Workflow & System Architecture
**Status:** Draft
**Last Updated:** 2026-01-08

## 🎯 Core Goal
To facilitate the creation and submission of a comprehensive **Order Form**.

---

## 🏗️ The Hybrid Engine: Code + Formulas
The system is not just code; it relies heavily on **in-sheet formulas** and **Named Ranges** as the structural skeleton.

### 1. The "Reset Sheet" Optimization
*   **Problem:** Clearing hundreds of input cells individually is slow.
*   **Solution (Current):**
    1.  **Named Range Inventory:** A function assigns Named Ranges to all product inputs.
    2.  **Tracking:** These names are stored in a list on the `SETTINGS` page.
    3.  **Smart Filtering:** A *formula* checks which of these ranges currently have `Value > 0` and creates a "Dirty List".
    4.  **Action:** The script only reads this "Dirty List" and clears those specific cells.
    *   *Constraint:* Any new product logic MUST effectively register itself with this system so `Reset` continues to work efficiently.

### 2. The Template Constraint
*   **The Look:** The output format is strictly dictated by the Boss. It is **non-negotiable**.
*   **The Logic:** The layout of the template is **non-linear / irregular**. It doesn't follow a simple loop (e.g. Header, Item, Item, Item, Footer).
    *   *Constraint:* We cannot refactor the visual template to be "easier to code". The code must bend to fit the template.

### 3. The "Add Product" Workflow Requirement
Adding a product is not just inserting a row. It involves:
*   Adding the visual row to the "Order Taking" sheet.
*   **Creating a Named Range** for that new input.
*   Ensuring the "Dirty List" formula tracks this new range.
*   Updating the "Middleware" and "Export" sheets to recognize the new SKU.

---

## 🔬 Calibration Plan
Before writing new code, we must map the *existing* formula/range ecosystem.
**Agent Task:** Create diagnostic functions to extract:
1.  All **Named Ranges** currently in the sheet.
2.  The current logic/formulas used for the "Dirty List" mechanism (so we don't break it).
