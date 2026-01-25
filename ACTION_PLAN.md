## Operational Protocol
- **Strict Single Agent Mode**: Do NOT spawn sub-agents (browser/terminal) without explicit user permission.
- **Resource Awareness**: Minimize parallel tasks.
- **Always product Agnostic**:Always assume that the product being worked on could be any product.
- **Update Version Numbers**: Always update the version number before every test
- **Design in Material Design**: Material design should be used as  aguide when implementing any GUIs 

## Version History

### v1.8.00 (Current)
- [x] **Hybrid Variation & Multi-Split Logic**
    - [x] Implemented per-variation Format selection (Single, Carton, Multi).
    - [x] Created "Split-on-Submit" macro for Multi rows (Auto-gen Single + Case).
    - [x] Redesigned Variation Sidebar for maximum spacing (24px gaps).
    - [x] Synchronized version tagging across all project files (index, sidebar, config).
- [ ] **Next Steps: Production Validation**
    - [ ] Perform batch upload of Multi-format products.
    - [ ] Verify SKU generation and price auto-calculation in the spreadsheet.

### v1.6.06 (History)
- [x] **Radical Sale Simplification**
    - [x] Stripped away all inferred sale logic.
    - [x] Strictly following `TRUE/FALSE` in the "Sale" column.
    - [x] Forced Parent row status to propagate to all variations.
    - [x] Added visual package check emoji 📦 anchor.
- [ ] **Next Steps: Verify Sale Tags**
    - [ ] Request user to HARD REFRESH and check for PACKAGE EMOJI 📦 in header.

### v1.6.05
    - [x] Iterated version major because minor reached 100.
    - [x] Synchronized version across all core files.

### v1.5.13 (History)
- [x] **Sale Display Restoration**
    - [x] Re-implemented `SALE` badge in product group headers.
    - [x] Added original price strikethrough in header next to the sale price.
- [x] **Branding Isolation & Fixes**
    - [x] Isolated "Main" category branding to the App Bar and Primary Buttons ONLY.
    - [x] Fixed root variable bleeding (labels/inputs now use standard MD3 colors).
    - [x] Implemented server-side style injection for zero-flicker branding on load.
    - [x] Enabled direct cell formatting reading (reads both background & font color).

### v1.4.13 (History)
- [x] **Grid Stabilization & Rendering Complete**
    - [x] Fixed `Service Spreadsheets failed` via chunked flushes (every 5 products).
    - [x] Resolved `Invalid Range Name` error using strict R/C notation blocking.
    - [x] Restored Numeric Quantity Inputs (removed accidental checkboxes).
    - [x] Optimized cleanup by rebuilding the sheet instead of clearing it.
    - [x] Confirmed **Categorization Order** (Pre-Rolls, Vapes, etc.) is respected.

### v1.3.94 (History)
`SheetService.gs` to clean state.
    - [x] Enforce **v1.3.89** tagging across all project files.

### v1.3.87 (History)

### v1.3.74 (Current)
- [x] **Product Management**
    - [x] **Edit Product Feature**:
        - [x] Add `updateProductGroup` to `ProductService.gs`.
        - [x] Update `addProductSidebar.html`.
