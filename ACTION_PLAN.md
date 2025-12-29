# Order System Implementation Tasks

- [ ] **PDF Template Configuration** <!-- id: 0 -->
    - [ ] Add metadata formulas (Date, Store, etc.) to `ORDERS_EXPORT` top section.
    - [ ] Add `SUMIF` formulas for Product Quantity (Singles) for all rows.
    - [ ] Add `SUMIF` formulas for Master Case (MC) for all rows.
    - [ ] Verify `ORDER_DATA` matching logic.

- [ ] **New Product Workflow** <!-- id: 1 -->
    - [ ] Add "Galaxy Gummies" via Web App Sidebar to confirm it saves to `PRODUCTS`.
    - [ ] Update `ORDERS_EXPORT`: Add row, copy formulas, update name match to `*Galaxy Gummies*`.

- [ ] **Order Processing Verification** <!-- id: 2 -->
    - [ ] **Web App Order**: Submit order -> Populate Staging -> Check PDF.
    - [ ] **Manual Adjustment**: Duplicate row -> Edit SKU/Qty -> Populate Staging -> Check PDF.

- [x] **UI Polish** <!-- id: 3 -->
    - [x] Fix Header Alignment (Total Amount) in `index.html`.

