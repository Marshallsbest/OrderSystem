# Antigravity Order System: Technical Specification & User Manual
**Version**: 1.8.00
**Environment**: Google Apps Script / Google Sheets / Material Design 3

## 1. Executive Summary
The Antigravity Order System is a high-performance, mobile-first B2B ordering platform built entirely within the Google Workspace ecosystem. It provides a seamless transition from spreadsheet-based inventory management to a modern, interactive web interface for clients to place orders, which are then processed back into structured data for fulfillment.

---

## 2. Core Capabilities

### 2.1 Multi-Mode Product Rendering
The system detects the structure of your product groups and chooses the most efficient layout:
- **Matrix View**: Triggered automatically for products with multiple variations (e.g., Flavor vs. Strength). It creates a responsive grid allowing rapid entry across multiple SKUs.
- **Compact Flex View**: Designed for simple lists, providing a clean "Single" vs "Case" pack breakdown with clear visual separation.
- **Hybrid Variation Entry**: A specialized sidebar workflow for administrators that allows for rapid entry. It features a per-variation "Format" selector (Single, Carton, Multi) and automatically splits "Multi" entries into both Single and Case catalog items during submission.

### 2.2 Intelligent Sale Management
The systems' pricing engine is uncommonly robust for a spreadsheet-driven app:
- **Inheritance**: If you mark a "Parent" row as on sale, every child variation automatically inherits that status.
- **Auto-Discovery**: If the `Sale Price` is lower than the regular `Price`, the system automatically flags the item as "On Sale" even if you forgot to check the box.
- **Group-Level Awareness**: The main header badge ("SALE") appears dynamically if *any* variation within that group is currently discounted.

### 2.3 Dynamic Brand Engine
Instead of hard-coded colors, the application's look and feel is controlled directly from the `SETTINGS` sheet:
- **What-You-See-Is-What-You-Get (WYSIWYG)**: The system reads the actual background and font colors of your spreadsheet cells.
- **Contrast Calculations**: If a font color is missing, the system uses the CCIR 601 (YIQ) algorithm to determine whether white or black text provides better legibility on your chosen background.

---

## 3. Technical Architecture (The "How It Works")

### 3.1 The Frontend (HTML/CSS/JS)
- **State Machine Architecture**: The app (index.html) operates as a Single-Page Application (SPA). It manages states (Login, Order, Review, Success) by toggling visibility, ensuring no reload is required during a session.
- **Material Design 3 (MD3)**: Utilizing CSS variables to maintain a consistent Google-grade aesthetic.
- **Accordion Animation System**: A custom height-measuring logic in `js.html` allows for smooth opening/closing of product categories while maintaining "Sticky Headers" (the category titles stay at the top as you scroll through products).

### 3.2 The Backend (Google Apps Script Services)
The codebase is modularized into specialized services:
- **ProductService.gs**: The brain of the data retrieval. It uses a **Dynamic Alias Mapper** to find columns. You can rename "Price" to "RP" or "Unit Cost" in your sheet, and the code will still find it.
- **Models.gs**: Standardizes diverse sheet rows into a unified JSON object. This is where the complex Parent-Child hierarchy is resolved.
- **OrderService.gs**: Handles the "Review to Sheet" pipeline. It calculates commissions, updates inventory counts, and prepares the "Line Items" for fulfillment.
- **PDFService.gs**: Manages the staging area (`ORDER_DATA`). It populates Named Ranges dynamically so your PDF template (`ORDERS_EXPORT`) can use simple VLOOKUPs or SUM formulas.

---

## 4. Key Data Models

### 4.1 The Variable Product Model
The system uses a unique "Node" system:
- **Parent Node**: Holds the "Source of Truth" for description, image, color coding, and brand name.
- **Child Node**: Holds the specific SKU, specific variation (e.g., "Indigo Flavor"), and price.
- **Inheritance Logic**:
  1. Check Child for value.
  2. If empty, check Parent.
  3. If empty, check Category defaults.

### 4.2 Dynamic Mapping
The `getProductHeaderMap` function in `ProductService.gs` scans the first row of your sheet. It looks for "Aliases" (e.g., `sku` matches `sku`, `item code`, `product code`). This makes the system "Headless"—you can change your spreadsheet layout without breaking the code.

---

## 5. Workflow Logic

### 5.1 The Order Pipeline
1. **Catalog Sync**: On load, `getProductCatalog` fetches all active rows.
2. **Local Calculation**: As the user types quantities, `calculateTotal` (client-side) performs subtotaling instantly without calling the server.
3. **Review Phase**: The `reviewOrder` function generates a structured summary of only the items ordered (quantity > 0).
4. **Final Submission**: `processOrder` sends the payload to the server, which writes it to the `ORDERS` sheet and triggers the PDF export loop.

### 5.2 Named Range Staging
To ensure the PDF is accurate:
- Every product SKU is associated with a `pdfRangeName`.
- The system automatically creates named ranges like `BBBS_SINGLE` and `BBBS_MULTI`.
- The `setupOrderDataSheet` function clears and repopulates these ranges every time an order is processed, ensuring the "Export" template always has live data.

---

## 6. Maintenance & Protocol

- **Versioning**: Follows a strict `vX.Y.ZZ` format.
- **Logging**: Server-side logs are written to Google Cloud Logging (`console.log`) and client-side logs are available via `F12`.
- **Styling Isolation**: All branding colors are scoped to specific UI elements (Header, Buttons) to prevent "bleeding" into standard form inputs.

---
*Created by Antigravity AI Engine v1.8.00*
