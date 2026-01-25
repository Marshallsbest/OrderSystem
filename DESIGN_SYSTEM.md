# Design System & Material Specifications

## Overview
This document defines the styling standards for the Order System, ensuring alignment with Material Design 3 (MD3) principles.

## 1. Input Fields (Data Table/Grid)
To maintain high density in order grids, we use a "Dense Outlined" variant of the Material Text Field.

### Specification
- **Component**: Outlined Text Field (Dense)
- **Class Name**: `.table-input` (and aliases `.qty-input`)
- **Visual Specs**:
  - **Height**: `48px` (MD3 Dense Target)
  - **Shape**: `border-radius: 4px` (Small Shape)
  - **Border**: `1px solid var(--md-sys-color-outline)`
  - **Focus**: `1px solid var(--md-sys-color-primary)` (using `outline`)
  - **Typography**: `16px` (Body Large - prevents zoom on iOS)
  - **Background**: `transparent` (Surface)

### Material Mappings
| Property | CSS Variable | MD3 Token |
| :--- | :--- | :--- |
| Container Color | `transparent` | `md.sys.color.surface` |
| Outline Color | `--md-sys-color-outline` | `md.sys.color.outline` |
| Active Indicator | `--md-sys-color-primary` | `md.sys.color.primary` |
| Text Color | `--md-sys-color-on-surface` | `md.sys.color.on-surface` |

## 2. Grid Layouts
- **Matrix Grid**: Uses CSS Grid/Flex hybrid for performance.
- **Headers**: Sticky positioning, `z-index: 20`.
