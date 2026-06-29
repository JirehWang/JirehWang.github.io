# DESIGN.md - LKC New Family UI System Specification

## Design Summary
This design system defines the visual guidelines for the Linkou Church New Family Management System (林口教會新家人管理系統). It focuses on high legibility, clean visual layout, subtle micro-motion, and a warm editorial forest palette to project trust and care.

## Brand Voice
- **Warm & Organic**: Organic forest greens represent growth and spiritual care.
- **Utilitarian & Precise**: Clean alignment, distinct borders, and responsive grid layouts prevent visual clutter.
- **Trustworthy**: Clean typography and state badges ensure volunteers can easily scan lists.

## Color Tokens
- **Background (`--bg`)**: `#f4f6f5` (soft green-grey off-white)
- **Surface (`--surface`)**: `#ffffff`
- **Surface Muted (`--surface-2`)**: `#e9efe3` (warm sage tint)
- **Text Primary (`--text`)**: `#16221f` (deep forest charcoal)
- **Text Muted (`--muted`)**: `#52635f`
- **Line/Border (`--line`)**: `#cbd5d0`
- **Primary Accent (`--primary`)**: `#1f5548` (deep forest green)
- **Primary Hover (`--primary-hover`)**: `#174339`
- **Warning/Danger (`--danger`)**: `#a63e3e`
- **Attention/Warning (`--warning`)**: `#d97706`

## Typography
- **Heading Font**: `system-ui, -apple-system, sans-serif`
- **Body Font**: `"Noto Sans TC", "Microsoft JhengHei", Arial, sans-serif`
- **Text Sizes**:
  - Main Title: `26px`, bold
  - Subsection Title: `16px`, semi-bold
  - Table Headers: `13px`, semi-bold, uppercase-like
  - Body/Cells: `14px`
  - Badges/Metas: `11px` / `12px`

## Spacing and Layout
- **App Max Width**: `min(1200px, calc(100% - 32px))`
- **App Padding**: `24px` on desktop, `16px` on mobile.
- **Border Radius**:
  - Cards / Panels: `10px`
  - Input fields / buttons: `6px`
  - Badges: `4px`
- **Grid gap**: `16px` for form inputs, `12px` for table toolbars.

## Components and CTA
- **Primary Action Button**: Deep forest green, solid background, smooth white text. Hover scales it up slightly (`transform: translateY(-1px)`) and darkens the color.
- **Secondary Action Button**: Transparent/white background with thin border, forest green text.
- **Table Grid**: Excel-like fixed headers, border outline, column drag resizers with active hover indicator (primary accent line).
- **Status Badges**:
  - Active: Sage green background, dark green text.
  - Closed: Soft grey-green background, muted text.
  - Warning (Overdue): Soft red background, dark red text.

## Responsive Rules
- Below `768px`, form grids collapse to 1-column.
- Toolbar options stack vertically.
- Table wrapper adds scrollbars (`overflow-x: auto`) for wide datasets.

## Accessibility Rules
- Maintain contrast ratios of at least `4.5:1` for body text.
- Use explicit labels (`aria-label`) on icon-only close buttons.

## Implementation Notes
- Add smooth transitions to all interactive button states (`transition: all 0.2s cubic-bezier(0.4, 0, 0.2, 1)`).
- Table headers should remain sticky during vertical scrolls.
