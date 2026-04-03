# GridCoords

A Revit add-in that automatically places coordinate labels at every grid intersection in a plan view.

## What It Does

When working in a floor plan, ceiling plan, or area plan, GridCoords detects all visible grids, finds where they cross, and places a label at each intersection point. For example, if horizontal grid "A" crosses vertical grid "3", the tool places a label reading **(A,3)** at that location.

The tool supports grids at any angle — not just orthogonal grids. Diagonal and curved grids are automatically detected, grouped by angle, and paired for intersection labeling.

This automates what would otherwise be a tedious manual task of identifying and labeling every grid intersection for coordination drawings.

## How to Use

1. Open a **plan view** in Revit (Floor Plan, Ceiling Plan, or Area Plan).
2. Click the **Grid Coords** button on the ribbon (under the Dev Tools panel).
3. The Grid Coords window opens and stays on top while you work.

### Settings (Left Side)

**Placement Mode**
- **TextNote** — Places a Revit text note at each intersection. Choose which Text Type to use from the dropdown.
- **Family Instance** — Places a Generic Annotation family at each intersection. Choose the family type and which text parameter should receive the label text.
- Dropdown selections are preserved when changing Grid Scope or refreshing.

**Label Format**
- **Template** — Controls what the label looks like. The default is `({H},{V})` where `{H}` is replaced with the "more horizontal" grid name and `{V}` with the "more vertical" grid name. Hover over the "Template:" label for format examples and diagonal token assignment details. If left empty, defaults to `({H},{V})`.
- **Order** — Choose whether horizontal or vertical grid names come first.
- **Preview** — Shows a live example of what labels will look like using actual grid names from the first enabled pairing.

**Placement Options**
- **Auto offset** — When checked, the tool automatically offsets labels from the exact intersection point so they don't sit directly on the grid lines. The offset scales with your view scale.
- **Offset X / Offset Y** — Manual offset values in paper inches. Only editable when Auto is unchecked.

**Existing Labels**
- **Delete existing labels first** — Removes all labels previously placed by this tool before placing new ones. This is the default.
- **Skip intersections with existing labels** — Only places labels where none exist yet (useful for adding labels after new grids are added).
- **Place duplicates anyway** — Places labels at all intersections regardless of what's already there.

### Grids (Right Side)

**Grid Scope** — Controls which grids appear in the lists below:
- **All Visible Grids** — Shows every grid in the current view. This is the default.
- **Currently Selected Grids** — Only shows grids you've selected in Revit. The list auto-updates as you select different grids in the canvas.
- **Pick Grids** — Click the "Pick" button, then click individual grids in Revit. Press Finish or Escape when done. Only the picked grids will appear. Use "Clear" to reset.

**Horizontal Grids** and **Vertical Grids** — Grids are automatically classified by their angle relative to the view's horizontal axis and separated into groups. Each group has:
- A **Select All** checkbox to quickly toggle all grids in that group.
- Individual checkboxes displayed in a compact wrap layout.
- A header showing how many grids are selected (e.g., "Horizontal Grids (5 of 5)").

**Diagonal and Curved Grid Groups** — If any grids are at non-orthogonal angles (not within 10 degrees of horizontal or vertical), they appear in their own angle-based groups below the standard H/V expanders. For example, grids at approximately 45 degrees are grouped as "Diagonal Grids (~45 degrees)". Curved grids get their own group. Each dynamic group has the same Select All and checkbox layout as the standard groups.

**Intersection Pairings** — Shows all cross-group pairings that will be used for intersection computation. Each pairing shows:
- An enable/disable checkbox to include or exclude that pairing from placement.
- Which group gets the `{H}` token and which gets the `{V}` token — the group with the smaller angle (more horizontal) is always assigned `{H}`.
- The estimated intersection count for that pairing.

**Intersection Count** — Shows the total intersections across all enabled pairings with a per-pairing breakdown.

### Action Buttons (Bottom)

- **Place Labels** — Calculates all intersections for each enabled pairing and places labels in the active view. Curved grids that intersect at multiple points will receive a label at each intersection.
- **Delete Labels** — Removes labels placed by this tool. When using "Currently Selected Grids" or "Pick Grids" scope, only deletes labels whose grid names match grids displayed in any group. When using "All Visible Grids", deletes all labels in the view.
- **Close** — Closes the window.

### Results

After placing or deleting labels, the Results section appears showing:
- How many labels were placed
- How many existing labels were deleted
- How many intersections were skipped
- How many errors occurred (with details if any)

## Supported Revit Versions

Revit 2022, 2023, 2024, 2025, and 2026.

## Notes

- The tool window stays open (modeless) so you can continue working in Revit while it's displayed.
- Labels placed by this tool are tracked internally using Extensible Storage, so the Delete function only removes labels created by Grid Coords — not other text notes or annotations in the view.
- If the view has a crop region active, only intersections inside the crop boundary will receive labels.
- Grid names are sorted naturally (1, 2, 3, ... 10, 11 rather than 1, 10, 11, 2).
- The grid list and view name auto-refresh when you switch to a different view.
- Grids are classified by computing their angle relative to the view's horizontal direction, using a 10-degree tolerance for clustering. Parallel grids always end up in the same group regardless of their line direction vector.
