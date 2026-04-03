# GridCoords

A Revit add-in that automatically places coordinate labels at every grid intersection in a plan view.

## What It Does

When working in a floor plan, ceiling plan, or area plan, GridCoords detects all visible grids, finds where they cross, and places a label at each intersection point. For example, if horizontal grid "A" crosses vertical grid "3", the tool places a label reading **(A,3)** at that location.

This automates what would otherwise be a tedious manual task of identifying and labeling every grid intersection for coordination drawings.

## How to Use

1. Open a **plan view** in Revit (Floor Plan, Ceiling Plan, or Area Plan).
2. Click the **Grid Coords** button on the ribbon (under the Dev Tools panel).
3. The Grid Coords window opens and stays on top while you work.

### Settings (Left Side)

**Placement Mode**
- **TextNote** — Places a Revit text note at each intersection. Choose which Text Type to use from the dropdown.
- **Family Instance** — Places a Generic Annotation family at each intersection. Choose the family type and which text parameter should receive the label text.

**Label Format**
- **Template** — Controls what the label looks like. The default is `({H},{V})` where `{H}` is replaced with the horizontal grid name and `{V}` with the vertical grid name. Hover over the "Template:" label for format examples. If left empty, defaults to `({H},{V})`.
- **Order** — Choose whether horizontal or vertical grid names come first.
- **Preview** — Shows a live example of what labels will look like using actual grid names from your view.

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

**Horizontal Grids** and **Vertical Grids** — Grids are automatically classified by their direction and separated into two groups. Each group has:
- A **Select All** checkbox to quickly toggle all grids in that group.
- Individual checkboxes displayed in a compact wrap layout.
- A header showing how many grids are selected (e.g., "Horizontal Grids (5 of 5)").

**Angled / Curved** — If any grids are detected at an angle (not aligned with the view's axes) or are curved, they appear in a separate collapsed section. You can reassign them as Horizontal or Vertical using the dropdown next to each grid name.

**Intersection Count** — Shows the total intersections that will be labeled, with the math breakdown (e.g., "25 intersections (5 H x 5 V)").

### Action Buttons (Bottom)

- **Place Labels** — Calculates all intersections and places labels in the active view.
- **Delete Labels** — Removes labels placed by this tool. When using "Currently Selected Grids" or "Pick Grids" scope, only deletes labels for the grids shown in the form. When using "All Visible Grids", deletes all labels in the view.
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
