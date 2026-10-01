#!/usr/bin/env python3
"""
Recipe Aggregator (with Nutrition)
------------------------------------
Interactively select recipes and servings, then outputs a
combined shopping list with per-meal calorie & protein totals —
both printed to terminal and saved as shopping_list.xlsx.

Can also be imported by app.py (Streamlit web version).

Requires a  nutrition_db.xlsx  file in the same folder as this script.
"""

import sys
from collections import defaultdict
from pathlib import Path

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
except ImportError:
    sys.exit("Missing dependency: run  pip install openpyxl  then try again.")

# ── Paths ────────────────────────────────────────────────────────────────────
SCRIPT_DIR     = Path(__file__).parent
OUTPUT_FILE    = SCRIPT_DIR / "shopping_list.xlsx"
NUTRITION_FILE = SCRIPT_DIR / "nutrition_db.xlsx"

# ── Nutrition DB per-unit conversion hints ────────────────────────────────────
# For ingredients measured in non-gram units, store approximate grams per unit
# so we can convert to kcal / protein.  Users can override these in the DB Notes.
# These are fallback defaults — the DB itself is the source of truth.
UNIT_GRAMS: dict[str, float] = {
    "cloves": 5.0,      # garlic clove
    "units":  50.0,     # generic unit (e.g. egg)
    "packs":  100.0,    # noodle pack (dry)
    "bells":  120.0,    # bell pepper
    "ml":     1.0,      # liquids (density ≈ water unless noted)
    "l":      1000.0,
    "kg":     1000.0,
    "g":      1.0,
}

# ── Styling ───────────────────────────────────────────────────────────────────
HEADER_COLOR    = "2E4057"
SECTION_COLOR   = "048A81"
ALT_ROW_COLOR   = "F0F7F6"
WHITE           = "FFFFFF"
ACCENT_COLOR    = "E8F4F3"
NUTRITION_COLOR = "1B4332"   # dark green for nutrition section


def _side():
    return Side(style="thin", color="CCCCCC")


def _border():
    s = _side()
    return Border(left=s, right=s, top=s, bottom=s)


# ── Nutrition DB ──────────────────────────────────────────────────────────────
def load_nutrition_db() -> dict[str, tuple[float, float, str]]:
    """
    Returns {ingredient_lower: (kcal_per_100g, protein_per_100g, unit_in_recipe)}
    Reads from nutrition_db.xlsx (columns: Ingredient, Unit, Calories, Protein, Notes).
    """
    if not NUTRITION_FILE.exists():
        print(f"  ⚠  nutrition_db.xlsx not found in {SCRIPT_DIR} — nutrition data will be skipped.\n")
        return {}
    wb = load_workbook(NUTRITION_FILE, data_only=True)
    ws = wb.active
    db: dict[str, tuple[float, float, str]] = {}
    for row in ws.iter_rows(min_row=4, values_only=True):
        if not row[0]:
            continue
        name = str(row[0]).strip().lower()
        unit = str(row[1]).strip().lower() if row[1] else ""
        try:
            kcal = float(row[2])
            prot = float(row[3])
        except (TypeError, ValueError):
            kcal, prot = 0.0, 0.0
        db[name] = (kcal, prot, unit)
    return db


def nutrition_for(ingredient: str, qty: float, unit: str,
                  db: dict) -> tuple[float, float]:
    """
    Returns (total_kcal, total_protein_g) for `qty` of ingredient in `unit`.
    Converts qty → grams using UNIT_GRAMS, then scales kcal/protein per 100 g.
    """
    key = ingredient.strip().lower()
    if key not in db:
        return 0.0, 0.0
    kcal_100, prot_100, _ = db[key]
    unit_lower = unit.strip().lower()
    grams_per_unit = UNIT_GRAMS.get(unit_lower, 1.0)
    total_grams = qty * grams_per_unit
    return (kcal_100 * total_grams / 100.0, prot_100 * total_grams / 100.0)


# ── Recipe reading ────────────────────────────────────────────────────────────
def read_recipe(filepath: Path) -> tuple[str, list[tuple[str, float, str]]]:
    wb = load_workbook(filepath, data_only=True)
    ws = wb.active
    recipe_name = str(ws["A1"].value or filepath.stem).strip()
    try:
        base_servings = float(ws["B2"].value or 1)
    except (TypeError, ValueError):
        base_servings = 1.0
    ingredients = []
    for row in ws.iter_rows(min_row=4, values_only=True):
        name, qty, unit = (row[0], row[1], row[2]) if len(row) >= 3 else (None, None, None)
        if not name:
            continue
        try:
            qty_per_serving = float(qty) / base_servings
        except (TypeError, ValueError):
            qty_per_serving = 0.0
        ingredients.append((str(name).strip(), qty_per_serving, str(unit or "").strip()))
    return recipe_name, ingredients


# ── Recipe discovery ──────────────────────────────────────────────────────────
def discover_recipes() -> dict[int, Path]:
    files = sorted(SCRIPT_DIR.glob("*.xlsx"))
    files = [f for f in files
             if f.name not in ("shopping_list.xlsx", "nutrition_db.xlsx")
             and not f.name.startswith("~$")]   # skip Excel temp/lock files
    return {i + 1: f for i, f in enumerate(files)}


# ── Terminal prompt (only used when run from the command line) ────────────────
def prompt_order(recipe_map: dict[int, Path]) -> list[tuple[Path, int]]:
    print("\n" + "═" * 50)
    print("  🍽   RECIPE AGGREGATOR")
    print("═" * 50)
    print("\nAvailable recipes:\n")
    for num, path in recipe_map.items():
        print(f"  [{num}]  {path.stem.replace('_', ' ').title()}")
    print("\n  [0]  Done — generate shopping list\n")

    order: list[tuple[Path, int]] = []
    added: set[str] = set()

    while True:
        raw = input("Select recipe number (or 0 to finish): ").strip()
        if raw == "0":
            if not order:
                print("  ⚠  Please add at least one recipe first.\n")
                continue
            break
        try:
            choice = int(raw)
        except ValueError:
            print("  ✗  Enter a number from the list.\n")
            continue
        if choice not in recipe_map:
            print(f"  ✗  No recipe with number {choice}.\n")
            continue
        path = recipe_map[choice]
        raw_s = input(f"  How many servings of '{path.stem.replace('_', ' ').title()}' would you like? ").strip()
        try:
            servings = int(raw_s)
            if servings <= 0:
                raise ValueError
        except ValueError:
            print("  ✗  Please enter a positive whole number.\n")
            continue
        key = path.stem
        if key in added:
            order = [(p, s + servings) if p.stem == key else (p, s) for p, s in order]
            new_s = next(s for p, s in order if p.stem == key)
            print(f"  ✓  Updated '{path.stem.replace('_', ' ').title()}' to {new_s} servings.\n")
        else:
            order.append((path, servings))
            added.add(key)
            print(f"  ✓  Added {servings}× '{path.stem.replace('_', ' ').title()}'\n")
    return order


# ── Aggregation ───────────────────────────────────────────────────────────────
def aggregate(order: list[tuple[Path, int]], db: dict):
    """
    Returns:
      summary:  {recipe_name: (servings, total_kcal, total_protein, kcal_per_srv, protein_per_srv)}
      combined: [(name, unit, qty, qty_fmt, kcal, protein)]
    """
    summary: dict[str, tuple[int, float, float, float, float]] = {}
    totals: dict[tuple[str, str], float] = defaultdict(float)
    nutri: dict[tuple[str, str], tuple[float, float]] = defaultdict(lambda: (0.0, 0.0))

    for path, servings in order:
        recipe_name, ingredients = read_recipe(path)
        meal_kcal = meal_prot = 0.0
        for name, qty_per, unit in ingredients:
            total_qty = qty_per * servings
            totals[(name, unit)] += total_qty
            k, p = nutrition_for(name, total_qty, unit, db)
            old_k, old_p = nutri[(name, unit)]
            nutri[(name, unit)] = (old_k + k, old_p + p)
            meal_kcal += k
            meal_prot += p
        summary[recipe_name] = (servings, round(meal_kcal, 1), round(meal_prot, 1),
                                round(meal_kcal / servings, 1), round(meal_prot / servings, 1))

    combined = []
    for (name, unit), qty in sorted(totals.items(), key=lambda x: x[0][0].lower()):
        k, p = nutri[(name, unit)]
        combined.append((name, unit, round(qty, 3), _smart_format(qty),
                         round(k, 1), round(p, 1)))
    return summary, combined


def _smart_format(qty: float) -> str:
    if qty == int(qty):
        return str(int(qty))
    return f"{qty:.2f}".rstrip("0").rstrip(".")


# ── Terminal output ───────────────────────────────────────────────────────────
def print_shopping_list(summary, combined):
    print("\n" + "═" * 60)
    print("  🛒  SHOPPING LIST")
    print("═" * 60)
    print("\nDishes ordered:\n")
    print(f"  {'Dish':<28} {'Srv':>4}  {'kcal':>8}  {'Protein':>10}  {'kcal/Srv':>10}  {'Protein/Srv':>10}")
    print("  " + "─" * 80)
    for name, (servings, kcal, prot, kcal_srv, prot_srv) in summary.items():
        kcal_s = f"{kcal:.0f}" if kcal else "—"
        prot_s = f"{prot:.1f} g" if prot else "—"
        kcal_srv = f"{kcal_srv:.0f}" if kcal_srv else "—"
        prot_srv = f"{prot_srv:.1f} g" if prot_srv else "—"
        print(f"  {name:<28} {servings:>4}  {kcal_s:>8}  {prot_s:>10}  {kcal_srv:>8}  {prot_srv:>12}")
    print()
    print(f"  {'Ingredient':<25} {'Qty':>8}  {'Unit':<8}  {'kcal':>8}  {'Protein':>10}")
    print("  " + "─" * 68)
    for name, unit, qty, qty_fmt, kcal, prot in combined:
        kcal_s = f"{kcal:.0f}" if kcal else "—"
        prot_s = f"{prot:.1f} g" if prot else "—"
        print(f"  {name:<25} {qty_fmt:>8}  {unit:<8}  {kcal_s:>8}  {prot_s:>10}")
    print()


# ── Excel output ──────────────────────────────────────────────────────────────
def write_excel(summary, combined, target=OUTPUT_FILE):
    """
    Builds the shopping list workbook.
    `target` can be a Path (saves to disk, the original behaviour)
    or a file-like object such as io.BytesIO (used by the web app).
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Shopping List"

    row = 1

    # Title
    ws.merge_cells(f"A{row}:G{row}")
    ws[f"A{row}"] = "🛒  Shopping List"
    ws[f"A{row}"].font = Font(name="Arial", bold=True, size=15, color=WHITE)
    ws[f"A{row}"].fill = PatternFill("solid", fgColor=HEADER_COLOR)
    ws[f"A{row}"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[row].height = 32
    row += 1

    # ── Dishes section ────────────────────────────────────────────────────────
    for col_idx, label in enumerate(
            ["Dish", "Servings", "", "Calories (kcal)", "Protein (g)",
             "Calories per Serving (kcal)", "Protein per Serving (g)"], 1):
        c = ws.cell(row=row, column=col_idx, value=label)
        c.font = Font(name="Arial", bold=True, size=10, color=WHITE)
        c.fill = PatternFill("solid", fgColor=NUTRITION_COLOR)
        c.alignment = Alignment(horizontal="center")
    row += 1

    for i, (name, (servings, kcal, prot, kcal_srv, prot_srv)) in enumerate(summary.items()):
        fill = PatternFill("solid", fgColor=ALT_ROW_COLOR if i % 2 == 0 else WHITE)
        vals = [name, f"×{servings}", "",
                kcal if kcal else "—", prot if prot else "—",
                kcal_srv if kcal_srv else "-", prot_srv if prot_srv else "-"]
        for col_idx, val in enumerate(vals, 1):
            c = ws.cell(row=row, column=col_idx, value=val)
            c.font = Font(name="Arial", size=10)
            c.fill = fill
            c.border = _border()
            if col_idx in (2, 4, 5, 6, 7):
                c.alignment = Alignment(horizontal="center")
        row += 1

    row += 1  # spacer

    # ── Ingredient header ─────────────────────────────────────────────────────
    for col_idx, label in enumerate(
            ["Ingredient", "Quantity", "Unit", "Calories (kcal)", "Protein (g)", ""], 1):
        c = ws.cell(row=row, column=col_idx, value=label)
        c.font = Font(name="Arial", bold=True, size=10, color=WHITE)
        c.fill = PatternFill("solid", fgColor=SECTION_COLOR)
        c.alignment = Alignment(horizontal="center")
        c.border = _border()
    row += 1

    # ── Ingredient rows ───────────────────────────────────────────────────────
    for i, (name, unit, qty, qty_fmt, kcal, prot) in enumerate(combined):
        fill = PatternFill("solid", fgColor=ALT_ROW_COLOR if i % 2 == 0 else WHITE)
        vals = [name, qty_fmt, unit,
                kcal if kcal else "—",
                prot if prot else "—",
                ""]
        for col_idx, val in enumerate(vals, 1):
            c = ws.cell(row=row, column=col_idx, value=val)
            c.font = Font(name="Arial", size=10)
            c.fill = fill
            c.border = _border()
            if col_idx in (2, 3, 4, 5):
                c.alignment = Alignment(horizontal="center")
        row += 1

    ws.column_dimensions["A"].width = 26
    ws.column_dimensions["B"].width = 12
    ws.column_dimensions["C"].width = 10
    ws.column_dimensions["D"].width = 18
    ws.column_dimensions["E"].width = 14
    ws.column_dimensions["F"].width = 26
    ws.column_dimensions["G"].width = 22

    wb.save(target)
    if isinstance(target, Path):
        print(f"  💾  Saved to: {target}\n")


# ── Main (command-line use) ───────────────────────────────────────────────────
def main():
    db = load_nutrition_db()
    recipe_map = discover_recipes()
    if not recipe_map:
        sys.exit(f"No recipe .xlsx files found in {SCRIPT_DIR}")

    order = prompt_order(recipe_map)
    summary, combined = aggregate(order, db)
    print_shopping_list(summary, combined)
    write_excel(summary, combined)


if __name__ == "__main__":
    main()
