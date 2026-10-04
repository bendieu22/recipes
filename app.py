import io

import pandas as pd
import streamlit as st

from recipe_aggregator import (
    load_ingredient_db,
    discover_recipes,
    aggregate,
    write_excel,
)

# ── Categories ───────────────────────────────────────────────────────────────
# List the recipe file names (without .xlsx) that belong under
# "Snacks & Breakfast". Every recipe not listed here goes under "Meals".
SNACKS_BREAKFAST = {
    "chocolate_cookies",
    "overnight_oats",
    "pancakes",
    "peanut_butter_balls",
}

MEALS = "Meals"
SNACKS = "Snacks & Breakfast"

st.set_page_config(page_title="Recipe Aggregator", page_icon="🛒")
st.title("🛒 Recipe Aggregator")


def pretty(stem: str) -> str:
    return stem.replace("_", " ").title()


@st.cache_data
def get_db():
    return load_ingredient_db()


db = get_db()
recipes = discover_recipes()

if not recipes:
    st.error("No recipe .xlsx files found next to app.py.")
    st.stop()

if not db.loaded:
    st.warning("ingredient_db.xlsx not found: calories, protein and costs will show as 0.")

# Split recipes by category: {category: {file_stem: Path}}
by_category: dict[str, dict[str, object]] = {MEALS: {}, SNACKS: {}}
for path in recipes.values():
    cat = SNACKS if path.stem in SNACKS_BREAKFAST else MEALS
    by_category[cat][path.stem] = path

# ── Pick recipes and servings ────────────────────────────────────────────────
order = []
for category, options in by_category.items():
    if not options:
        continue
    st.subheader(category)
    chosen = st.multiselect(
        f"Add to {category.lower()}",
        options=list(options),
        format_func=pretty,
        placeholder="Choose recipes…",
        key=f"pick_{category}",
    )
    for stem in chosen:
        servings = st.number_input(
            f"{pretty(stem)}: servings",
            min_value=1,
            value=1,
            step=1,
            key=f"srv_{stem}",
        )
        order.append((options[stem], int(servings)))

st.divider()

# ── Results ──────────────────────────────────────────────────────────────────
if not order:
    st.info("Pick at least one recipe above to see the shopping list, cost and nutrition.")
    st.stop()

plan = aggregate(order, db)
t = plan.totals

# Headline numbers
c1, c2 = st.columns(2)
c1.metric("Estimated spend (whole packs)", f"${t['cost_buy']:,.2f}")
c2.metric("Cost of amount actually used", f"${t['cost_used']:,.2f}")
st.caption(
    "“Estimated spend” rounds each ingredient up to the pack you buy in store; "
    "“amount used” only counts the grams in the recipes (the rest is leftovers)."
)

# Nutrition and cost for each meal
st.subheader("Nutrition and cost per meal")
st.dataframe(
    pd.DataFrame(
        [(name, s["servings"], s["kcal_srv"], s["protein_srv"], s["cost_srv"],
          s["kcal"], s["protein"], s["cost"])
         for name, s in plan.summary.items()],
        columns=["Dish", "Servings", "kcal / serving", "Protein / serving (g)", "Cost / serving",
                 "Total kcal", "Total protein (g)", "Total cost (used)"],
    ),
    hide_index=True,
    column_config={
        "kcal / serving": st.column_config.NumberColumn(format="%.0f"),
        "Protein / serving (g)": st.column_config.NumberColumn(format="%.1f"),
        "Cost / serving": st.column_config.NumberColumn(format="$%.2f"),
        "Total kcal": st.column_config.NumberColumn(format="%.0f"),
        "Total protein (g)": st.column_config.NumberColumn(format="%.1f"),
        "Total cost (used)": st.column_config.NumberColumn(format="$%.2f"),
    },
)

# What to buy
st.subheader("What to buy")
st.dataframe(
    pd.DataFrame(
        [(p["name"], p["needed"], p["product"],
          p["packs"] if p["packs"] is not None else None,
          p["cost_buy"], ("⚠ " if p["estimate"] else "") + p["status"])
         for p in plan.purchases],
        columns=["Ingredient", "Needed", "Pack (what you buy)", "Packs", "Cost to buy", "Price status"],
    ),
    hide_index=True,
    column_config={"Cost to buy": st.column_config.NumberColumn(format="$%.2f")},
)

if t["n_estimates"]:
    st.warning(
        f"{t['n_estimates']} price(s) are estimates, not Provigo prices (marked ⚠). "
        "Update them in ingredient_db.xlsx."
    )
for w in plan.warnings:
    st.warning(w)

with st.expander("Ingredient quantities exactly as written in the recipes"):
    st.dataframe(
        pd.DataFrame(
            [(l["name"], l["qty_fmt"], l["unit"], l["kcal"], l["protein"], l["cost"]) for l in plan.lines],
            columns=["Ingredient", "Quantity", "Unit", "kcal", "Protein (g)", "Cost (used)"],
        ),
        hide_index=True,
        column_config={
            "kcal": st.column_config.NumberColumn(format="%.0f"),
            "Protein (g)": st.column_config.NumberColumn(format="%.1f"),
            "Cost (used)": st.column_config.NumberColumn(format="$%.2f"),
        },
    )

buf = io.BytesIO()
write_excel(plan, buf)
st.download_button(
    "⬇️ Download shopping_list.xlsx",
    data=buf.getvalue(),
    file_name="shopping_list.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
st.caption("Prices: Provigo (CAD), see the status and source of each price in ingredient_db.xlsx.")
