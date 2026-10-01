import io

import pandas as pd
import streamlit as st

from recipe_aggregator import (
    load_nutrition_db,
    discover_recipes,
    aggregate,
    write_excel,
)

st.set_page_config(page_title="Recipe Aggregator", page_icon="🛒")
st.title("🛒 Recipe Aggregator")


@st.cache_data
def get_db():
    return load_nutrition_db()


db = get_db()
recipes = discover_recipes()

if not recipes:
    st.error("No recipe .xlsx files found next to app.py.")
    st.stop()

if not db:
    st.warning("nutrition_db.xlsx not found: calories and protein will show as 0.")

# ── Servings per recipe (0 = not selected) ───────────────────────────────────
st.subheader("Servings per recipe")
order = []
for path in recipes.values():
    label = path.stem.replace("_", " ").title()
    servings = st.number_input(label, min_value=0, value=0, step=1, key=path.stem)
    if servings > 0:
        order.append((path, int(servings)))

# ── Results ──────────────────────────────────────────────────────────────────
if order:
    summary, combined = aggregate(order, db)

    st.subheader("Dishes")
    st.dataframe(
        pd.DataFrame(
            [(name, *vals) for name, vals in summary.items()],
            columns=["Dish", "Servings", "kcal", "Protein (g)",
                     "kcal / serving", "Protein / serving (g)"],
        ),
        hide_index=True,
        use_container_width=True,
    )

    st.subheader("Shopping list")
    st.dataframe(
        pd.DataFrame(
            [(name, qty_fmt, unit, kcal, prot)
             for name, unit, _qty, qty_fmt, kcal, prot in combined],
            columns=["Ingredient", "Quantity", "Unit", "kcal", "Protein (g)"],
        ),
        hide_index=True,
        use_container_width=True,
    )

    buf = io.BytesIO()
    write_excel(summary, combined, buf)
    st.download_button(
        "⬇️ Download shopping_list.xlsx",
        data=buf.getvalue(),
        file_name="shopping_list.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
else:
    st.info("Enter servings for at least one recipe to generate the list.")
