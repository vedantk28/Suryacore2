import streamlit as st
from excel_backend import evaluate_cells

FILE_PATH = "2experiment.xlsx"

# Ingredient names from A2 to A6
ingredient_names = ["Ingredient 1", "Ingredient 2", "Ingredient 3", "Ingredient 4", "Ingredient 5"]
ingredient_cells = [f"B{i}" for i in range(2, 7)]

# Output cells
output_cells = [f"F{i}" for i in range(1, 23)] + [f"F{i}" for i in range(24, 42)] + [f"H{i}" for i in range(24, 40)]

st.set_page_config(layout="wide")
st.title("🧪 Ingredient Impact Simulator")

col1, col2 = st.columns([1, 2])

with col1:
    st.header("🔧 Adjust Ingredient Quantities")
    quantities = []
    for name in ingredient_names:
        qty = st.slider(f"{name}", 0, 100, 50)
        quantities.append(qty)

with col2:
    st.header("📊 Computed Results")
    results, labels = evaluate_cells(FILE_PATH, quantities, output_cells)
    for label, cell, value in zip(labels, output_cells, results.values()):
        st.write(f"**{label}** ({cell}): `{value}`")
