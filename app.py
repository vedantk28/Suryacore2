import streamlit as st
from excel_backend import evaluate_cells
from io import BytesIO

FILE_PATH = "2experiment.xlsx"

ingredient_names = ["Maize", "Jowar", "B.Rice", "Wheat", "Bajra"]
ingredient_cells = [f"B{i}" for i in range(2, 7)]
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

    run_button = st.button("Run Simulation")

with col2:
    st.header("📊 Computed Results")

    if run_button:
        try:
            with open(FILE_PATH, "rb") as f:
                file_bytes = BytesIO(f.read())

            results, labels = evaluate_cells(file_bytes, quantities, output_cells)

            for label, cell, value in zip(labels, output_cells, results.values()):
                st.write(f"**{label}** ({cell}): `{value}`")
        except Exception as e:
            st.error(f"❌ Failed to run simulation: {e}")
