
import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt

# Load Excel files once at the top level
excel_files = {
    "Base": "BaseScenario.xlsx",
    "High": "HighScenario.xlsx",
    "Low": "LowScenario.xlsx"
}

# Function to read sector names from any scenario file
@st.cache_data
def get_sectors():
    df = pd.read_excel(excel_files["Base"], sheet_name=0, header=None)
    return df.loc[2:20, 2].tolist()

# Function to extract rent growth or return data
def extract_data(scenario, sectors, data_type):
    all_data = {}

    scenario_list = [scenario] if scenario != "All" else list(excel_files.keys())

    for scen in scenario_list:
        df = pd.read_excel(excel_files[scen], sheet_name=0, header=None)
        sector_names = df.loc[2:20, 2].tolist()

        for sector in sectors:
            try:
                idx = sector_names.index(sector) + 2
                if data_type == "Rent Growth":
                    data = df.loc[idx, 13:18].values.tolist()  # Columns N to S
                elif data_type == "Return Forecast":
                    data = [
                        df.loc[idx, 22],  # Column W - Unlevered 6YR
                        df.loc[idx, 31]   # Column AF - Net Levered Fund Level
                    ]
                all_data.setdefault(sector, {})[scen] = data
            except ValueError:
                continue  # sector not found
    return all_data

# Streamlit app
st.set_page_config(layout="centered")
st.title("TownsendAI Forecasting Tool")

# First dropdown: Forecast type
forecast_type = st.selectbox("Select forecast type:", ["", "Rent Growth", "Return Forecast"])

if forecast_type:
    scenario = st.selectbox("Select Scenario:", ["", "Base", "High", "Low", "All"])
    sectors = st.multiselect("Select up to 3 sectors:", get_sectors(), max_selections=3)

    if scenario and sectors:
        data = extract_data(scenario, sectors, forecast_type)

        if forecast_type == "Rent Growth":
            years = ["Year 1", "Year 2", "Year 3", "Year 4", "Year 5", "Year 6"]
            st.subheader("Rent Growth Forecasts")
            for sector, scenarios in data.items():
                for scen, values in scenarios.items():
                    fig, ax = plt.subplots()
                    ax.bar(years, values)
                    ax.set_title(f"{sector} - {scen} Scenario")
                    ax.set_ylabel("% Growth")
                    st.pyplot(fig)

        elif forecast_type == "Return Forecast":
            st.subheader("6-Year Returns (Unlevered and Levered)")
            for sector, scenarios in data.items():
                for scen, values in scenarios.items():
                    st.write(f"**{sector}** - {scen} Scenario")
                    st.write(f"- Unlevered Return: {values[0]:.2f}%")
                    st.write(f"- Net Levered Fund Level Return: {values[1]:.2f}%")
