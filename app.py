
import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt

# File mapping
excel_files = {
    "Base": "BaseScenario.xlsx",
    "High": "HighScenario.xlsx",
    "Low": "LowScenario.xlsx"
}

# Load sectors from Base file
@st.cache_data
def get_sectors():
    df = pd.read_excel(excel_files["Base"], sheet_name=0, header=None)
    return df.loc[2:20, 2].tolist()

# Load data based on forecast type and scenario
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
                    values = df.loc[idx, 13:18].values.tolist()  # N:S
                elif data_type == "Return Forecast":
                    values = [
                        df.loc[idx, 22],  # W
                        df.loc[idx, 31]   # AF
                    ]
                all_data.setdefault(sector, {})[scen] = values
            except ValueError:
                continue
    return all_data

# Streamlit UI
st.set_page_config(layout="centered")
st.title("TownsendAI Forecasting Tool")

forecast_type = st.selectbox("Select forecast type:", ["", "Rent Growth", "Return Forecast"])

if forecast_type:
    scenario = st.selectbox("Select Scenario:", ["", "Base", "High", "Low", "All"])
    sectors = st.multiselect("Select up to 3 sectors:", get_sectors(), max_selections=3)

    if scenario and sectors:
        data = extract_data(scenario, sectors, forecast_type)

        if forecast_type == "Rent Growth":
            years = ["Year 1", "Year 2", "Year 3", "Year 4", "Year 5", "Year 6"]
            st.subheader("Rent Growth Forecasts")
            for sector, scenario_values in data.items():
                fig, ax = plt.subplots()
                bar_width = 0.2
                x = list(range(len(years)))

                for i, (scen, values) in enumerate(scenario_values.items()):
                    offset = [xi + (i - 1) * bar_width for xi in x]
                    ax.bar(offset, values, width=bar_width, label=scen)

                ax.set_xticks(x)
                ax.set_xticklabels(years)
                ax.set_title(f"{sector} - Rent Growth")
                ax.set_ylabel("% Growth")
                ax.legend()
                st.pyplot(fig)

        elif forecast_type == "Return Forecast":
            st.subheader("6-Year Returns (Unlevered and Levered)")
            for sector, scenario_values in data.items():
                fig, ax = plt.subplots()
                width = 0.35
                x = list(range(len(scenario_values)))
                scenarios = list(scenario_values.keys())
                unlevered = [scenario_values[scen][0] for scen in scenarios]
                levered = [scenario_values[scen][1] for scen in scenarios]

                ax.bar([i - width/2 for i in x], unlevered, width=width, label="Unlevered")
                ax.bar([i + width/2 for i in x], levered, width=width, label="Levered")
                ax.set_xticks(x)
                ax.set_xticklabels(scenarios)
                ax.set_title(f"{sector} - Return Forecast")
                ax.set_ylabel("% Return")
                ax.legend()
                st.pyplot(fig)
