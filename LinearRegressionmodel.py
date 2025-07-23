import pandas as pd
from sklearn.linear_model import LinearRegression
import numpy as np

turbine_path = "project.xlsx"
boiler_path = "Boiler.xlsx"


df_turbine = pd.read_excel(turbine_path, sheet_name="Turbine", skiprows=1)
df_boiler = pd.read_excel(boiler_path, sheet_name="Boiler", skiprows=1)

df_turbine = df_turbine.loc[:, ~df_turbine.columns.str.contains("Unnamed")]
df_boiler = df_boiler.loc[:, ~df_boiler.columns.str.contains("Unnamed")]

df_turbine.columns = df_turbine.columns.str.strip()
df_boiler.columns = df_boiler.columns.str.strip()

months_available = ["October", "November", "December", "January", "February", "March"]
forecast_months = ["April", "May", "June", "July", "August", "September"]
forecast_year = 2025
future_month_labels = [f"{month} {forecast_year}" for month in forecast_months]

future_month_index = np.array([[6], [7], [8], [9], [10], [11]])  # April to Sept

def run_combined_linear_forecast(df, source_name, time_columns):
    combined_forecast = []

    print(f"\n📊 Linear Forecast Results for {source_name}:\n")
    print(f"{'Parameter':<30} " + "  ".join([f"{m:<15}" for m in future_month_labels]))
    print("-" * (30 + 18 * len(future_month_labels)))

    for idx, row in df.iterrows():
        param_name = str(row.iloc[0])
        values = pd.to_numeric(row[time_columns], errors='coerce').dropna().values

        if len(values) >= 3:
            try:
                X = np.arange(len(values)).reshape(-1, 1)
                y = values.reshape(-1, 1)

                model = LinearRegression()
                model.fit(X, y)

                predictions = model.predict(future_month_index).flatten()

                forecast_row = [param_name] + list(predictions)
                combined_forecast.append(forecast_row)

                # Print in terminal
                print(f"{param_name:<30} " + "  ".join([f"{val:.2f}".ljust(15) for val in predictions]))

            except Exception as e:
                print(f"⚠️ Skipping {param_name}: {e}")

    # Create and save DataFrame
    forecast_df = pd.DataFrame(
        combined_forecast,
        columns=["Parameter"] + future_month_labels
    )

    
time_columns_turbine = [col for col in months_available if col in df_turbine.columns]
time_columns_boiler = [col for col in months_available if col in df_boiler.columns]

run_combined_linear_forecast(df_turbine, "Turbine", time_columns_turbine)
run_combined_linear_forecast(df_boiler, "Boiler", time_columns_boiler)
