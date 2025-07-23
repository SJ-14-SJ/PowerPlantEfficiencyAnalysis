import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns

# Load the boiler sheet
df_boiler = pd.read_excel("Boiler.xlsx", sheet_name='Boiler', skiprows=1)

# Clean and rename
df_boiler = df_boiler.iloc[:, :10]
df_boiler.columns = [
    "Parameter", "Design", "October", "November", "December",
    "January", "February", "March", "Performance", "Extra"
]
df_boiler = df_boiler.dropna(subset=["Parameter"])

# Convert to numeric
months = ["October", "November", "December", "January", "February", "March"]
df_boiler[["Design"] + months] = df_boiler[["Design"] + months].apply(pd.to_numeric, errors='coerce')

# Calculate best performance
df_boiler["Calculated Performance"] = df_boiler.apply(
    lambda row: min(
        (row[m] for m in months if pd.notnull(row[m])),
        key=lambda x: abs(x - row["Design"]),
        default=np.nan
    ), axis=1
)

# % Deviation
df_boiler["% Deviation"] = ((df_boiler["Calculated Performance"] - df_boiler["Design"]) / df_boiler["Design"]) * 100
df_boiler["Significant Deviation"] = df_boiler["% Deviation"].abs() > 5

# Recommendations
def get_recommendation(row):
    param = str(row["Parameter"]).lower()
    dev = row["% Deviation"]
    if "temperature" in param:
        return "Check overheating or fuel quality" if dev > 5 else "Possible heat loss or sensor issue" if dev < -5 else "OK"
    elif "pressure" in param:
        return "Check for overpressure" if dev > 5 else "Low pressure, investigate combustion or flow" if dev < -5 else "OK"
    elif "flow" in param:
        return "Inspect flow controls, fans, ducts"
    elif "ash" in param:
        return "Review coal quality and combustion process"
    return "Within acceptable range"

df_boiler["Recommendation"] = df_boiler.apply(get_recommendation, axis=1)

# Create and save the deviation bar chart
plt.figure(figsize=(12, 6))
sns.barplot(x="Parameter", y="% Deviation", data=df_boiler, palette="coolwarm")
plt.axhline(0, color='black', linestyle='--')
plt.title("📉 Boiler Parameter Deviation from Design to Best Performance")
plt.ylabel("% Deviation")
plt.xticks(rotation=45, ha='right')
plt.tight_layout()
plt.show()

# Select month columns
months = ["October", "November", "December", "January", "February", "March"]

# Prepare the plot
plt.figure(figsize=(14, 7))

# Loop through each parameter and plot its line
for i, row in df_boiler.iterrows():
    values = row[months].values.astype(float)
    plt.plot(months, values, marker='o', label=row["Parameter"])

plt.title("📈 Combined Monthly Trend for All Boiler Parameters")
plt.xlabel("Month")
plt.ylabel("Operational Value")
plt.grid(True)
plt.xticks(rotation=45)
plt.legend(loc='center left', bbox_to_anchor=(1, 0.5), fontsize='small')
plt.tight_layout()
plt.show()

