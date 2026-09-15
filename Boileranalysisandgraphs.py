import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from analysis import ROOT, MONTHS, load_measurements, summarize

df_boiler = summarize(load_measurements(ROOT / "Boiler.xlsx", "Boiler"))
months = MONTHS

# Recommendations
def get_recommendation(row):
    if pd.isna(row["% Deviation"]):
        return "Insufficient data or zero design baseline"
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
sns.barplot(x="Parameter", y="% Deviation", data=df_boiler, hue="Parameter", legend=False, palette="coolwarm")
plt.axhline(0, color='black', linestyle='--')
plt.title("Boiler Parameter Deviation from Design to Best Performance")
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

plt.title("Combined Monthly Trend for All Boiler Parameters")
plt.xlabel("Month")
plt.ylabel("Operational Value")
plt.grid(True)
plt.xticks(rotation=45)
plt.legend(loc='center left', bbox_to_anchor=(1, 0.5), fontsize='small')
plt.tight_layout()
plt.show()

