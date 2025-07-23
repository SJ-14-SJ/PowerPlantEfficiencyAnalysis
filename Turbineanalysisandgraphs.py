import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns


file_path = "project.xlsx"
df = pd.read_excel(file_path, sheet_name='Turbine', skiprows=1)

df = df.iloc[:, :9].dropna(subset=["Parameters"]) # Changed 'Parameter' to 'Parameters'
df.columns = [
    "Parameter", "Design", "October", "November", "December",
    "January", "February", "March", "Performance"
]

months = ["October", "November", "December", "January", "February", "March"]
for col in ["Design", "Performance"] + months:
    df[col] = pd.to_numeric(df[col], errors='coerce')

df["Calculated Performance"] = df.apply(
    lambda row: min(
        (row[month] for month in months if pd.notnull(row[month])),
        key=lambda x: abs(x - row["Design"]),
        default=None
    ),
    axis=1
)

df["% Deviation"] = ((df["Calculated Performance"] - df["Design"]) / df["Design"]) * 100
df["Significant Deviation"] = df["% Deviation"].abs() > 5

def generate_recommendation(row):
    param = str(row["Parameter"]).lower()
    deviation = row["% Deviation"]
    if "pressure" in param:
        return "Check for valve issues or leaks" if deviation < -5 else "Monitor overpressure conditions"
    elif "temperature" in param:
        return "Improve insulation or reheat cycle" if deviation < -5 else "Check for overheating"
    elif "vibration" in param:
        return "Inspect rotor alignment or balancing" if deviation > 5 else "Normal"
    elif "efficiency" in param:
        return "Investigate heat rate or steam flow"
    return "Within acceptable range"

df["Recommendation"] = df.apply(generate_recommendation, axis=1)


plt.figure(figsize=(12, 6))
sns.barplot(x="Parameter", y="% Deviation", hue="Parameter", data=df, palette="coolwarm", dodge=False)
plt.axhline(0, color='black', linestyle='--')
plt.axhline(5, color='green', linestyle=':', label='+5% threshold')
plt.axhline(-5, color='red', linestyle=':', label='-5% threshold')
plt.title("Deviation from Design to Best Performance (Oct–Mar)")
plt.xticks(rotation=45, ha='right')
plt.legend(loc='upper right')
plt.tight_layout()
plt.show()

df_melt = df.melt(id_vars=["Parameter"], value_vars=months, var_name="Month", value_name="Value")
plt.figure(figsize=(12, 6))
sns.lineplot(data=df_melt, x="Month", y="Value", hue="Parameter", marker='o')
plt.title("Monthly Trends (Oct 2024 – Mar 2025)")
plt.tight_layout()
plt.show()

print("\nFINAL ANALYSIS TABLE:\n")
print(df[["Parameter", "Design", "Calculated Performance", "% Deviation", "Recommendation"]].to_string(index=False))


