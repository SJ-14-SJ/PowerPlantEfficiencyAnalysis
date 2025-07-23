import pandas as pd

turbine_path = "project.xlsx"
boiler_path = "Boiler.xlsx"

df_boiler = pd.read_excel(boiler_path, sheet_name="Boiler", skiprows=1)
unnamed_cols = [col for col in df_boiler.columns if "Unnamed:" in col]
if unnamed_cols:
    df_boiler = df_boiler.drop(columns=unnamed_cols)


boiler_columns = [
    "Parameter", "Design", "October", "November", "December",
    "January", "February", "March", "Performance"
]
df_boiler.columns = boiler_columns
df_boiler["Source"] = "Boiler"

df_turbine = pd.read_excel(turbine_path, sheet_name="Turbine", skiprows=1)
unnamed_cols = [col for col in df_turbine.columns if "Unnamed:" in col]
if unnamed_cols:
    df_turbine = df_turbine.drop(columns=unnamed_cols)

turbine_columns = [
    "Parameter", "Design", "October", "November", "December",
    "January", "February", "March", "Performance"
]
df_turbine.columns = turbine_columns
df_turbine["Source"] = "Turbine"

combined_df = pd.concat([df_boiler, df_turbine], ignore_index=True)

combined_df["% Deviation"] = ((combined_df["Performance"] - combined_df["Design"]) / combined_df["Design"]) * 100


def gpt_like_recommendation(row):
    param = str(row["Parameter"]).lower()
    deviation = row["% Deviation"]

    if abs(deviation) < 0.5:
        return f"The {row['Parameter']} is operating well within acceptable limits. No action is required currently."

    elif "pressure" in param:
        if deviation > 0:
            return f"The pressure for {row['Parameter']} is {deviation:.2f}% above the design value. Consider checking for over-pressurization or issues with PRVs and relief valves."
        else:
            return f"The pressure for {row['Parameter']} is {abs(deviation):.2f}% below the design value. Investigate for leakages, valve malfunction, or underperformance in feed systems."

    elif "temperature" in param:
        if deviation > 0:
            return f"{row['Parameter']} is running hotter than expected by {deviation:.2f}%. Evaluate temperature sensors and ensure proper cooling or heat dissipation."
        else:
            return f"{row['Parameter']} shows a {abs(deviation):.2f}% temperature drop. Verify combustion quality, heat exchangers, and insulation."

    elif "air" in param or "flue" in param:
        return f"Deviation in {row['Parameter']} is {deviation:.2f}%. Inspect fans, dampers, and APH for flow inconsistencies or mechanical issues."

    elif "fuel" in param or "coal" in param:
        return f"The deviation in {row['Parameter']} is {deviation:.2f}%. Investigate coal feed rate and combustion efficiency to optimize performance."

    else:
        return f"{row['Parameter']} has a deviation of {deviation:.2f}%. Further investigation is recommended to maintain optimal operation."

combined_df["GPT_Recommendation"] = combined_df.apply(gpt_like_recommendation, axis=1)



def chatbot():
    print("\n🤖 Power Plant Chatbot (Boiler + Turbine): Ask me about any parameter!")
    print("Type 'exit' to quit.\n")

    combined_df["Parameter"] = combined_df["Parameter"].astype(str).fillna("")

    while True:
        user_input = input("🔍 Enter parameter name (or keyword): ").strip().lower()

        if user_input == "exit":
            print("Goodbye! 💡")
            break

        matches = combined_df[combined_df["Parameter"].str.lower().str.contains(user_input)]

        if matches.empty:
            print("❌ Sorry, I couldn't find that parameter. Try a different keyword.\n")
        else:
            for _, row in matches.iterrows():
                print(f"\n📌 Source: {row['Source']}")
                print(f"🔧 Parameter: {row['Parameter']}")
                print(f"📊 Deviation: {row['% Deviation']:.2f}%")
                print(f"🧠 Recommendation: {row['GPT_Recommendation']}\n")

chatbot()