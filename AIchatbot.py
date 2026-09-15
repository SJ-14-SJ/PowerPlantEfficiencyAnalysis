import pandas as pd
from analysis import ROOT, load_measurements, deviation_percent

def rule_based_recommendation(row):
    if pd.isna(row["% Deviation"]):
        return "Insufficient data or zero design baseline"
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





def chatbot():
    frames = []
    for source, filename in [("Boiler", "Boiler.xlsx"), ("Turbine", "project.xlsx")]:
        frame = load_measurements(ROOT / filename, source)
        frame["Source"] = source
        frames.append(frame)
    combined_df = pd.concat(frames, ignore_index=True)
    combined_df["% Deviation"] = combined_df.apply(lambda r: deviation_percent(r["Performance"], r["Design"]), axis=1)
    combined_df["Recommendation"] = combined_df.apply(rule_based_recommendation, axis=1)
    print("\n🤖 Power Plant Chatbot (Boiler + Turbine): Ask me about any parameter!")
    print("Type 'exit' to quit.\n")

    combined_df["Parameter"] = combined_df["Parameter"].astype(str).fillna("")

    while True:
        user_input = input("🔍 Enter parameter name (or keyword): ").strip().lower()

        if user_input == "exit":
            print("Goodbye! 💡")
            break

        if not user_input:
            print("Enter a parameter name or keyword.")
            continue

        matches = combined_df[combined_df["Parameter"].str.lower().str.contains(user_input, regex=False, na=False)]

        if matches.empty:
            print("❌ Sorry, I couldn't find that parameter. Try a different keyword.\n")
        else:
            for _, row in matches.iterrows():
                print(f"\n📌 Source: {row['Source']}")
                print(f"🔧 Parameter: {row['Parameter']}")
                print(f"📊 Deviation: {row['% Deviation']:.2f}%")
                print(f"🧠 Recommendation: {row['Recommendation']}\n")

if __name__ == "__main__":
    chatbot()
