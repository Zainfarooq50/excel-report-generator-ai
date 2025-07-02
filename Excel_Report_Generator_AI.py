import os
import pandas as pd
import numpy as np
from GPT_utils_excel import get_ai_insights


# ---------------- Main Excel Report Generator ----------------
def generate_excel_report(csv_file, output_folder="outputs"):
    """
    Generates a sales report with charts and AI-generated insights.
    """
    os.makedirs(output_folder, exist_ok=True)

    # Load Data
    try:
        df = pd.read_csv(csv_file)
    except FileNotFoundError:
        print(f"❌ File '{csv_file}' not found.")
        return

    print("✅ Data Loaded:\n", df.head())

    # Data Cleaning
    df["City"] = df["City"].str.title()
    df["Name"] = df["Name"].str.upper()

    # Performance Classification
    df["Performance"] = np.where(df["Sales"] > 6500, "High", "Low")

    # Summaries
    city_summary = df.groupby("City")["Sales"].sum().reset_index()
    perf_counts = df["Performance"].value_counts().reset_index()
    perf_counts.columns = ["Performance", "Count"]

    # Basic Text Summary
    summary_text = (
        f"Total Sales Records: {len(df)}\n"
        f"Total Sales Amount: ${df['Sales'].sum():,.2f}\n"
        f"High Performers: {perf_counts.loc[perf_counts['Performance'] == 'High', 'Count'].values[0]}\n"
        f"Low Performers: {perf_counts.loc[perf_counts['Performance'] == 'Low', 'Count'].values[0]}"
    )

    # AI Insights
    ai_text = get_ai_insights(summary_text)

    # Output File Path
    output_file = os.path.join(output_folder, f"AI_Sales_Report_{os.path.splitext(csv_file)[0]}.xlsx")

    # ---------------- Excel Report ----------------
    writer = pd.ExcelWriter(output_file, engine="xlsxwriter")

    df.to_excel(writer, sheet_name="All_Sales_Data", index=False)
    city_summary.to_excel(writer, sheet_name="City_Summary", index=False)
    perf_counts.to_excel(writer, sheet_name="Performance_Summary", index=False)

    workbook = writer.book
    ws_city = writer.sheets["City_Summary"]
    ws_perf = writer.sheets["Performance_Summary"]

    # Bar Chart - Sales by City
    bar_chart = workbook.add_chart({"type": "column"})
    bar_chart.add_series({
        "name": "Total Sales by City",
        "categories": ["City_Summary", 1, 0, len(city_summary), 0],
        "values": ["City_Summary", 1, 1, len(city_summary), 1],
    })
    bar_chart.set_title({"name": "Sales by City"})
    ws_city.insert_chart("D2", bar_chart)

    # Pie Chart - Performance Split
    pie_chart = workbook.add_chart({"type": "pie"})
    pie_chart.add_series({
        "name": "Performance Distribution",
        "categories": ["Performance_Summary", 1, 0, len(perf_counts), 0],
        "values": ["Performance_Summary", 1, 1, len(perf_counts), 1],
    })
    pie_chart.set_title({"name": "High vs Low Performers"})
    ws_perf.insert_chart("D2", pie_chart)

    # AI Insights Sheet
    ws_ai = workbook.add_worksheet("AI_Insights")
    ws_ai.write("A1", "Summary:")
    ws_ai.write("A2", summary_text)
    ws_ai.write("A5", "AI Generated Insights:")
    ws_ai.write("A6", ai_text)

    writer.close()
    print(f"\n✅ Full Report with Charts & AI Insights saved to '{output_file}'.")


# ---------------- Script Entry ----------------
if __name__ == "__main__":
    csv_file = input("Enter sales CSV file name (with .csv extension): ")
    generate_excel_report(csv_file)
