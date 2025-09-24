# Excel Project: Interactive Sales Dashboard

## Project Objective

The objective of this project was to transform a raw sales dataset into a fully interactive and insightful dashboard. This project showcases a comprehensive mastery of data visualization in Excel, from initial data summarization with PivotTables to the creation of advanced, user-friendly reports with slicers. The goal was to build a tool that allows stakeholders to easily filter and analyze sales performance across different regions and product categories.

---

## Skills and Tools Demonstrated

This project demonstrates a wide range of intermediate to advanced Excel skills:

*   **Data Structuring:**
    *   Formatting raw data into a clean, usable format.
    *   Converting data ranges into named Excel Tables (`tblSalesData`) for dynamic referencing.

*   **Data Summarization:**
    *   Proficient use of **PivotTables** to aggregate and summarize data for analysis (e.g., sales over time, sales by category, sales by region).
    *   Changing value field settings (e.g., from `Sum` to `Average`) to derive correct insights.
    *   Manual and automatic grouping of date fields for time-series analysis.

*   **Data Visualization (Charting):**
    *   **Chart Selection:** Demonstrated the ability to choose the correct chart type for a specific analytical purpose:
        *   **Line Chart:** To visualize trends over time.
        *   **Column & Bar Charts:** For clear categorical comparisons and ranking.
        *   **Pie Chart:** To show compositional breakdowns (parts of a whole).
        *   **Combination (Combo) Chart:** To display two different data types (e.g., Sales in $ and Profit Margin in %) on a single, coherent chart using a secondary axis.
        *   **Advanced Custom Charts:** Created specialized visuals like a **Pareto Chart** for 80/20 analysis, a **Scatter Plot** to investigate correlations, and a custom **Gantt Chart** and **Thermometer/KPI Chart**.

*   **Dashboarding & Interactivity:**
    *   Assembling multiple charts and visuals into a single, cohesive dashboard layout.
    *   Implementing **Slicers** to create a dynamic and user-friendly filtering experience.
    *   Using **Report Connections** to link a single slicer to multiple PivotTables, ensuring the entire dashboard updates in unison.

*   **Data Analysis & Storytelling:**
    *   Adding analytical overlays like **Trendlines** to charts.
    *   Deriving actionable business insights from the visualized data, such as identifying top-performing product categories and regions or analyzing the relationship between sales volume and profitability.

---

## Key Project Steps

1.  **Data Preparation:** The raw dataset was cleaned, formatted, and converted into a named Excel Table.
2.  **Backend Creation:** A dedicated `PivotTables` sheet was created to serve as the "engine" for the dashboard. Four separate PivotTables were built to summarize the data from different perspectives.
3.  **Chart Construction:** Each required chart was built on a `Chart_Examples` sheet, ensuring each was correctly configured and formatted before being moved to the dashboard. This included troubleshooting common issues like data grouping and chart type limitations.
4.  **Dashboard Assembly:** The final, polished charts were moved and arranged onto the main `Dashboard` sheet.
5.  **Interactivity Implementation:** Slicers for "Region" and "Product Category" were added and connected to all PivotTables, making the dashboard fully interactive.
6.  **Final Touches:** The dashboard was given a professional look and feel by adding a title, setting a background color, and removing gridlines.

---

## How to Use This Dashboard

1.  Open the `Sales_Dashboard_Project.xlsx` file.
2.  Navigate to the **`Dashboard`** sheet.
3.  Use the slicers at the bottom of the dashboard to filter the data. Click on a Region or Product Category to see all charts update instantly.
4.  To select multiple items, hold down `Ctrl` (on Windows) or `Cmd` (on Mac) while clicking.
5.  To clear a filter, click the "Clear Filter" icon at the top-right of the slicer.
6.  The `PivotTables` sheet can be viewed to see the underlying data summaries that power the charts.
