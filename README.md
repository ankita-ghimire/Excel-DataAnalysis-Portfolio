Title: Sales Data — Advanced Sorting & Priority Analysis

Problem: Demonstrate thorough Excel sorting techniques to prioritize sales actions and discover top performers.

Data: Synthetic Employee_Sales dataset (20 rows): EmpID, FirstName, LastName, Department, Country, Quarter, Priority, Sales.

Steps:
1. Created OriginalOrder and converted data to a Table.
2. Cleaned Sales and trimmed text.
3. Demonstrated single-column sorts (LastName, Sales).
4. Performed multi-level sort: Country → Priority (Custom List) → Sales (desc).
5. Created Custom Lists for Priority and Quarter (Critical > High > Medium > Low; Q1->Q4).
6. Applied cell color formatting and sorted by Cell Color.
7. Built PriorityRank helper column and used it for numeric sorting.
8. Created a dynamic sorted view (SORT/SORTBY) — if Excel 365 available.
9. Built PivotTable to show Top 3 employees by Sales.

Key insights:
- Custom lists and helper columns are reliable for non-alphabetical orders.
- Sorting by Sales quickly reveals top earners; multi-level sort groups data for targeted action.
- Pivot Top-N finds key employees per department.

Files included:
- Sales_Sorting_Project.xlsx
- custom_list.png
- README.md
