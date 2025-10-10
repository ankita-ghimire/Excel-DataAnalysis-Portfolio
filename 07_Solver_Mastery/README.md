# Excel Project: Operations Research Portfolio with Solver

**Project Status:** Completed

## Project Objective

This project's goal was to model and resolve a number of traditional, real-world "Operations Research" problems in order to show off proficiency with the Excel Solver Add-in.  This portfolio demonstrates the capacity to use Solver's optimization engine to identify the optimal solution after converting intricate business problems with numerous constraints into a mathematical model.

 This project goes beyond basic data reporting and into the field of **prescriptive analytics**, which uses data to suggest the best course of action in order to maximize profit, minimize expenses, and utilize scarce resources as efficiently as possible.
## Skills and Tools Demonstrated

This project is a comprehensive demonstration of advanced analytical and problem-solving skills using Excel's most powerful optimization tool.

*   **Excel Solver Add-in:**
    *   **Objective Setting:** Proficiently defined the objective function for each problem, whether it was to **Minimize** (cost), **Maximize** (profit/efficiency), or reach a specific target.
    *   **Variable Identification:** Correctly identified and set the "changing variable cells" that Solver is allowed to manipulate to find the optimal solution.
    *   **Constraint Modeling:** Mastered the ability to define a complex set of business rules and constraints, including:
        *   `<=` (e.g., not exceeding a supply or budget).
        *   `=` (e.g., meeting an exact demand).
        *   `>=` (e.g., ensuring non-negative results).
        *   **`binary`** constraints for "Yes/No" or "Go/No-Go" decisions, such as in the assignment and investment problems.
    *   **Solving Methods:** Correctly chose the appropriate solving method (**Simplex LP**) for linear optimization problems.

*   **Advanced Analytical Techniques:**
    *   **SUMPRODUCT Function:** Utilized `SUMPRODUCT` as the core function to build efficient and scalable objective functions and constraints.
    *   **Sensitivity Analysis:** Generated and interpreted a **Sensitivity Report**, demonstrating an advanced understanding of post-solution analysis. This includes explaining the business implications of key metrics like **Shadow Price** and **Reduced Cost**.

*   **Business Problem Modeling:**
    *   Successfully structured and solved a diverse portfolio of classic business challenges, proving the versatility of the skillset.

---

## Portfolio of Solved Problems

This workbook contains a separate sheet for each of the following optimization problems:

### 1. The Transportation Problem (Cost Minimization)
*   **Business Question:** Given a set of factories with limited supply and warehouses with specific demands, what is the cheapest possible shipping plan to satisfy all orders?
*   **Solver's Role:** To determine the exact quantity of goods to ship on each route to minimize the total transportation cost.
*   **Result:** Solver found the optimal shipping plan with a minimum total cost of **$3,570**, while ensuring all supply and demand constraints were perfectly met.

### 2. The Assignment Problem (Efficiency Maximization)
*   **Business Question:** Given a team of employees and a list of tasks, with each employee having a different skill score for each task, how do we assign exactly one person to each task to get the highest possible total efficiency?
*   **Solver's Role:** To find the perfect one-to-one matching of employees to tasks that maximizes the sum of the efficiency scores.
*   **Result:** Solver found the optimal assignment, achieving a maximum possible team efficiency score of **277**.

### 3. The Capital Investment Problem (Budget Optimization)
*   **Business Question:** Given a limited investment budget and a list of potential projects, each with a specific cost and expected return (NPV), which combination of projects should we fund to get the absolute highest total return?
*   **Solver's Role:** To use binary constraints (1=Invest, 0=Don't Invest) to select the optimal portfolio of projects.
*   **Result:** Solver identified the combination of Project B and Project C as the optimal investment strategy, maximizing the total return at **$165,000** while staying within the $120,000 budget.

### 4. Sensitivity Analysis
*   **Business Question:** After finding the cheapest transportation plan, how sensitive is our solution to change? How much would costs have to change to alter our plan? What's the value of increasing our factory supply?
*   **Solver's Role:** To generate a post-solution **Sensitivity Report**.
*   **Result:** The report provided key strategic insights, such as identifying a "Shadow Price" of -$3 for one of the factory supply constraints, indicating that each additional unit of supply from that factory would reduce total costs by $3.

---

## How to Use This Workbook

1.  Open the `Solver_Mastery_Portfolio.xlsx` file. Ensure the Solver Add-in is enabled (**Data -> Solver**).
2.  Navigate to any of the problem sheets (`Transportation`, `Assignment`, `Capital_Investment`).
3.  Open the Solver dialog box (**Data -> Solver**) to view the pre-configured objective, variables, and constraints for that problem.
4.  You can re-run the solution by clicking "Solve".
5.  Navigate to the `Sensitivity Report 1` sheet to view the post-solution analysis for the Transportation Problem.
