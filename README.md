# KPI Dashboard (Streamlit)

This project is a robust, production-ready Streamlit dashboard for analyzing employee performance KPIs across multiple teams (Sales, Ads, Website Ads, Portfolio Holders). It provides dynamic filtering, clear metric displays, and advanced logic for identifying top performers and the "Star of the Month."

## Features
- **Dynamic Sidebar Filters:** Filter by Team, Month, Week, and Employee.
- **Employee Performance Cards:** View detailed metrics for each employee, including handling of missing/unknown names.
- **Metric Cards & Trend Charts:** Visualize weekly and monthly sales, targets, and performance trends.
- **Star of the Month Logic:** Automatically highlights employees with at least one perfect score (10) in the month, and among them, the one(s) with the highest total net sales.
- **Team-Agnostic:** Supports all team types and their unique metrics via robust column mapping and cleaning logic.
- **Production-Ready:** Clean code, type safety, and no debug prints.

## Usage
1. Upload your KPI Excel file via the sidebar.
2. Use the filters to explore performance by team, month, week, or employee.
3. Review the dashboard for sales breakdowns, trends, and top performers.

## Requirements
- Python 3.8+
- streamlit
- pandas
- plotly

Install dependencies:
```bash
pip install streamlit pandas plotly
```

## Running the Dashboard
```bash
streamlit run KPI_DASH.py
```

## GitHub Deployment
- This repository is managed via git. To update, commit your changes and push to the repo.

## Author
- digitwebai

---
For any issues or feature requests, please open an issue on GitHub.
