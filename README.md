# Monte Carlo Retirement Portfolio Simulator

A comprehensive Excel + VBA application that uses historical bootstrap Monte Carlo simulation to predict retirement portfolio success rates and ending balances.

Features

-Historical Bootstrap Sampling with optional block bootstrap (preserves market cycles)
-Inflation-adjusted withdrawals using actual historical CPI data
-Social Security integration (inflation-adjusted, user-defined start year)
-Flexible portfolio modeling (stock/bond allocation + annual fees)
-Three visualizations:
-Line chart showing one full retirement path
-Histogram of all ending portfolio outcomes (in $1M ranges)
-Summary bar chart highlighting key metrics (Median, 5th percentile, Average)

User-friendly GUI (VBA UserForm) – no need to edit cells directly
Uses real historical data from Aswath Damodaran (S&P 500 total returns, 10-year T.Bonds, and CPI)

How It Works

The simulator runs thousands of possible retirement scenarios by randomly sampling real historical market years (with replacement). Each simulation applies actual past returns, inflation rates, and withdrawal needs to test whether the portfolio survives the full retirement period.
Project Purpose
This was developed as a final project for a VBA / Excel programming course. It demonstrates:

Monte Carlo simulation techniques
Historical bootstrap methodology
User interface design (UserForms)
Data visualization and result interpretation
Professional Excel application development

Technologies Used

Microsoft Excel + VBA
UserForms for GUI
Chart objects for visualization
Historical financial data (Damodaran dataset)
