# Retirement Portfolio Monte Carlo Simulator

A comprehensive **Excel + VBA** application that uses historical bootstrap Monte Carlo simulation to evaluate retirement portfolio survival and growth.

## Features

- Historical bootstrap sampling from real S&P 500 and CPI data (Aswath Damodaran dataset)
- Block bootstrap option to better preserve market cycles
- Inflation-adjusted withdrawals
- Social Security integration (inflation-adjusted)
- Customizable stock/bond allocation and annual fees
- User-friendly graphical interface (VBA UserForm)
- Three visualizations:
  - Line chart of a single retirement path
  - Histogram showing distribution of ending balances
  - Summary bar chart of key outcomes

## How It Works

The simulator runs thousands of possible retirement scenarios by randomly sampling real historical market years.  
Each simulation applies actual past returns, inflation rates, and withdrawal needs to determine whether the portfolio survives the entire retirement period.

## Project Purpose

This project was developed as the final assignment for a VBA / Excel programming course.  
It demonstrates:
- Monte Carlo simulation techniques
- Historical bootstrap methodology
- Professional user interface design
- Data visualization and result communication

## Technologies Used

- Microsoft Excel
- VBA (Visual Basic for Applications)
- UserForms for GUI
- Dynamic Chart generation

## How to Use

1. Open the Excel workbook
2. Click the **"RUN SIMULATION"** button on the Inputs sheet (or run the macro)
3. Enter your assumptions in the UserForm
4. View results on the **Results**, **First Simulation Path**, and **Outcome Histogram** sheets

## Data Source

Historical data sourced from Aswath Damodaran’s dataset (S&P 500 Total Returns, 10-year Treasury Bonds, and CPI Inflation).

---

**Built as a semester-long final project** to demonstrate practical VBA application development, simulation modeling, and data visualization.
