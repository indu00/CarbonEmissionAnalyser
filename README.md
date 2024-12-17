# CarbonEmissionAnalyser
A Python-based application to calculate, analyze, and visualize carbon emissions from energy consumption, waste generation, and business travel. This tool helps users monitor their environmental impact by storing data in Excel files and providing detailed insights through tabular reports, trend analysis.


## Features
1. **Data Input:**
   - Users can input data for:
     - Monthly electricity, natural gas, and fuel bills.
     - Monthly waste generation and recycling percentage.
     - Distance traveled for business purposes and fuel efficiency.
   - Automatic calculations for carbon emissions from these categories.

2. **Data Storage:**
   - Data is saved in an Excel file (`UserData.xlsx`) for future reference using the `openpyxl` library.

3. **Visualization:**
   - View data trends over time using line charts.
   - Analyze overall carbon emissions with a pie chart.

4. **Tabular Display:**
   - View all user data in a neatly formatted table using the `tabulate` library.

5. **Trends Analysis:**
   - Highlight users with the highest emissions.
   - Identify which category (Energy, Waste, or Travel) contributes the most.

---


### Setup
1. Clone the repository:
   ```bash
   git clone https://github.com/indu00/CarbonEmissionAnalyser.git
   ```
2. Navigate to the directory:
   ```bash
   cd carbon-footprint-tracker
   ```
## How to Use
1. Run the program:
   ```bash
   python tracker.py
   ```
2. Follow the prompts to input data.
3. Use the menu to:
   - Input additional data.
   - View stored data in tabular format.
   - Analyze trends with line..
4. Exit the application when finished.
   
---

### Tabular Data
Displays data like this:
Data comes from UserData.xlsx file
```
╒════════════╤═══════════╤════════════════╤════════════════════════════╤═══════════════════════════╤══════════════════════════╕
│ Email      │ Username  │ Company Name   │ Carbon Emission (Energy)   │ Carbon Emission (Waste)   │ Carbon Emission (Travel) │
╘════════════╧═══════════╧════════════════╧════════════════════════════╧═══════════════════════════╧══════════════════════════╛
```
