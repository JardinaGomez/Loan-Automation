# Loan Management Portfolio

## Overview

This Loan Management Portfolio is an automated reporting tool designed in Excel VBA to efficiently process and categorize loan data. It accepts raw loan data in Excel or CSV format and generates comprehensive summaries, categorizing loans based on their due status such as Past Due, Non-Accrual, Matured, and more.

## Features

- **Categorization:** Organizes loans into distinct sections based on their status
  - 1-30 Days

  - 30 Days

  - 60 Days

  - Over 90 Days

  - Past Due

  - Matured

  - Non-Accrual

- Generates comprehensive summary reports, including:

  - Total Principal Balance

  - Amount Past Due

  - Exposure

  - % Net Portfolio

- **Data-driven approach:** Accepts raw loan data from Excel (.xlsx) or CSV formats.

- **Dynamic Reporting:** Automatically calculates and updates metrics, such as total exposure and percentage of the net portfolio for each loan category.

# Requirements

## Data Input

- Input files should be Excel or CSV formatted.

- Required headers in input data:

  - Company Name

  - Loan Nbr

  - Next Payment Due Date

  - Note Date

  - G/L Balance

  - USER18N

  - Maturity Date

  - Risk Factor

  - Int Paid to Date

  - sba Guarantee Percentage

  - Part Id

  - Non-Accrual Date

# Instructions
**1. Clone the repository:**
  > git clone 

**2.  Setup your Workbook:**
   - Make sure the workbook contains the folowing sheets before execution:
       - Data
       - Button
  
**3. **Running the Report:**
  1. Open the Loan Management Portfolio.xlsm Excel workbook.

  2. Import your data : Click the "Import Data" button to imput you data file
![Screenshot 2025-03-09 at 5 04 09 PM](https://github.com/user-attachments/assets/da7a1ed2-4ed2-454c-b3ca-c6ad6ddd6f32)


  4. Ensure only "Data" and "Button" sheets are present.
       - If not press the "RESET" button 

  5. Use the provided dropdown or manually run the RUNALL() macro to automatically create a report:
  ![image](https://github.com/user-attachments/assets/b75e3de4-c6b5-48f6-a7d6-25d11db7701d)

  The final report will .. 
    - Create categorized sections based on delinquency.
    - Populate and calculate values for the Summary box (including % Net Portfolio).

# Important Instructions

- Before running the RUNALL() macro, ensure:

  - The imported data file contains the required headers.

  - All prior data sheets (except Button) have been cleared or deleted.
  - 
## Troubleshooting
- For questions, bugs, or enhancement requests, please contact Jardina directly through internal github. 

# Collaborations

- Developed collaboratively with the Pursuit Team to enhance reporting efficiency and accuracy.

- Regular feedback was provided by stakeholders at Pursuit to align the tool with operational requirements.

# Usage Guidelines

- Ensure input data accuracy to guarantee correct categorization.

- Regularly backup reports prior to generating new data.

- Verify macro-enabled Excel workbook (.xlsm) settings before execution.

# Technical Specs 

- **Programming Language:** Excel VBA

- **File Formats:** Excel Macro-enabled Workbook (.xlsm), CSV


# Future Development

- Plans for further automation include:

  - Additional reports: Concentration Report, Credit Reporting.

  - Integration of more robust data validation and error handling.

# Maintainer

- Jardina Gomez

- For questions or support, please contact Jardina directly through github.

