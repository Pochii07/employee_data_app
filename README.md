# Employee Data App

This project is a simple desktop application that allows users to generate synthetic employee data and export it to an Excel file. The application is built using PySide6 for the GUI, pandas for data handling, and Faker for realistic name generation.

## Features

- Input the number of employees to generate.
- Select a folder to save the generated Excel file.
- Generate synthetic employee data with the following fields:
  - Employee ID (auto-incrementing)
  - Full Name (random realistic name)
  - Department (randomly chosen from IT, HR, Operations, Administration, Finance)
  - Salary (random number between 25,000 and 120,000)
  - Hire Date (random date between 2020-01-01 and today)
- Export the generated data to an Excel file named `employees.xlsx`.
- Include a summary sheet showing the average salary per department.
- Display messages to the user regarding the status of operations.

## Requirements

To run this application, you need to install the following dependencies:

- PySide6
- pandas
- faker
- openpyxl

You can install the required packages using pip. Make sure to create a virtual environment for better package management.

## Installation

1. Clone the repository:
   ```
   git clone <repository-url>
   cd employee-data-app
   ```

2. Install the required packages:
   ```
   pip install -r requirements.txt
   ```

## Running the Application

To run the application, execute the following command in your terminal:

```
python employee_app.py
```

## Usage

1. Enter the number of employees you want to generate in the input field.
2. Click the "Select Folder" button to choose a directory where the Excel file will be saved.
3. Click the "Generate Data" button to create the synthetic employee data.
4. Once the data is generated, click the "Export to Excel" button to save the data to an Excel file.