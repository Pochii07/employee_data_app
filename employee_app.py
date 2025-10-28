import sys
import os
import random
import pandas as pd
from faker import Faker
from datetime import datetime, date

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QApplication, QFileDialog, QTableWidgetItem

from employee_app_ui import EmployeeDataUI


class EmployeeDataApp(EmployeeDataUI):
    def __init__(self):
        super().__init__()
        self.folder_path = ""
        self.employee_data = pd.DataFrame()

        self.select_folder_button.clicked.connect(self.select_folder)
        self.generate_button.clicked.connect(self.generate_data)
        self.export_button.clicked.connect(self.export_to_excel)

    def select_folder(self):
        self.folder_path = QFileDialog.getExistingDirectory(self, "Select Folder")
        if self.folder_path:
            self.message_label.setText("Folder selected: " + self.folder_path)
        else:
            self.message_label.setText("No folder selected.")

    def generate_data(self):
        try:
            num_employees = int(self.num_employees_input.text())
            if num_employees <= 0:
                raise ValueError("Number of employees must be greater than 0.")

            fake = Faker()
            departments = ["IT", "HR", "Operations", "Administration", "Finance"]
            data = []

            start_date = datetime(2020, 1, 1).date()
            end_date = datetime.now().date()

            for emp_id in range(1, num_employees + 1):
                full_name = fake.name()
                department = random.choice(departments)
                salary = random.randint(25000, 120000)
                hire_date = fake.date_between(start_date=start_date, end_date=end_date)
                data.append([emp_id, full_name, department, salary, hire_date])

            self.employee_data = pd.DataFrame(
                data,
                columns=["emp_id", "full_name", "department", "salary", "hire_date"]
            )

            self.populate_table(self.employee_data)
            self.message_label.setText(f"{num_employees} records generated.")
            self.export_button.setEnabled(True)

        except ValueError as e:
            self.message_label.setText(str(e))

    def populate_table(self, df: pd.DataFrame):
        self.table.clear()
        self.table.setRowCount(len(df))
        self.table.setColumnCount(len(df.columns))
        self.table.setHorizontalHeaderLabels(df.columns.tolist())

        for r in range(len(df)):
            for c in range(len(df.columns)):
                val = df.iat[r, c]
                if isinstance(val, (datetime, date)):
                    text = val.strftime("%Y-%m-%d")
                else:
                    text = str(val)
                item = QTableWidgetItem(text)
                item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                self.table.setItem(r, c, item)

        self.table.setSortingEnabled(True)

    def _unique_export_path(self, base_filename: str = "employees.xlsx") -> str:
        path = os.path.join(self.folder_path, base_filename)
        if not os.path.exists(path):
            return path
        base, ext = os.path.splitext(path)
        i = 1
        while os.path.exists(f"{base}_{i}{ext}"):
            i += 1
        return f"{base}_{i}{ext}"

    def export_to_excel(self):
        if self.employee_data.empty:
            self.message_label.setText("No data to export. Generate data first.")
            return
        if not self.folder_path:
            self.message_label.setText("No folder selected.")
            return

        def write_workbook(writer: pd.ExcelWriter):
            self.employee_data.to_excel(writer, sheet_name='Employees', index=False)
            summary = (
                self.employee_data.groupby('department', as_index=False)['salary']
                .mean()
                .rename(columns={'salary': 'average_salary'})
            )
            summary['average_salary'] = summary['average_salary'].round(2)
            summary.to_excel(writer, sheet_name='Summary', index=False)
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            writer.sheets['Summary'].cell(row=1, column=3, value=f"Exported on: {timestamp}")

        dest_path = self._unique_export_path("employees.xlsx")
        try:
            with pd.ExcelWriter(dest_path, engine='openpyxl') as writer:
                write_workbook(writer)
            self.message_label.setText("File generated: " + dest_path)
        except PermissionError:
            ts_path = os.path.join(
                self.folder_path,
                f"employees_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            )
            try:
                with pd.ExcelWriter(ts_path, engine='openpyxl') as writer:
                    write_workbook(writer)
                self.message_label.setText("Target file in use. Saved as: " + ts_path)
            except PermissionError:
                self.message_label.setText("Permission denied. Close the Excel file and try again.")
        except Exception as e:
            self.message_label.setText(f"Export failed: {e}")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = EmployeeDataApp()
    window.show()
    sys.exit(app.exec())