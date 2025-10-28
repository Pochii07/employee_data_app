from PySide6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit,
    QPushButton, QFileDialog, QTableWidget, QTableWidgetItem, QHeaderView
)


class EmployeeDataUI(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Employee Data Generator")
        self.setGeometry(100, 100, 1000, 600)
        self.setMinimumSize(900, 550)

        self.main_layout = QHBoxLayout(self)

        left_layout = QVBoxLayout()

        self.num_employees_input = QLineEdit(self)
        self.num_employees_input.setPlaceholderText("Enter number of employees")
        left_layout.addWidget(self.num_employees_input)

        self.select_folder_button = QPushButton("Select Folder", self)
        left_layout.addWidget(self.select_folder_button)

        self.generate_button = QPushButton("Generate Data", self)
        left_layout.addWidget(self.generate_button)

        self.export_button = QPushButton("Export to Excel", self)
        self.export_button.setEnabled(False)
        left_layout.addWidget(self.export_button)

        self.message_label = QLabel("", self)
        left_layout.addWidget(self.message_label)
        left_layout.addStretch(1)

        self.table = QTableWidget(self)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)

        self.main_layout.addLayout(left_layout)
        self.main_layout.addWidget(self.table)
        self.main_layout.setStretch(0, 1) 
        self.main_layout.setStretch(1, 3)  