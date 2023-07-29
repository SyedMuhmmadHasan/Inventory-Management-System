import sys
import os
import csv
import sqlite3
import pandas as pd
from PyQt5.QtWidgets import QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, \
    QPushButton, QMessageBox, QTableWidget, QTableWidgetItem, QFileDialog


class InventorySystem(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Inventory Management System")
        self.setGeometry(100, 100, 600, 400)

        # Create a new database file "inventory.db" if it doesn't exist
        if not os.path.exists("inventory.db"):
            self.conn = sqlite3.connect("inventory.db")
            self.cursor = self.conn.cursor()
            self.cursor.execute('''CREATE TABLE IF NOT EXISTS items
                                  (id INTEGER PRIMARY KEY AUTOINCREMENT,
                                   name TEXT NOT NULL,
                                   quantity INTEGER NOT NULL)''')
            self.conn.commit()
        else:
            self.conn = sqlite3.connect("inventory.db")
            self.cursor = self.conn.cursor()

        self.init_ui()

        # Load items from database on startup
        self.load_items()

    def init_ui(self):
        # Widgets
        self.item_name_label = QLabel("Item Name:")
        self.item_name_input = QLineEdit()
        self.quantity_label = QLabel("Quantity:")
        self.quantity_input = QLineEdit()
        self.add_button = QPushButton("Add Item")
        self.delete_button = QPushButton("Delete Item")
        self.remove_all_button = QPushButton("Remove All")  # New button: Remove All
        self.save_button = QPushButton("Save List")

        # Table to display items
        self.table = QTableWidget()
        self.table.setColumnCount(2)  # Removed the ID column
        self.table.setHorizontalHeaderLabels(["Item Name", "Quantity"])  # Removed the ID label

        # Layout setup
        input_layout = QHBoxLayout()
        input_layout.addWidget(self.item_name_label)
        input_layout.addWidget(self.item_name_input)
        input_layout.addWidget(self.quantity_label)
        input_layout.addWidget(self.quantity_input)
        input_layout.addWidget(self.add_button)
        input_layout.addWidget(self.delete_button)
        input_layout.addWidget(self.remove_all_button)  # Added the Remove All button
        input_layout.addWidget(self.save_button)

        main_layout = QVBoxLayout()
        main_layout.addLayout(input_layout)
        main_layout.addWidget(self.table)

        central_widget = QWidget()
        central_widget.setLayout(main_layout)
        self.setCentralWidget(central_widget)

        # Button connections
        self.add_button.clicked.connect(self.add_item)
        self.delete_button.clicked.connect(self.delete_item)
        self.remove_all_button.clicked.connect(self.remove_all_items)  # Connected the button to the method
        self.save_button.clicked.connect(self.save_list)

        # Apply style to the widgets
        self.setStyleSheet("""
            QLabel {
                font-size: 14px;
                font-weight: bold;
            }
            QLineEdit, QTableWidget {
                font-size: 14px;
            }
            QPushButton {
                font-size: 14px;
                background-color: #4CAF50;
                color: white;
                border: none;
                padding: 10px;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
        """)

    def add_item(self):
        name = self.item_name_input.text().strip()
        quantity = self.quantity_input.text().strip()

        if name and quantity:
            try:
                quantity = int(quantity)

                # Check if the item name is already in the list
                if name in [self.table.item(row, 0).text() for row in range(self.table.rowCount())]:
                    self.show_message("Duplicate Item", "An item with the same name already exists.")
                else:
                    self.cursor.execute("INSERT INTO items (name, quantity) VALUES (?, ?)", (name, quantity))
                    self.conn.commit()
                    self.load_items()
                    self.item_name_input.clear()
                    self.quantity_input.clear()
            except ValueError:
                self.show_message("Invalid input", "Quantity must be an integer.")
        else:
            self.show_message("Incomplete Information", "Please enter item name and quantity.")

    def delete_item(self):
        selected_rows = self.table.selectionModel().selectedRows()

        if selected_rows:
            confirm = QMessageBox.question(self, "Confirmation", "Are you sure you want to delete the selected item?",
                                           QMessageBox.Yes | QMessageBox.No)
            if confirm == QMessageBox.Yes:
                name = self.table.item(selected_rows[0].row(), 0).text()
                self.cursor.execute("DELETE FROM items WHERE name=?", (name,))
                self.conn.commit()
                self.load_items()
        else:
            self.show_message("No Item Selected", "Please select an item to delete.")

    def remove_all_items(self):
        confirm = QMessageBox.question(self, "Confirmation", "Are you sure you want to remove all items?",
                                       QMessageBox.Yes | QMessageBox.No)
        if confirm == QMessageBox.Yes:
            self.cursor.execute("DELETE FROM items")
            self.conn.commit()
            self.load_items()

    def save_list(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        file_name, _ = QFileDialog.getSaveFileName(self, "Save List as Excel", "", "Excel Files (*.xlsx);;All Files (*)", options=options)

        if file_name:
            # Check if the file name ends with '.xlsx' and add it if missing
            if not file_name.lower().endswith('.xlsx'):
                file_name += '.xlsx'

            # Convert the table data into a pandas DataFrame
            data = []
            for row in range(self.table.rowCount()):
                name_val = self.table.item(row, 0).text()
                quantity_val = self.table.item(row, 1).text()
                data.append([name_val, quantity_val])

            df = pd.DataFrame(data, columns=["Item Name", "Quantity"])

            # Save the DataFrame to an Excel file with adjusted column width
            try:
                with pd.ExcelWriter(file_name, engine='xlsxwriter') as writer:
                    df.to_excel(writer, index=False, sheet_name="Inventory")

                    # Access the XlsxWriter workbook and worksheet objects
                    workbook = writer.book
                    worksheet = writer.sheets["Inventory"]

                    # Set the column width for the "Item Name" column
                    worksheet.set_column("A:A", 17)  # Adjust the value (20) to increase or decrease the width

                    # Set the column width for the "Quantity" column
                    worksheet.set_column("B:B", 10)  # Adjust the value (9) to increase or decrease the width

                self.show_message("Success", "List saved successfully.")
            except Exception as e:
                self.show_message("Error", f"Failed to save the list. Error: {str(e)}")

    def load_items(self):
        self.table.setRowCount(0)
        self.cursor.execute("SELECT * FROM items")
        rows = self.cursor.fetchall()

        for row_index, row_data in enumerate(rows):
            self.table.insertRow(row_index)
            for column_index, data in enumerate(row_data[1:]):  # Exclude the ID column
                item = QTableWidgetItem(str(data))
                self.table.setItem(row_index, column_index, item)

    def show_message(self, title, message):
        msg_box = QMessageBox(self)
        msg_box.setWindowTitle(title)
        msg_box.setText(message)
        msg_box.exec_()

    def closeEvent(self, event):
        self.conn.close()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = InventorySystem()
    window.show()
    sys.exit(app.exec_())
