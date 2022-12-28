import sqlite3
from datetime import date
from openpyxl import Workbook, load_workbook
import pandas as pd


def main():
    main_menu = int(input('''What do you want to do?
1 - add expense 
2 - export expenses to Excel
3 - exit   
Please enter the number:  '''))

    while True:
        try:
            if main_menu == 1:
                Expense.add_expense()
            if main_menu == 2:
                export_to_excel()
            elif main_menu == 3:
                print("Exiting the program...")
                break
        finally:
            break


def export_to_excel():
    while True:
        try:
            wb = Workbook()
            ws = wb.active
            wb.save("C:\\Users\\hp store\\OneDrive\\Рабочий стол\\Python Assignment 2!\\2022-11-20.xlsx")
            connection = sqlite3.connect(
                "C:\\Users\\hp store\\OneDrive\\Рабочий стол\\Python Assignment 2!\\Expenses.sqlite")
            connection.execute("PRAGMA foreign_keys = ON")
            connection.commit()
            cursor = connection.cursor()
            cursor.execute('SELECT Amount FROM Expense')
            items = cursor.fetchall()
            filepath = "C:\\Users\\hp store\\OneDrive\\Рабочий стол\\Python Assignment 2!\\2022-11-20.xlsx"
            cursor.execute('SELECT Name from Category, Expense WHERE Category.CategoryID = Expense.CategoryId')
            items2 = cursor.fetchall()
            cursor.execute('SELECT Date from Expense')
            items3 = cursor.fetchall()
            connection.commit()
            table = {"Amount": items, "Name": items2, "Date": items3}
            dataframe = pd.DataFrame(table)
            dataframe.to_excel(filepath)

        finally:
            print("Expenses have been exported to Excel, the file name is the current date.")
            return main()


class Expense():

    def __init__(self):
        """For the sake of initializing"""

    def add_expense():
        while True:
            try:

                connection = sqlite3.connect(
                    "C:\\Users\\hp store\\OneDrive\\Рабочий стол\\Python Assignment 2!\\Expenses.sqlite")
                cursor = connection.cursor()
                cursor.execute('SELECT * FROM Category')
                items = cursor.fetchall()
                today = date.today()
                for item in items:
                    print(item[0], "-", item[1])
                user_choice = int(input('To which category belongs this expense. Choose the number:  '))
                amount = float(input("What is the amount of the expense?: "))
                expenseid = None
                cursor.execute('''INSERT INTO Expense VALUES(?,?,?,?)''', (expenseid, amount, user_choice, today))
                connection.commit()
            finally:
                print("Expense is added!")
                return main()


main()
