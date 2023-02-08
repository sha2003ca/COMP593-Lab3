import sys
import openpyxl
import os
from openpyxl import Workbook
from datetime import date
import pandas as pd
def main():
    sales_csv = get_sales_csv()
    orders_dir = create_orders_dir(sales_csv)
    process_sales_data(sales_csv, orders_dir)
# Get path of sales data CSV file from the command line
def get_sales_csv():
    if len(sys.argv) < 2:
        print("Error: No command line parameter provided.")
        sys.exit()
    csv_file = sys.argv[1]
    if not os.path.exists(csv_file):
        print("Error: Provided file path does not exist.")
        sys.exit()
    return csv_file
# Create the directory to hold the individual order Excel sheets
def create_orders_dir(sales_csv):
    csv_dir = os.path.dirname(sales_csv)
    today = date.today().strftime("%Y-%m-%d")
    orders_dir = os.path.join(csv_dir, "Orders_" + today)
    if not os.path.exists(orders_dir):
        os.mkdir(orders_dir)
    return orders_dir
# Split the sales data into individual orders and save to Excel sheets
def process_sales_data(sales_csv, orders_dir):
    sales_data = pd.read_csv(sales_csv)
    sales_data['TOTAL PRICE'] = sales_data['ITEM QUANTITY'] * sales_data['ITEM PRICE']
    sales_data = sales_data[['ORDER DATE', 'ITEM NUMBER', 'PRODUCT LINE', 'PRODUCT CODE', 'ITEM QUANTITY', 'ITEM PRICE', 'TOTAL PRICE', 'STATUS', 'CUSTOMER NAME']]
    orders = sales_data.groupby(by='ORDER ID')
    for order_id, order_data in orders:
        order_data = order_data.sort_values(by='ITEM NUMBER')
        order_data = order_data.reset_index(drop=True)
        grand_total = order_data['TOTAL PRICE'].sum()
        order_data = order_data.append(
            {'ITEM NUMBER': 'GRAND TOTAL', 'TOTAL PRICE': grand_total}, ignore_index=True)
        order_file = os.path.join(orders_dir, str(order_id) + ".xlsx")
        order_data.to_excel(order_file, index=False, engine='openpyxl')
        # TODO: Format the Excel sheet
        wb = openpyxl.load_workbook(order_file)
        sheet = wb.worksheets[0]
        # Format the Excel sheet
        sheet.column_dimensions['A'].width = 11
        sheet.column_dimensions['B'].width = 13
        sheet.column_dimensions['C'].width = 15
        sheet.column_dimensions['D'].width = 15
        sheet.column_dimensions['E'].width = 15
        sheet.column_dimensions['F'].width = 13
        sheet.column_dimensions['G'].width = 13
        sheet.column_dimensions['H'].width = 10
        sheet.column_dimensions['I'].width = 30
        for row in sheet.iter_rows(min_row=2, min_col=4, max_col=5):
            for cell in row:
                cell.number_format = '"$"#,##0.00'
        wb.save(order_file)
if __name__ == '__main__':
    main()