import sys
import os
import datetime
import pandas as pd
import re

def main():
    sales_csv = get_sales_csv()
    orders_dir = create_orders_dir(sales_csv)
    process_sales_data(sales_csv, orders_dir)

def get_sales_csv():
    num_params =len(sys.argv) - 1
    if num_params >= 1:
        csv_path = sys.argv[1]
        if os.path.isfile(csv_path):
            return os.path.abspath(csv_path)
        else:
            print("Error: CSV file does not exist.")
            sys.exit(1)
    else:
        print('Error: Missing CSV file path.')
        sys.exit(1)  

def create_orders_dir(sales_csv): 

    sales_dir = os.path.dirname(sales_csv)
    todays_date = datetime.date.today().isoformat()
    orders_dir_name = f'Orders_{todays_date}'
    orders_dir_path = os.path.join(sales_dir , orders_dir_name)
    if not os.path.isdir(orders_dir_path):
        os.makedirs(orders_dir_path)
    return orders_dir_path

def process_sales_data(sales_csv, orders_dir):
    sales_df = pd.read_csv(sales_csv)
    sales_df.insert(7, 'TOTAL PRICE', sales_df['ITEM QUANTITY'] * sales_df['ITEM PRICE'])  
    sales_df.drop(columns=['ADDRESS', 'CITY', 'STATE', 'POSTAL CODE', 'COUNTRY'], inplace=True)
    for order_id, order_df in sales_df.groupby('ORDER ID'):
        
        order_df.drop(columns=['ORDER ID'], inplace=True)
        order_df.sort_values(by='ITEM NUMBER', inplace=True)
        grand_total = order_df['TOTAL PRICE'].sum()
        grand_total_df = pd.DataFrame({'ITEM PRICE': ['GRAND TOTAL:'], 'TOTAL PRICE':[grand_total]})
        order_df = pd.concat([order_df, grand_total_df])

        customer_name = order_df['CUSTOMER NAME'].values[0]
        customer_name = re.sub(r'\W', '', customer_name)
        order_file_name = f'Orders{order_id}_{customer_name}.xlsx'
        order_file_path = os.path.join(orders_dir, order_file_name)
        
        sheet_name = f'Order {order_id}'
        writer = pd.ExcelWriter(order_file_path, engine='xlsxwriter')
        order_df.to_excel(writer, sheet_name=sheet_name)
        workbook  = writer.book
        worksheet = writer.sheets[sheet_name]
        format1 = workbook.add_format({'num_format': '$#,##0.00'}) 
        worksheet.set_column('B:B',11)
        worksheet.set_column('G:H',13, format1)
        worksheet.set_column('D:F',15)

        worksheet.set_column('C:C',13)
        worksheet.set_column('I:I',10)
        worksheet.set_column('J:J',30)
        
        writer.close()
        


if __name__ == '__main__':
    main()    