import threading
from flask import Flask, render_template, request, redirect
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pymysql
import time
from crm import *

app = Flask(__name__)

# Define your Google Sheets credentials
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name('mailer-400406-83227f4a1b2d.json', scope)
client = gspread.authorize(creds)


@app.route('/')
def crm():
    return render_template('crm.html')


@app.route('/fetch', methods=['POST'])
def fetch_data():
    if request.method == 'POST':
        item_no = request.form['item_no']
        style = request.form['style']
        gs_link = request.form['gs_link']

        try:
            sheet = client.open_by_url(gs_link).worksheet('Sheet4')
        except gspread.exceptions.SpreadsheetNotFound:
            return "Sheet not found"

        # Search for the item number in the cell B3
        cell = sheet.acell('B3')
        if cell.value == item_no:
            # Fetch all data from the Google Sheet
            excel_data = sheet.get_all_records(head=11)  # Skip the first 10 rows

            # Connect to MySQL database
            conn = pymysql.connect(
                host='172.20.50.169',
                user='internal_apps',
                password='Schemax@2023',
                database='internal_apps',
                charset='utf8mb4',
                cursorclass=pymysql.cursors.DictCursor
            )

            cursor = conn.cursor()
            # Insert fetched data into the database
            for row in excel_data:
                cursor.execute('''INSERT INTO crm_fetched_data (`item_no`, `group`, cost_group, item_type, item_group, product_group, operation,  fabric_type, structure, quality, finish, 
                                                material_code, name, description, budget_price, currency, unit_of_measure, alternate_unit_of_measure,
                                                consumption, waste, factor, development_responsible, placement, mapping, status) VALUES (
                                                %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s
                                            )''', (
                    item_no, row['Grp'], row['Cost Grp'], row['Item Type'], row['Item Grp'],
                    row['Product Grp'], row['Operation'], row['Fabric Type'], row['Structure'], row['Quality'],
                    row['Finish'], row['Material Code'], row['Name'],
                    row['Description'], row['Budget Price'], row['Currency'], row['UOM'], row['ALT UOM'],
                    row['Consp.'], row['Waste'], row['Factor'], row['Development Responsible'], row['Placement'],
                    row['MAPPING'], 'Pending'  # Assuming 'Pending' is the initial status
                ))

            # Commit changes and close connection
            conn.commit()  # Commit after each insert
            cursor.close()
            conn.close()  # Close the connection after all inserts
            return redirect('/result')  # Redirect to the result page after fetching data


@app.route('/result')
def show_result():
    # Fetch data from the database to display in the result page
    conn = pymysql.connect(
        host='172.20.50.169',
        user='internal_apps',
        password='Schemax@2023',
        database='internal_apps',
        charset='utf8mb4',
        cursorclass=pymysql.cursors.DictCursor
    )
    cursor = conn.cursor()
    cursor.execute("SELECT id, item_no, `group`, product_group, item_group, description, operation, fabric_type, structure, quality, finish, material_code, unit_of_measure, alternate_unit_of_measure, currency, budget_price, consumption, waste, factor, development_responsible, placement, status FROM crm_fetched_data")
    fetched_data = cursor.fetchall()
    cursor.close()
    conn.close()
    return render_template('result.html', data=fetched_data)  # Pass fetched data to the template


# New route to run the bot after clicking "CREATE" button in result.html
@app.route('/run_bot', methods=['POST'])
def run_bot():
    if request.method == 'POST':
        fetched_data = database()
        # # Iterate over fetched_data and extract item_no and group
        for data in fetched_data:
            id = data['id']
            item_no = data['item_no']
            group = data['group']
            product_group = data['product_group']
            item_group = data['item_group']
            description = data['description']
            operation = data['operation']
            fabric_type = data['fabric_type']
            structure = data['structure']
            quality = data['quality']
            finish = data['finish']
            material_code = data['material_code']
            unit_of_measure = data['unit_of_measure']
            alternate_unit_of_measure = data['alternate_unit_of_measure']
            factor = data['factor']
            development_responsible = data['development_responsible']
            currency = data['currency']
            budget_price = data['budget_price']
            consumption = data['consumption']
            waste = data['waste']
            placement = data['placement']
            status = data['Status']
            print("Data from database:")
            print(data)
            # # Skip running item_field if group is 'FAB' or empty
            # if group in ('FAB', ''):
            #     continue
            if status != "Created":
                # Check if Status is 'Pending'
                print("Calling item_field for ID:", id)
                # Call the item_field function with the fetched item_no and group
                item_field(id, item_no, group, product_group, item_group, description, operation, unit_of_measure,
                           alternate_unit_of_measure,
                           factor, currency, budget_price, consumption, waste, placement, material_code, fabric_type,
                           structure, quality, finish, development_responsible, status)

        return "Bot has been triggered."
    else:
        return "Invalid request method."


if __name__ == '__main__':
    app.run(debug=True)
