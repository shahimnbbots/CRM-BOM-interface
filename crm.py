import traceback
from tkinter import messagebox
from selenium import webdriver
from selenium.webdriver import Keys, ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from subprocess import CREATE_NO_WINDOW
import time
import pymysql
import tkinter as tk


def database():
    # Establish a connection to the database
    conn = pymysql.connect(
        host='172.20.50.169',
        user='internal_apps',
        password='Schemax@2023',
        database='internal_apps',
        charset='utf8mb4',
        cursorclass=pymysql.cursors.DictCursor
    )
    # Create a cursor object
    cursor = conn.cursor()
    # Execute the SQL query to fetch all data
    cursor.execute("SELECT * FROM crm_fetched_data WHERE Status = 'Pending'")
    # Fetch all rows of data
    fetched_data = cursor.fetchall()
    # Close the cursor and connection
    cursor.close()
    conn.close()
    # Return the fetched data
    return fetched_data
# Call the function to fetch data from the database
all_data = database()
# Print the fetched data
for row in all_data:
    print(row)


def update_status(id, new_status):
    # Establish a connection to the database
    conn = pymysql.connect(
        host='172.20.50.169',
        user='internal_apps',
        password='Schemax@2023',
        database='internal_apps',
        charset='utf8mb4',
        cursorclass=pymysql.cursors.DictCursor
    )
    # Create a cursor object
    cursor = conn.cursor()
    # Update the Status for the given ID
    update_query = "UPDATE crm_fetched_data SET Status = %s WHERE id = %s"
    cursor.execute(update_query, (new_status, id))
    # Commit the changes
    conn.commit()
    # Close the cursor and connection
    cursor.close()
    conn.close()


def item_field(id, item_no, group, product_group, item_group, description, operation, unit_of_measure, alternate_unit_of_measure, factor, currency, budget_price,
               consumption, waste, placement, material_code, fabric_type, structure, quality, finish, development_responsible, status):
    try:
        # Your existing code for browser automation here...
        service = Service()
        options = Options()
        options.add_experimental_option('detach', True)
        service.creation_flags = CREATE_NO_WINDOW
        driver = webdriver.Chrome(service=service, options=options)
        wait = WebDriverWait(driver, 30)
        action = webdriver.ActionChains(driver)
        driver.get("http://intranetn.shahi.co.in:8080/ShahiExportIntranet/login")
        username = wait.until(ec.presence_of_element_located((By.ID, 'username')))
        username.send_keys("755921")
        password = driver.find_element(By.ID, "password")
        password.send_keys("vis123")
        driver.find_element(By.ID, 'savebutton').click()
        time.sleep(2)
        handles = [driver.current_window_handle]
        try:
            driver.execute_script(
                "javascript:openMenuPage('null' , 'CRM - Work Place (New M3)' , '2447' , 'F' , 'Applications'  );")
        except:
            driver.execute_script(
                "javascript:openMenuPage('null' , 'CRM - Work Place (New M3)' , '2447' , 'F' , 'Applications'  );")
        for i in driver.window_handles:
            if i not in handles:
                driver.switch_to.window(i)
                break
        handles.append(driver.current_window_handle)
        time.sleep(2)
        try:
            driver.execute_script(
                f"javascript:openAccessPage('http://crmm4.shahi.co.in:8080/CRMPRDN/CRMPRDNEW.jsp' , 'CRM' , '2448' , 'R' , '755921' , 'N', '50020096', 'N');")
        except:
            driver.execute_script(
                f"javascript:openAccessPage('http://crmm4.shahi.co.in:8080/CRMPRDN/CRMPRDNEW.jsp' , 'CRM' , '2448' , 'R' , '755921' , 'N', '50020096', 'N');")
        for i in driver.window_handles:
            if i not in handles:
                driver.switch_to.window(i)
                break
        driver.execute_script("getHome('340')")
        frame = driver.find_element(By.ID, "mainFrame")
        driver.switch_to.frame(frame)
        apps = wait.until(
            ec.presence_of_element_located((By.XPATH, '//*[@id="mainContainer"]/div[1]/div[4]/div/span[1]')))
        apps.click()
        time.sleep(5)
        material = (driver.find_element(By.XPATH, '//*[@id="mainContainer"]/div[1]/div[5]/div/span[1]'))
        driver.execute_script(f"arguments[0].click();", material)
        # item_no
        item = (driver.find_element(By.XPATH, '//*[@id="styleid4"]'))
        driver.execute_script(f"arguments[0].value='{item_no}'", item)
        time.sleep(10)
        # Group
        if group == 'FAB':
            fabric = driver.find_element(By.XPATH, '//*[@id="addDiv"]')
            fabric.click()
            time.sleep(5)
            search_fab_code = driver.find_element(By.XPATH, '//*[@id="mtpageid"]/tbody/tr[2]/td[2]/table/tbody/tr/td[2]/a/img')
            search_fab_code.click()
            time.sleep(5)
            iframe = driver.find_element(By.ID, 'fablistpagefrm')
            driver.switch_to.frame(iframe)
            select_desc = driver.find_element(By.XPATH, '/html/body/form/table/tbody/tr/td/select')
            # Get all options
            options = select_desc.find_elements(By.TAG_NAME, "option")
            # Iterate over options and print their text
            # Iterate over options and print their text
            for option in options:
                print(option.text)
                if description in option.text:
                    print("Match found")
                    print("Matching description found. Scrolling into view...")
                    print("Performing double-click action...")
                    ActionChains(driver).move_to_element(option).double_click(option).perform()
                    print("Double-click action performed.")
                    break
            # After interacting with elements inside the iframe, switch back to the default content
            driver.switch_to.default_content()
            driver.switch_to.frame(frame)
            time.sleep(1)
            fabric_type_value = driver.find_element(By.ID, 'itemgrpID_TXT')
            fabric_type_value.send_keys(fabric_type)
            fabric_type_value.send_keys(Keys.ARROW_DOWN)
            fabric_type_value.send_keys(Keys.ENTER)
            structure_value = driver.find_element(By.ID, 'FRE3F_TXT')
            structure_value.send_keys(structure)
            structure_value.send_keys(Keys.ARROW_DOWN)
            structure_value.send_keys(Keys.ENTER)
            placement_val = driver.find_element(By.ID, 'itemtypeID_TXT')
            placement_val.send_keys(placement)
            placement_val.send_keys(Keys.ARROW_DOWN)
            placement_val.send_keys(Keys.ENTER)
            quality_val = driver.find_element(By.ID, 'productgrpID_TXT')
            quality_val.send_keys(quality)
            quality_val.send_keys(Keys.ARROW_DOWN)
            quality_val.send_keys(Keys.ENTER)
            fab_finish = driver.find_element(By.ID, 'FRE4_TXT')
            fab_finish.clear()
            fab_finish.send_keys(finish)
            fab_finish.send_keys(Keys.ARROW_DOWN)
            fab_finish.send_keys(Keys.ENTER)
            # Development Responsible
            if development_responsible:
                resp_person = driver.find_element(By.ID, 'appuserfab_TXT')
                resp_person.send_keys(development_responsible)
                resp_person.send_keys(Keys.ARROW_DOWN)
                resp_person.send_keys(Keys.ENTER)
            # UOM
            uom_value = driver.find_element(By.XPATH, '//*[@id="basicUOM_TXT"]')
            uom_value.clear()
            uom_value.send_keys(unit_of_measure)
            uom_value.send_keys(Keys.ARROW_DOWN)
            uom_value.send_keys(Keys.ENTER)
            time.sleep(2)
            # ALT_UOM
            if alternate_unit_of_measure:
                alt_uom = driver.find_element(By.XPATH, '//*[@id="alunID_TXT"]')
                alt_uom.send_keys(alternate_unit_of_measure)
                alt_uom.send_keys(Keys.ARROW_DOWN)
                alt_uom.send_keys(Keys.ENTER)
            # Factor
            if factor:
                factor_value = driver.find_element(By.XPATH, '//*[@id="conversion"]')
                factor_value.send_keys(factor)
            # Currency
            currency_value = driver.find_element(By.XPATH, '//*[@id="currID_TXT"]')
            currency_value.clear()
            currency_value.send_keys(currency)
            currency_value.send_keys(Keys.ARROW_DOWN)
            currency_value.send_keys(Keys.ENTER)
            # Budget_price
            price_value = driver.find_element(By.XPATH, '//*[@id="pricepur"]')
            price_value.clear()
            price_value.send_keys(budget_price)
            # Consumption
            consp = driver.find_element(By.XPATH, '//*[@id="consumption1"]')
            consp.clear()
            consp.send_keys(consumption)
            # Wastage
            wastage = driver.find_element(By.XPATH, '//*[@id="wastage1"]')
            wastage.clear()
            wastage.send_keys(waste)
            # Cost Group
            costgroup = driver.find_element(By.XPATH, '//*[@id="costgroup_TXT"]')
            costgroup.clear()
            costgroup.send_keys("NORMAL")
            # driver.execute_script("argum")
            time.sleep(1)
            costgroup.send_keys(Keys.ARROW_DOWN)
            costgroup.send_keys(Keys.ENTER)
            # If everything runs without errors, update the Status to 'Created'
            update_status(id, 'Created')
        if group == 'TRM':
            create_trim = driver.find_element(By.XPATH, '//*[@id="updDiv"]')
            create_trim.click()
            time.sleep(5)
            # product_group
            input_field = driver.find_element(By.XPATH, '//*[@id="productgrpID_TXT"]')
            input_field.clear()
            input_field.send_keys(product_group)
            input_field.send_keys(Keys.ARROW_DOWN)
            input_field.send_keys(Keys.ENTER)
            time.sleep(3)
            # material_code-trim-code
            material_code_value = driver.find_element(By.XPATH, '//*[@id="fabcode"]')
            material_code_value.send_keys(material_code)
            # Item_group
            group_drop = driver.find_element(By.XPATH, '//*[@id="itemgrpID"]')
            group_drop.send_keys(item_group[0])
            group_drop.send_keys(Keys.ENTER)
            # Description
            description_value = driver.find_element(By.XPATH, '//*[@id="fabdescnew"]')
            description_value.send_keys(description)
            # Operation
            operation_drop = driver.find_element(By.XPATH, '//*[@id="FRE3TR_TXT"]')
            operation_drop.send_keys(operation)
            operation_drop.send_keys(Keys.ARROW_DOWN)
            operation_drop.send_keys(Keys.ENTER)
            # Development Responsible
            if development_responsible:
                resp_person = driver.find_element(By.ID, 'appuserfab_TXT')
                resp_person.send_keys(development_responsible)
                resp_person.send_keys(Keys.ARROW_DOWN)
                resp_person.send_keys(Keys.ENTER)
            # UOM
            uom_value = driver.find_element(By.XPATH, '//*[@id="basicUOM_TXT"]')
            uom_value.clear()
            uom_value.send_keys(unit_of_measure)
            uom_value.send_keys(Keys.ARROW_DOWN)
            uom_value.send_keys(Keys.ENTER)
            time.sleep(2)
            # ALT_UOM
            if alternate_unit_of_measure:
                alt_uom = driver.find_element(By.XPATH, '//*[@id="alunID_TXT"]')
                alt_uom.send_keys(alternate_unit_of_measure)
                alt_uom.send_keys(Keys.ARROW_DOWN)
                alt_uom.send_keys(Keys.ENTER)
            # Factor
            if factor:
                factor_value = driver.find_element(By.XPATH, '//*[@id="conversion"]')
                factor_value.send_keys(factor)
            # Currency
            currency_value = driver.find_element(By.XPATH, '//*[@id="currID_TXT"]')
            currency_value.clear()
            currency_value.send_keys(currency)
            currency_value.send_keys(Keys.ARROW_DOWN)
            currency_value.send_keys(Keys.ENTER)
            # Budget_price
            price_value = driver.find_element(By.XPATH, '//*[@id="pricepur"]')
            price_value.clear()
            price_value.send_keys(budget_price)
            # Consumption
            consp = driver.find_element(By.XPATH, '//*[@id="consumption1"]')
            consp.clear()
            consp.send_keys(consumption)
            # Wastage
            wastage = driver.find_element(By.XPATH, '//*[@id="wastage1"]')
            wastage.clear()
            wastage.send_keys(waste)
            # Cost Group
            costgroup = driver.find_element(By.XPATH, '//*[@id="costgroup_TXT"]')
            costgroup.clear()
            costgroup.send_keys("NORMAL")
            # driver.execute_script("argum")
            time.sleep(1)
            costgroup.send_keys(Keys.ARROW_DOWN)
            costgroup.send_keys(Keys.ENTER)
            # Placement
            if placement:
                placement_value = driver.find_element(By.XPATH, '//*[@id="remarks"]')
                time.sleep(1)
                placement_value.send_keys(placement)
            # If everything runs without errors, update the Status to 'Created'
            update_status(id, 'Created')

    except Exception as e:
        # Handle any exceptions that might occur
        traceback.print_exc()
        messagebox.showerror("Error", str(e))
        # If there's an error, update the Status to 'Error'
        update_status(id, 'Error')


# def test_entry_bot():
# Call the database function to fetch the item_no
fetched_data = database()

# Iterate over fetched_data and extract item_no and group
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
    # # Skip running item_field if group is 'FAB' or empty
    # if group in ('FAB', ''):
    #     continue
    # if status != "Created":
    # Call the item_field function with the fetched item_no and group
    item_field(id, item_no, group, product_group, item_group, description, operation, unit_of_measure, alternate_unit_of_measure,
     factor, currency, budget_price, consumption, waste, placement, material_code, fabric_type, structure, quality, finish, development_responsible, status)

    time.sleep(5)  # Optionally, add a delay between function calls



# app = tk.Tk()
# app.geometry("300x100")
# app.title("CRM AUTOMATION.")
#
# # Create the button
# button = tk.Button(app, text="Run Bot", command=test_entry_bot, width=10)
# # Use grid to place the button in the center and span it across 4 columns
# button.grid(row=0, column=0, columnspan=4, sticky=tk.NSEW)
#
# # Configure grid weights to center the button
# app.grid_rowconfigure(0, weight=1)
# app.grid_columnconfigure(0, weight=1)
#
# # Start the Tkinter event loop
# app.mainloop()

