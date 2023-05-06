"""Insert the sales data for the week and export it as a PDF"""

import os
import csv
import time

from RPA.Browser.Selenium import Selenium
from RPA.HTTP import HTTP
from RPA.Excel.Files import Files, XlsxWorkbook, XlsWorkbook
from RPA.PDF import PDF
from RPA.Archive import Archive

from SeleniumLibrary.errors import ElementNotFound
from selenium.common.exceptions import ElementClickInterceptedException

from selenium.webdriver.remote.webelement import WebElement

def try_until_find_it(browser:Selenium, locator) -> WebElement:
    '''Insists on finding the element until it shows its face'''
    try:
        element = browser.find_element(locator)
        return element
    except ElementNotFound:
        return try_until_find_it(browser,locator)

# ***** First Certification! *****
def insert_the_sales_data_for_the_week_and_export_it_as_pdf():
    # Open intranet website
    browser = Selenium()
    browser.auto_close = False
    browser.open_available_browser("https://robotsparebinindustries.com/")

    # Log in
    username: WebElement = browser.find_element("css:#username")
    username.send_keys("maria")

    password: WebElement = browser.find_element("css:#password")
    password.send_keys("thoushallnotpass")

    browser.submit_form()

    # Download excel file
    http = HTTP()
    http.download("https://robotsparebinindustries.com/SalesData.xlsx", overwrite=True)

    # Fill and submit form using the files data
    excel = Files()
    workbook: XlsxWorkbook = excel.open_workbook("SalesData.xlsx")
    records = workbook.read_worksheet(header=True)
    
    for record in records:
        if all(record.values()):
            browser.wait_until_page_contains_element("css:#sales-form")

            first_name: WebElement = browser.find_element("css:#firstname")
            first_name.send_keys(record["First Name"])

            last_name: WebElement = browser.find_element("css:#lastname")
            last_name.send_keys(record["Last Name"])

            sales_result: WebElement = browser.find_element("css:#salesresult")
            sales_result.send_keys(record["Sales"])

            browser.select_from_list_by_value("css:#salestarget", str(record["Sales Target"]))

            browser.click_button('css:button[type="submit"]')
        else:
            continue

    # Collect the results
    browser.screenshot(
        'css:div.sales-summary', os.path.join(".", "output", "sales_summary.png")
    )

    # Export table as pdf
    table_locator = "css:div#sales-results"
    browser.wait_until_element_is_visible(table_locator)
    table: WebElement = browser.find_element(table_locator)
    table_html = table.get_attribute("outerHTML")
    
    pdf = PDF()
    pdf.html_to_pdf( # Why isn't it recognizing this method?
        table_html, os.path.join(".", "output", "table.pdf")
    ) 

    # Log out and close browser (Teardown)
    browser.click_button('css:button#logout')
    browser.close_browser()

    print("Done.")

# ***** Second Certification! *****
def build_robots_based_on_csv_file():
    # Download csv file
    http = HTTP()
    order_path = os.path.join(".","output","orders.csv")
    http.download(
        "https://robotsparebinindustries.com/orders.csv", order_path
    )



    # Open site
    browser = Selenium()
    browser.auto_close = False
    browser.open_available_browser("https://robotsparebinindustries.com/#/robot-order")

    # Get orders from file
    with open(order_path) as csv_file:
        orders = csv.reader(csv_file)
    # Iterate over orders
        for index, order in enumerate(orders):
            # Ignore headers
            if index == 0:
                continue
            else:
                # Close Box
                browser.click_button('css:div[class="alert-buttons"] > button')

                # Create Robot
                # Remember order number
                order_number = order[0]
                
                # Select Head
                head = order[1]
                browser.select_from_list_by_index("css:select#head", head)
                
                # Select Body
                ### Couldn't understand how the "select_radio_button" method works
                ### So i'm going with the old reliable: xpath patterns
                body = order[2]
                browser.click_button(f'xpath://*[@id="id-body-{body}"]') 
                
                # Select Legs
                leg = order[3]
                leg_field: WebElement = browser.find_element(
                    'css:input[placeholder="Enter the part number for the legs"]'
                )
                leg_field.send_keys(leg)
                
                # Input Address
                address = order[4]
                address_field: WebElement = browser.find_element(
                    'css:input[placeholder="Shipping address"]'
                )
                address_field.send_keys(address)

                # Screenshot Preview 
                
                image_name = f"order-{order_number}-robot.png"
                image_path = os.path.join(".", "output", image_name)
                browser.click_button('css:button#preview')
                browser.wait_until_element_is_visible("css:div#robot-preview-image > img + img + img")
                browser.screenshot('css:div#robot-preview-image', image_path)

                # Order
                browser.click_button('css:button#order')

                # Get Receipt HTML
                receipt_name = f"order-{order_number}-receipt.pdf"
                receipt_locator = "css:div#receipt"
                receipt_path = os.path.join(".","output", "receipts", receipt_name)
                try:
                    receipt: WebElement = browser.find_element(receipt_locator)
                # If it does not find receipt, it means we have a server error
                # In this case, we will keep clicking the order button until we get it
                except ElementNotFound:
                    while not browser.does_page_contain(receipt_locator):
                        # It can raise Element not found if it loads while checking
                        try:
                            browser.click_button('css:button#order')
                        except ElementNotFound:
                            break
                    receipt: WebElement = browser.find_element(receipt_locator)

                receipt_html = receipt.get_attribute("outerHTML")

                # Create Receipt PDF
                pdf = PDF()
                pdf.html_to_pdf(receipt_html, receipt_path)

                # Put Preview in receipt
                pdf.add_files_to_pdf([image_path,], receipt_path, append=True)
                
                # Alert
                print(f" *** Order of number {order_number} processed! ***")

                # Click "order another" and reload page
                browser.click_button("css:button#order-another")
                
                ### Keeps raising ElementClickInterceptedException, but why??
                browser.wait_until_page_contains_element('css:div[class="alert-buttons"] > button')

        # Archive Receipts            
        zip = Archive()
        zip.archive_folder_with_zip(
            os.path.join(".","output","receipts"), "receipts.zip"
        )      

    # Teardown
    browser.close_browser()
    
    print("Done")


if __name__ == "__main__":
    build_robots_based_on_csv_file()
