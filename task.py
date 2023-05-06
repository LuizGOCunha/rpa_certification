"""Insert the sales data for the week and export it as a PDF"""

import os

from RPA.Browser.Selenium import Selenium
from RPA.HTTP import HTTP
from RPA.Excel.Files import Files, XlsxWorkbook, XlsWorkbook
from RPA.PDF import PDF
from selenium.webdriver.remote.webelement import WebElement


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


if __name__ == "__main__":
    insert_the_sales_data_for_the_week_and_export_it_as_pdf()
