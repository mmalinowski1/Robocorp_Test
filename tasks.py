from robocorp.tasks import task
from robocorp import browser

from RPA.HTTP import HTTP
from RPA.Excel.Files import Files
from RPA.PDF import PDF
from RPA.Tables import Tables
from RPA.Archive import Archive

@task
def order_robots_from_RobotSpareBin():
    """
    Orders robots from RobotSpareBin Industries Inc.
    Saves the order HTML receipt as a PDF file.
    Saves the screenshot of the ordered robot.
    Embeds the screenshot of the robot to the PDF receipt.
    Creates ZIP archive of the receipts and the images.
    """
    browser.configure(
        slowmo=100,
    )
    open_robot_order_website()
    orders = get_orders()
    for row in orders:
        close_annoying_modal()
        fill_the_form(row)
    archive_receipts()
    
    
def open_robot_order_website():
    """opens the website for automation
    """
    browser.goto("https://robotsparebinindustries.com/#/robot-order")
    
def get_orders():
    """Downloads excel file from the given URL"""
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True)
    table = Tables()
    data = table.read_table_from_csv("C:/Users/MMalinowski/Robocorp/my-rsb-robot2/orders.csv")
    return data

def close_annoying_modal():
    """Clicks and closes the pop-up"""
    page = browser.page()
    page.click("text=OK")
    
def fill_the_form(rows):
    """Fills in the form with the data from CSV file"""
    page = browser.page()
    page.select_option("#head", rows['Head'])
    selector_template = 'input[type="radio"][value="{}"]'
    page.check(selector_template.format(rows['Body']))
    page.get_by_placeholder("Enter the part number for the legs").fill(str(rows["Legs"]))
    page.fill("#address", rows["Address"])
    page.click("#preview")
    page.click("#order")
    if(page.query_selector('.alert-danger')):
        page.click("#order")
    store_receipt_as_pdf(rows["Order number"])
    screenshot_robot(rows["Order number"])
    combine_screenshot_with_pdf("output/Archive/{}.png".format(rows["Order number"]),"output/Archive/{}.pdf".format(rows["Order number"]),"output/Archive/{}.pdf".format(rows["Order number"]))
    page.click("#order-another")
    
def store_receipt_as_pdf(order_number):
    """Export the data to a pdf file"""
    page = browser.page()
    sales_results_html = page.locator("#receipt").inner_html()

    pdf = PDF()
    pdf.html_to_pdf(sales_results_html, "output/Archive/{}.pdf".format(order_number))

def screenshot_robot(order_number):
    page = browser.page()
    page.screenshot(path="output/Archive/{}.png".format(order_number))   
    
def combine_screenshot_with_pdf(screenshot_path, pdf_path, output_pdf_path):
    pdf = PDF()
    # Add the screenshot to the existing PDF
    pdf.add_files_to_pdf(files=[screenshot_path, pdf_path], target_document=output_pdf_path)
    
def archive_receipts():
    archive = Archive()
    # Archive the folder with PDFs into a zip file
    archive.archive_folder_with_zip('output/Archive/', 'grouped_pdfs.zip', recursive=True)

