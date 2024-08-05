from robocorp.tasks import task
from robocorp import browser

from RPA.HTTP import HTTP
from RPA.Excel.Files import Files
from RPA.Tables import Tables
from RPA.PDF import PDF
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
        slowmo=500,
    )
    open_robot_order_website()
    close_annoying_modal()
    orders = get_orders()
    for row in orders:
        fill_the_form(row)
    
    archive_receipts()
        
    
    

def open_robot_order_website():
    """Navigates and open the order website."""
    browser.goto("https://robotsparebinindustries.com/#/robot-order")


def close_annoying_modal():
    """To close the pop."""
    page = browser.page()
    page.click("button:text('OK')")


def get_orders():
    """
    Download the orders file, read it as a table
    and return the result.
    """
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True)

    csv_file = Tables()
    csv_orders = csv_file.read_table_from_csv("orders.csv", header=True)

    return csv_orders


def fill_the_form(order_row):
    """Fills in the orders form."""
    page = browser.page()
    page.select_option("#head", value=order_row["Head"])
    page.check("#id-body-" + order_row["Body"])
    page.fill(".form-control", order_row["Legs"])
    page.fill("#address", order_row["Address"])
    page.click("#preview")
    page.click("#order")

    while not page.query_selector("#order-another"):
        page.click("#order")

    store_receipt_as_pdf(order_row["Order number"])  

    page.click("#order-another")
    close_annoying_modal()


def store_receipt_as_pdf(order_number):
    """Store the receipt of each order to a pdf file."""
    page = browser.page()
    receipt_html = page.locator("#receipt").inner_html()

    pdf = PDF()
    pdf_file = f"output/receipts/receipt-{order_number}.pdf"
    pdf.html_to_pdf(receipt_html, pdf_file)

    # screenshot
    screenshot_robot(order_number)
    screenshot_file = f"output/screenshots/screenshot-{order_number}.png"

    # embed screenshot to pdf file
    embed_screenshot_to_receipt(screenshot_file, pdf_file)


def screenshot_robot(order_number):
    """Takes a screenshot of each preview of robot."""
    page = browser.page()
    screenshot_path = f"output/screenshots/screenshot-{order_number}.png"
    page.locator("#robot-preview-image").screenshot(path=screenshot_path)
   


def embed_screenshot_to_receipt(screenshot, pdf_file):
    """Embeds the screenshot to the receipt pdf file."""    
    pdf = PDF()
    # pdf.add_files_to_pdf(  # this one adds a new page to a pdf file
    #     files=[screenshot],
    #     target_document=pdf_file,
    #     append=True
    # )

    pdf.open_pdf(pdf_file)

    # embed the pdf + image from the folders receipts and screenshots
    pdf.add_watermark_image_to_pdf(
        image_path = screenshot,
        # pdf_path = pdf_file,
        output_path = pdf_file
        # coverage = 0.5 (defaults to 0.2)
    )


def archive_receipts():
    """Archive all the receipts in a zip file."""
    file_path = "output/receipts"
    zip_lib = Archive()
    zip_lib.archive_folder_with_zip(file_path, "output/receipts.zip", include="*.pdf")
    
        
   
    