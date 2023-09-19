from robocorp.tasks import task
from robocorp import browser, http
from RPA.Tables import Tables
from RPA.PDF import PDF
import zipfile

@task
def order_robots_from_RobotSpareBin():
    """
    Orders robots from RobotSpareBin Industries Inc.
    Saves the order HTML receipt as a PDF file.
    Saves the screenshot of the ordered robot.
    Embeds the screenshot of the robot to the PDF receipt.
    Creates ZIP archive of the receipts and the images.
    """    
    download_and_prep_spreadsheet()
    open_robot_order_website()
    fill_form_with_excel_date()

def download_and_prep_spreadsheet():
    """Downloads Excel File"""
    http.download(url = "https://robotsparebinindustries.com/orders.csv", overwrite = True)
    tables = Tables()
    table = tables.read_table_from_csv("orders.csv")
    return table

def open_robot_order_website():
    browser.goto("https://robotsparebinindustries.com/")
    page = browser.page()
    page.click("#root > header > div > ul > li:nth-child(2) > a")
    page.click("button:text('OK')")

def fill_form_with_excel_date():
    """Loop through CSV and fill in an entry for each row representing a robot order"""
    table = download_and_prep_spreadsheet()
    pdfFiles = [] #initialize the list
    page = browser.page()
    for row in table:
        orderNumber = row['Order number']
        head = row['Head']
        body = row['Body']
        legs = row['Legs']
        address = row['Address']

        """fill head"""
        print('Head: ' + head)
        page.select_option('css=#head', f"{head}")

        """fill body"""
        print('Body: ' + body)
        page.click(f'#id-body-{body}')

        """fill legs"""
        print('Legs: ' + legs)
        page.fill("css=input[placeholder='Enter the part number for the legs']", legs)

        """fill address"""
        print('Address: ' + address)
        page.fill('#address', address)

        page.click('#order')
        
        while page.is_visible(".alert.alert-danger") == True:
            page.click('#order')
            if not page.is_visible(".alert.alert-danger"):
                break
        
        """create the pdf receipt"""
        pdf_file = store_receipt_as_pdf(orderNumber, pdfFiles)
        screenshot = screenshot_robot(orderNumber)
        embed_screenshot_to_receipt(screenshot, pdf_file)

        page.click('#order-another')
        page.click("button:text('OK')")

    outputZip = "output/all_receipts.zip"
    archive_receipts(pdfFiles, outputZip)

def store_receipt_as_pdf(orderNumber, pdfFiles):
    """Saves the screenshot of the ordered robot."""
    page = browser.page()
    receipt_html = page.locator("#order-completion").inner_html()

    pdf = PDF()
    pdf_file = f"output/{orderNumber}.pdf"
    pdf.html_to_pdf(receipt_html, pdf_file)
    pdfFiles.append(pdf_file)  # Append to the list
    return pdf_file

def screenshot_robot(orderNumber):
    """Saves the screenshot of the ordered robot."""
    page = browser.page()
    screenshot = f"output/screenshots/{orderNumber}.png"
    page.screenshot(path = screenshot)
    return screenshot

def embed_screenshot_to_receipt(screenshot, pdf_file):
    """Embeds the screenshot of the robot to the PDF receipt."""
    pdf = PDF()
    list_of_files = [
        pdf_file,
        screenshot
    ]
    pdf.add_files_to_pdf(files=list_of_files, target_document = pdf_file)

def archive_receipts(pdfFiles, outputZip):
    """Creates ZIP archive of the receipts and the images."""
    with zipfile.ZipFile(outputZip, 'w') as zipf:
        for file in pdfFiles:
            zipf.write(file)