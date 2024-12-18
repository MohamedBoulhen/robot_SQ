from robocorp.tasks import task
from robocorp import browser
from RPA.HTTP import HTTP
from RPA.Excel.Files import Files
from RPA.PDF import PDF
import os
import logging

# Configure logging
logging.basicConfig(level=logging.INFO)

# Load credentials securely from environment variables or a config file
USERNAME = os.getenv("BOT_USERNAME", "maria")
PASSWORD = os.getenv("BOT_PASSWORD", "thoushallnotpass")
EXCEL_URL = "https://robotsparebinindustries.com/SalesData.xlsx"
BASE_URL = "https://robotsparebinindustries.com/"

@task
def robot_spare_bin_python():
    """Insert the sales data for the week and export it as a PDF."""
    browser.configure(slowmo=100)
    logging.info("Bot started.")
    
    try:
        open_the_intranet_website()
        log_in()
        download_excel_file()
        fill_form_with_excel_data()
        collect_results()
        export_as_pdf()
    except Exception as e:
        logging.error(f"An unexpected error occurred: {e}")
    finally:
        log_out()
        logging.info("Bot execution finished.")


def open_the_intranet_website():
    """Navigates to the given URL."""
    try:
        browser.goto(BASE_URL)
        logging.info("Successfully opened the intranet website.")
    except Exception as e:
        logging.error(f"Failed to open the intranet website: {e}")
        raise


def log_in():
    """Logs in using secure credentials."""
    try:
        page = browser.page()
        page.fill("#username", USERNAME)
        page.fill("#password", PASSWORD)
        page.click("button:text('Log in')")
        logging.info("Successfully logged in.")
    except Exception as e:
        logging.error(f"Failed to log in: {e}")
        raise


def download_excel_file():
    """Downloads the Excel file from the given URL."""
    try:
        http = HTTP()
        http.download(url=EXCEL_URL, overwrite=True)
        logging.info("Excel file downloaded successfully.")
    except Exception as e:
        logging.error(f"Failed to download the Excel file: {e}")
        raise


def fill_form_with_excel_data():
    """Processes data from Excel and submits the form."""
    try:
        excel = Files()
        excel.open_workbook("SalesData.xlsx")
        worksheet = excel.read_worksheet_as_table("data", header=True)
        excel.close_workbook()

        logging.info(f"Processing {len(worksheet)} records from Excel.")
        for row in worksheet:
            try:
                fill_and_submit_sales_form(row)
            except Exception as e:
                logging.warning(f"Failed to process record {row}: {e}")
    except Exception as e:
        logging.error(f"Failed to read or process the Excel file: {e}")
        raise


def fill_and_submit_sales_form(sales_rep):
    """Submits a single sales record through the form."""
    try:
        page = browser.page()
        page.fill("#firstname", sales_rep["First Name"])
        page.fill("#lastname", sales_rep["Last Name"])
        page.select_option("#salestarget", str(sales_rep["Sales Target"]))
        page.fill("#salesresult", str(sales_rep["Sales"]))
        page.click("text=Submit")
        logging.info(f"Submitted form for {sales_rep['First Name']} {sales_rep['Last Name']}.")
    except Exception as e:
        logging.error(f"Failed to submit sales form for {sales_rep}: {e}")
        raise


def collect_results():
    """Takes a screenshot of the results page."""
    try:
        page = browser.page()
        page.screenshot(path="output/sales_summary.png")
        logging.info("Sales summary screenshot saved.")
    except Exception as e:
        logging.error(f"Failed to take a screenshot of sales summary: {e}")
        raise


def export_as_pdf():
    """Exports sales results as a PDF."""
    try:
        page = browser.page()
        sales_results_html = page.locator("#sales-results").inner_html()

        pdf = PDF()
        pdf.html_to_pdf(sales_results_html, "output/sales_results.pdf")
        logging.info("Sales results exported as PDF.")
    except Exception as e:
        logging.error(f"Failed to export sales results as a PDF: {e}")
        raise


def log_out():
    """Logs out from the application."""
    try:
        page = browser.page()
        page.click("text=Log out")
        logging.info("Successfully logged out.")
    except Exception as e:
        logging.error(f"Failed to log out: {e}")
