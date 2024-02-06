from robocorp.tasks import task
from robocorp import browser
from RPA.HTTP import HTTP
from RPA.Excel.Files import Files
from RPA.PDF import PDF

browser.configure(
	# browser_engine="chrome",
	slowmo=500,
)
page = browser.page()
http = HTTP()
excel = Files()
pdf = PDF()




@task
def robot_spare_bin_python():
    """ inserts the sales data for the week and export it as pdf """
    open_the_intranet_website()
    log_in()
    download_excel_file()
    #fill_and_submit_sales_form()
    fill_form_with_excel_data()
    collect_results()
    export_as_pdf()
    log_out()


def open_the_intranet_website():
    """Navigates to the given URL"""
    browser.goto("https://robotsparebinindustries.com/")


def log_in():
    """Logs in to the intranet"""
    page.fill("#username", "maria")
    page.fill("#password", "thoushallnotpass")
    page.click("button:text('Log in')")


def fill_and_submit_sales_form(sales_rep):
    """Fills in the sales data and click the 'Submit' button"""
    page = browser.page()

    page.fill("#firstname", sales_rep["First Name"])
    page.fill("#lastname", sales_rep["Last Name"])
    page.select_option("#salestarget", str(sales_rep["Sales Target"]))
    page.fill("#salesresult", str(sales_rep["Sales"]))
    page.click("text=Submit")


def download_excel_file():
    """Downloads the sales data as an Excel file"""
    http.download(url="https://robotsparebinindustries.com/SalesData.xlsx", overwrite=True)


def fill_form_with_excel_data():
    excel.open_workbook("SalesData.xlsx")
    worksheet = excel.read_worksheet_as_table("data", header=True)
    excel.close_workbook()
    for row in worksheet:
        fill_and_submit_sales_form(row)


def collect_results():
    """Take a screenshot of the page"""
    page.screenshot(path="output/sales_summary.png")




def export_as_pdf():
    sales_result_html = page.locator("#sales-results").inner_html()
    pdf.html_to_pdf(sales_result_html, "output/sales_results.pdf")


def log_out():
    """Logs out from the intranet"""
    page.click("button:text('Log out')")



    