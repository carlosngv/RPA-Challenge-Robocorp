from robocorp import browser
from robocorp.tasks import task

from RPA.HTTP import HTTP
from RPA.Excel.Files import Files

http = HTTP()
lib = Files()

EXCEL_DATA_URL = 'https://rpachallenge.com/assets/downloadFiles/challenge.xlsx'
CHALLENGE_URL = 'https://rpachallenge.com/'
DATA_PATH = './output/data.xlsx'

@task
def input_forms_challenge():
    browser.configure(
        slowmo=100,
    )

    download_data()
    start_challenge()
    process_data()

def start_challenge():
    browser.goto(CHALLENGE_URL)
    page = browser.page()
    page.click('//button[contains(text(),"Start")]')


def fill_form(data):
    page = browser.page()
    page.fill('//input[@ng-reflect-name="labelPhone"]', str(data['Phone Number']))
    page.fill('//input[@ng-reflect-name="labelEmail"]', data['Email'])
    page.fill('//input[@ng-reflect-name="labelCompanyName"]', data['Company Name'])
    page.fill('//input[@ng-reflect-name="labelAddress"]', data['Address'])
    page.fill('//input[@ng-reflect-name="labelLastName"]', data['Last Name'])
    page.fill('//input[@ng-reflect-name="labelFirstName"]', data['First Name'])
    page.fill('//input[@ng-reflect-name="labelRole"]', data['Role in Company'])
    page.click('//input[@value="Submit"]')

def download_data():
    http.download(EXCEL_DATA_URL, target_file=DATA_PATH)

def process_data():
    lib.open_workbook(DATA_PATH)
    try:
        table = lib.read_worksheet_as_table(header=True)
    except:
        raise RuntimeError('Contact your admnistrator')
    finally:
        lib.close_workbook()

    for data in table:
        fill_form(data)
