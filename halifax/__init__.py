from .etf import halifax_etf
from .investment import halifax_investment
from .mutual_fund import halifax_mutual_funds
from utils import clean_spreadsheet, get_xlsx_filepath, setup_driver, delay


def halifax_runner():
    out = get_xlsx_filepath("halifax.xlsx")
    clean_spreadsheet(out)

    driver = setup_driver(True)
    halifax_etf(driver, out)
    driver.quit()

    delay(10, 20)

    driver = setup_driver(True)
    halifax_investment(driver, out)
    driver.quit()

    delay(10, 20)

    driver = setup_driver(True)
    halifax_mutual_funds(driver, out)
    driver.quit()
