import argparse
import json
# external modules
from selenium import webdriver
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
import sys
import time
import xlwt
import xlrd


def parse_config_file(path):
    try:
        with open(path) as data_file:
            data = json.load(data_file)
    except:
        print "Error while parsing config json file."
        sys.exit(1)

    return data


def validate_vat_numbers(path, column, has_title):
    """ Function checks if file suposed to conatin comapnies vat number indeed has those numbers in correct format
    :return: true if successfull, false otherwise
    """
    vat_numbers = list()
    print "Opening file with VAT numbers."
    try:
        vat_numbers_book = xlrd.open_workbook(path)
        vat_numbers_sheet = vat_numbers_book.sheet_by_index(0)
    except:
        print "Couldn't open specified file with VAT numbers or couldn't open sheet 0."
        return vat_numbers
    print "Opened successfully. Checking VAT numbers . . ."
    i = 0
    if has_title:
        i += 1 # skip cell with column title
    while vat_numbers_sheet.cell_value(i, column) != '':
        if len(str((int(float(vat_numbers_sheet.cell_value(i, column)))))) != 10:
            print "Cell number " + str(i) + "contains vat number with incorrect format: " + str(vat_numbers_sheet.cell_value(i, column))
            i += 1
            continue
        vat_numbers.append((int(float(vat_numbers_sheet.cell_value(i, column)))))
        i += 1
    print "Done."
    return vat_numbers


def init_web_browser(path, url, use_display):
    """ Function initializes web driver and goes to passed url

    :rtype : WebDriver
    :param url: address to which browser should go after initialization
    :return: initialized driver handle
    """
    print "Initializing web browser engine ..."
    dcap = dict(DesiredCapabilities.PHANTOMJS)
    dcap["phantomjs.page.settings.userAgent"] = ("Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/53 (KHTML, like Gecko) "
                                                 "Chrome/15.0.87")
    browser = webdriver.PhantomJS(executable_path=path, desired_capabilities=dcap)
    browser.get(url)
    print "Done."
    return browser


def get_web_page_element(driver, label, method):
    """ Look for an element on web page and wait until it's available

    :param driver: WebDriver handle
    :param label: name, id or xpath to searched element
    :param method: method of looking up label (by id, name or xpath)
    :return: handle to obtained element if successful, 0 otherwise
    """
    unit = 0.001  # milliseconds
    inc = 100 * unit
    limit = 10 * inc  # waiting limit
    c = 0

    while c < limit:
        try:
            element = None
            if method == "id":
                element = driver.find_element_by_id(label)
            elif method == "xpath":
                element = driver.find_element_by_xpath(label)
            elif method == "name":
                element = driver.find_element_by_name(label)
            return element  # Success
        except:
            time.sleep(inc)
            c += inc
    return 0


def create_results_sheet(path, columns_titles):
    """ Function creates excel sheet for results

    :type columns_titles: list
    :param columns_titles:  titles of columns
    :return: created book and sheet handles
    """
    print "Creating file with results ..."
    book = xlwt.Workbook(encoding="utf-8")
    sheet = book.add_sheet("Python Sheet 1")
    for index, value in enumerate(columns_titles):
        sheet.write(0, index, value)
    book.save(path)
    print "Done."
    return [book, sheet]


def serialize_to_sheet(sheet, results, row):
    for index, value in enumerate(results):
        sheet.write(row, index, value)


def get_company_data(driver, vat):
    """
    :param driver: WebDriver driver used for scrapping
    :param vat: vat number of company whose data is going to be scrapped
    :return: dict with company financial data
    """

    # enter company's vat and search
    try:
        vat_input = get_web_page_element(driver, '//*[@id="search_box_query"]', 'xpath')
        vat_input.clear()
        vat_input.send_keys(vat)
        search_button = get_web_page_element(driver, '//*[@id="search_box"]/form/input[2]', 'xpath')
        search_button.click()
    except:
        print "Error when searching for " + vat + ", possibly due to script implementation or browser related."
        return ["ERR", "ERR", "ERR", "ERR", "ERR"]

    # go to page with detailed description
    try:
        detailed_description = get_web_page_element(driver, '/html/body/div[3]/div/div/table/tbody/tr[2]/td[1]/a', 'xpath')
        detailed_description.click()
    except:
        print "Company with VAT number " + str(vat) + " not present in databse."
        return [str(vat), "NOT_FOUND", "NOT_FOUND", "NOT_FOUND", "NOT_FOUND"]

    # obtain company's data
    try:
        vat_number = get_web_page_element(driver, '//*[@id="tab_detail"]/tbody/tr[1]/td', 'xpath').text.split('\n', 1)[0]
    except:
        print "Couldn't get " + str(vat) + " vat number - something is wrong!"
        vat_number = str(vat)

    try:
        name = get_web_page_element(driver, '//*[@id="debt_card_header"]/div/h1', 'xpath').text.split('\n', 1)[0]
    except:
        print "Couldn't get " + str(vat) + " company's name"
        name = "NOT_FOUND"

    try:
        city = get_web_page_element(driver, '//*[@id="tab_detail"]/tbody/tr[2]/td', 'xpath').text.split('\n', 1)[0]
    except:
        print "Couldn't get " + str(vat) + " city"
        city = "NOT_FOUND"

    try:
        category = get_web_page_element(driver, '//*[@id="tab_detail"]/tbody/tr[4]/td', 'xpath').text.split('\n', 1)[0]
    except:
        print "Couldn't get " + str(vat) + "category"
        category = "NOT_FOUND"

    try:
        debt_value = get_web_page_element(driver, '//*[@id="tab_detail"]/tbody/tr[3]/td', 'xpath').text.split('\n', 1)[0]
    except:
        print "Couldn't get " + str(vat) + "debt_value"
        debt_value = "NOT_FOUND"

    # print data (debug only)
    print ">> vat_number: " + vat_number
    print ">> name: " + name
    print ">> city: " + city
    print ">> category: " + category
    print ">> debt_value: " + debt_value

    return [vat_number, name, city, category, debt_value]


def main():
    parser = argparse.ArgumentParser(description='Webscrapper bot arguments parser.')
    parser.add_argument('Config_file_path')
    args = parser.parse_args()
    parsed_args = vars(args)
    config_file_path = parsed_args['Config_file_path']
    config = parse_config_file(config_file_path)

    phantomjs_path = config['PHANTOMJS_PATH']
    vat_source_path = config['VAT_SOURCE']
    output_file_path = config['OUTPUT_FILE']
    vat_column_index = int(config['VAT_COLUMN_INDEX'])
    vat_column_contains_title = config['VAT_COLUMN_CONTAINS_TITLE'] == 'True'

    vat_numbers = validate_vat_numbers(vat_source_path, vat_column_index, vat_column_contains_title)

    if len(vat_numbers) == 0:
        print "Vat file contains errors. Terminating script."
        sys.exit(1)

    # some hardcoded constant data
    url = "https://www.dlugi.info" # address of web page that's gonna be scrapped
    titles = ["VAT_Number", "Name", "City", "Category", "Debt_value"] # titles of columns in result sheet

    browser = init_web_browser(phantomjs_path, url, False)
    [book, sheet] = create_results_sheet(output_file_path, titles)

    # scrap
    for i, val in enumerate(vat_numbers):
        company_data = get_company_data(browser, val)
        serialize_to_sheet(sheet, company_data, i+1)
        book.save(output_file_path)

if __name__ == '__main__':
    main()




