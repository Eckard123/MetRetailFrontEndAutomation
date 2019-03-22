from selenium import webdriver
import datetime
import calendar
import time
import xlrd
import xlwings
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.keys import Keys
import selenium.common.exceptions
import urllib3
from tryagain import retries
import pyautogui


# Setting details for the rest of operations to happen
excelfile_source= "C:\\Users\\EcBerry\\PycharmProjects\\webAutomation\\Reports\\ExcelScenarioInfoMetRetail\\Book1.xlsm"
from_row = 7
to_row = 8
i = from_row - 1
j = to_row - 1
k = from_row
scenarios_to_run = int(to_row - from_row)


# Defining some functions
def add_one_month(orig_date):
    new_year = orig_date.year
    new_month = orig_date.month + 1
    new_day = orig_date.day

    last_day_of_month = calendar.monthrange(new_year, new_month)[1]
    new_day = min(orig_date.day, last_day_of_month)
    Edited_date = orig_date.replace(day=1, month=new_month, year=new_year)
    return Edited_date.strftime("%d/%m/%Y")


def create_dob(date, excel_age):
    dob_year = date.year
    dob_month = date.month
    dob_day = date.day

    last_day_of_month = calendar.monthrange(dob_year, dob_month)[1]
    new_day = min(date.day, last_day_of_month)
    new_year = int(dob_year - excel_age)
    created_dob = date.replace(day=new_day, month=1, year=new_year)
    return created_dob.strftime("%d/%m/%Y")


def change_string(old_value):
    new_my_sum = old_value.replace(".0","")
    if new_my_sum.__len__() == 5:
        if new_my_sum[1] == "0":
           result = str(new_my_sum.replace("0000","0 000"))
        elif new_my_sum[1] != "0":
            result = str(new_my_sum.replace("000"," 000"))
    if new_my_sum.__len__() < 5:
        result = str(new_my_sum.replace("000"," 000"))
    elif new_my_sum.__len__() > 5:
        result = str(new_my_sum.replace("0000","0 000"))

    return result


Todays_date = datetime.date.today()
Start_date = add_one_month(Todays_date)
##########################################

for x in range(i, j):
    print("Value of i:")
    print(i)

    # Getting the workbook ready to write data to "Web_Output" sheet
    outputworkbook = xlwings.Book(excelfile_source)
    outputworksheet = outputworkbook.sheets[1]  # This gets the worksheet 'Web_Output'

    # Getting the ID number for the scenario
    workbook = xlrd.open_workbook(excelfile_source).sheet_by_name("Scenarios").cell_value(rowx=i, colx=0)
    Scenario_id = int(workbook)
    print(Scenario_id)
    print(Todays_date)
    print("Start date is: " + Start_date)

    # Locating and calling URL
    IE = webdriver.Chrome()
    IE.get("https://retail-dev.metropolitan.co.za/mmih-cdi-search-client/login")
    IE.maximize_window()

    # Setting username constant
    try:
        WebDriverWait(IE,30).until(EC.visibility_of_element_located((By.XPATH,'/html/body/div[2]/div[2]/div/div/div/form/div[1]/input')))
        username = IE.find_element_by_xpath("//input[@name = 'email']")
        username.send_keys("aeptest1@metropolitan.co.za")

        # Setting password constant
        password = IE.find_element_by_xpath("//input[@name = 'password']")
        password.send_keys("metro2875!!")

        # Click the login button
        Login = IE.find_element_by_xpath("//button[@type = 'submit']")
        Login.click()

    except(RuntimeError,urllib3.exceptions.MaxRetryError,selenium.common.exceptions.NoSuchElementException):
        print("Error occured with logon, please try again")
        IE.refresh()

    # Set the ID_number variable on web
    idelement = WebDriverWait(IE,30).until(EC.element_to_be_clickable((By.ID,'idnumber')))
    id_number = IE.find_element_by_xpath("//input[@id = 'idnumber']")
    id_number.send_keys(Scenario_id)

    # Click the search button
    Search = IE.find_element_by_xpath("//button[@type = 'submit']")
    Search.click()

    # In order to click 'Create new client'
    try:
        WebDriverWait(IE,60).until(EC.element_to_be_clickable((By.XPATH,'//*[@id="searchResultSection"]/div[2]/div/div[2]/button')))
        createnewclients = IE.find_elements_by_xpath('//*[@id="searchResultSection"]/div[2]/div/div[2]/button')
        for createnew in createnewclients:
            if createnew.get_attribute('class') == "btn btn-sm btn-mmih-primary":
                createnew.click()

    except(selenium.common.exceptions.NoSuchElementException):
            print("Session not responding, closing session.")
            IE.quit()


    try:
        WebDriverWait(IE,20).until(EC.url_contains("https://retail-dev.metropolitan.co.za/mmih-cdi-maintain-client/client-details/personal-details"))

        WebDriverWait(IE,20).until(EC.visibility_of_element_located((By.XPATH,'//*[@id="firstName"]')))
        firstname = IE.find_element_by_xpath('//*[@id="firstName"]')
        workbook = xlrd.open_workbook(excelfile_source).sheet_by_name("Scenarios").cell_value(rowx=i, colx=1)
        firstname.send_keys(workbook)

        surname = IE.find_element_by_xpath('//*[@id="lastName"]')
        surname_workbook = xlrd.open_workbook(excelfile_source).sheet_by_name("Scenarios").cell_value(rowx=i, colx=3)
        surname.send_keys(surname_workbook)

        DOB = IE.find_element_by_xpath('//*[@id="clientDOB"]')
        workbook = xlrd.open_workbook(excelfile_source).sheet_by_name("Scenarios").cell_value(rowx=i, colx=4)
        DOB.send_keys(workbook)

        Initials = IE.find_element_by_xpath('//*[@id="initials"]')
        workbook = xlrd.open_workbook(excelfile_source).sheet_by_name('Scenarios').cell_value(rowx=i,colx=2)
        Initials.send_keys(workbook)

        country = Select(IE.find_element_by_xpath('//*[@id="countryOfIdIssue"]'))
        country.select_by_index(index=1)
        time.sleep(1)

        gender = Select(IE.find_element_by_xpath('//*[@id="gender"]'))
        workbook_gender = xlrd.open_workbook(excelfile_source).sheet_by_name("Scenarios").cell_value(rowx=i, colx=5)
        time.sleep(1)

        if workbook_gender == "Male":
            gender.select_by_index(1)
        elif workbook_gender == "Female":
            gender.select_by_index(2)

        time.sleep(1)

        title = Select(IE.find_element_by_xpath('//*[@id="title"]'))
        if workbook_gender =="Male":
            title.select_by_index(1)
        elif workbook_gender == "Female":
            title.select_by_index(1)

    except(selenium.common.exceptions.TimeoutException, selenium.common.exceptions.NoSuchElementException):
        print("Not all elements located, closing session")
        IE.quit()

    # In order to add the client once all details have been entered
    add_wait = WebDriverWait(IE,10).until(EC.element_to_be_clickable((By.XPATH,'/html/body/div[2]/div[2]/md-content/section/div/div/section/form/div/div[14]/button[1]')))
    add = IE.find_element_by_xpath('/html/body/div[2]/div[2]/md-content/section/div/div/section/form/div/div[14]/button[1]')
    IE.execute_script("arguments[0].click();",add)
    time.sleep(2)

    # click small round close button
    try:
        round_close_button_wait = WebDriverWait(IE, 20).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div/div/div/div/button/span')))
        closebutton = IE.find_element_by_xpath('/html/body/div[1]/div/div/div/div/button/span')
        closebutton.click()

    except selenium.common.exceptions.TimeoutException:
        print("Timeout error, please try again")

    # click the 'close' sign
    close_button_wait = WebDriverWait(IE, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/div[2]/div/div/div[1]/div/div[1]/a/span')))
    closesign = IE.find_element_by_xpath('/html/body/div[2]/div[2]/div/div/div[1]/div/div[1]/a/span')
    closesign.click()

    # Click the funeral option in the circle
    WebDriverWait(IE, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="benefit_3_1"]')))

    @retries(max_attempts=3)
    def funeral_function():
        funerals = IE.find_elements_by_xpath('//*[@id="benefit_3_1"]')
        for funeral in funerals:
            if funeral.get_attribute('id') == "benefit_3_1":
                funeral.click()
    result = funeral_function() # This is the actual call of the function 'funeral_function' within the function itself.

    # click the plan button on the right to add details
    WebDriverWait(IE, 60).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/div[2]/section[2]/div/div[3]/div[2]/life-goals/div[1]/div/div/div[4]/div/div[2]/button[1]')))
    plan = IE.find_element_by_xpath('/html/body/div[2]/div[2]/section[2]/div/div[3]/div[2]/life-goals/div[1]/div/div/div[4]/div/div[2]/button[1]')
    plan.click()

    # click the 'ok' button to continue to next page
    try:
        WebDriverWait(IE, 120).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[3]/div/div/div/div[3]/div/button')))
        okcontinue = IE.find_element_by_xpath('/html/body/div[3]/div/div/div/div[3]/div/button')
        okcontinue.click()

    except(selenium.common.exceptions.TimeoutException):
        print("Proceeding...")

    # Setting workbook options for first benefit
    WebDriverWait(IE,60).until(EC.element_to_be_clickable((By.XPATH,'/html/body/div[3]/div/div/div/div[3]/div[2]/div[1]/table/tbody/tr/td[1]/div[2]/div/a')))
    benefit_1_option = xlrd.open_workbook(excelfile_source).sheet_by_name("AS400").cell_value(rowx=i, colx=14) # This is the actual option, i.e. MFPA or MFMM etc.
    benefit_one_sum_assured = str(xlrd.open_workbook(excelfile_source).sheet_by_name('AS400').cell_value(rowx=i,colx=31)) # The actual sum assured value
    new_benefit_1_sum_assured = change_string(benefit_one_sum_assured)
    print("First benefit: " + benefit_1_option)
    print("First benefit sum assured:" + new_benefit_1_sum_assured)
    print()
    benefit_1_indicator = False
    benefit_1_MFMM_indicator = False
    benefit_1_MFSP_indicator = False
    benefit_1_MFOC_indicator = False
    benefit_1_MFUC_indicator = False
    benefit_1_MFPA_indicator = False
    benefit_1_MFEF_indicator = False

    YOB_source = int(xlrd.open_workbook(excelfile_source).sheet_by_name("AS400").cell_value(rowx=i, colx=33))
    First_benefit_DOB = create_dob(Todays_date, YOB_source)

    if benefit_1_option == "MFMM":
        firstoption = IE.find_element_by_xpath('/html/body/div[3]/div/div/div/div[3]/div[2]/div[1]/table/tbody/tr/td[1]/div[2]/div/a')
        firstoption.click()
        benefit_1_indicator = True
        benefit_1_MFMM_indicator = True
    elif benefit_1_option == "MFSP":
        firstoption = IE.find_element_by_xpath('/html/body/div[3]/div/div/div/div[3]/div[2]/div[1]/table/tbody/tr/td[2]/div[2]/div/a')
        firstoption.click()
        textbox = IE.find_element_by_xpath('/html/body/div[3]/div/div/div/div[3]/div[2]/div[1]/table/tbody/tr/td[2]/div[2]/div/div[1]/input')
        textbox.send_keys("Happy")
        time.sleep(1)
        benefit_1_indicator = True
        benefit_1_MFSP_indicator = True
    elif benefit_1_option == "MFUC":
        firstoption = IE.find_element_by_xpath('/html/body/div[3]/div/div/div/div[3]/div[2]/div[1]/table/tbody/tr/td[3]/div[2]/div/a')
        firstoption.click()
        textbox = IE.find_element_by_xpath('/html/body/div[3]/div/div/div/div[3]/div[2]/div[1]/table/tbody/tr/td[3]/div[2]/div/div[1]/input')
        textbox.send_keys("Sad")
        time.sleep(1)
        benefit_1_indicator = True
        benefit_1_MFUC_indicator = True
    elif benefit_1_option == "MFPA":
        firstoption = IE.find_element_by_xpath('/html/body/div[3]/div/div/div/div[3]/div[2]/div[2]/table/tbody/tr/td[1]/div[2]/div/a')
        firstoption.click()
        textbox = IE.find_element_by_xpath('/html/body/div[3]/div/div/div/div[3]/div[2]/div[2]/table/tbody/tr/td[1]/div[2]/div/div[1]/input')
        textbox.send_keys("China")
        time.sleep(1)
        benefit_1_indicator = True
        benefit_1_MFPA_indicator = True
    elif benefit_1_option == "MFEF":
        firstoption = IE.find_element_by_xpath('/html/body/div[3]/div/div/div/div[3]/div[2]/div[2]/table/tbody/tr/td[2]/div[2]/div/a')
        firstoption.click()
        textbox = IE.find_element_by_xpath('/html/body/div[3]/div/div/div/div[3]/div[2]/div[2]/table/tbody/tr/td[2]/div[2]/div/div[1]/input')
        textbox.send_keys("Japan")
        time.sleep(1)
        benefit_1_indicator = True
        benefit_1_MFEF_indicator = True

    # Setting workbook options for second benefit
    benefit_2_option = xlrd.open_workbook(excelfile_source).sheet_by_name("AS400").cell_value(rowx=i, colx=52)
    print("Second benefit is: " + benefit_2_option)
    print()

    benefit_2_indicator = False
    benefit_2_PURE_indicator = False
    benefit_2_PUDT_indicator = False
    benefit_2_PUDS_indicator = False
    benefit_2_PURO_indicator = False

    if benefit_2_option == "PUDT" or "PUDS" or "PURE" or "PURO":
        benefit_2_indicator = True
        if benefit_2_option == "PUDT":
            benefit_2_PUDT_indicator = True
        elif benefit_2_option == "PUDS":
            benefit_2_PUDS_indicator = True
        elif benefit_2_option == "PURE":
            benefit_2_PURE_indicator = True
        elif benefit_2_option == "PURO":
            benefit_2_PURO_indicator = True

    #Setting workbook options for third benefit
    benefit_3_option = xlrd.open_workbook(excelfile_source).sheet_by_name("AS400").cell_value(rowx=i, colx=90)
    print("Third benefit is: " + benefit_3_option)
    print()

    benefit_3_indicator = False
    benefit_3_PURE_indicator = False
    benefit_3_PURO_indicator = False
    benefit_3_PUDT_indicator = False
    benefit_3_PUDS_indicator = False

    if benefit_3_option == "PUDT" or "PUDS" or "PURO" or "PURE":
        benefit_3_indicator = True
        if benefit_3_option == "PUDT":
            benefit_3_PUDT_indicator = True
        elif benefit_3_option == "PUDS":
            benefit_3_PUDS_indicator = True
        elif benefit_3_option == "PURE":
            benefit_3_PURE_indicator = True
        elif benefit_3_option == "PURO":
            benefit_3_PURO_indicator = True

    # Setting workbook options for Fourth benefit
    benefit_4_option = xlrd.open_workbook(excelfile_source).sheet_by_name("AS400").cell_value(rowx=i, colx=128)
    print("Fourth benefit is: " + benefit_4_option)
    print()

    benefit_4_indicator = False
    benefit_4_PURE_indicator = False
    benefit_4_PURO_indicator = False
    benefit_4_PUDT_indicator = False
    benefit_4_PUDS_indicator = False

    if benefit_4_option == "PUDT" or "PUDS" or "PURO" or "PURE":
        benefit_4_indicator = True
        if benefit_4_option == "PUDT":
            benefit_4_PUDT_indicator = True
        elif benefit_4_option == "PUDS":
            benefit_4_PUDS_indicator = True
        elif benefit_4_option == "PURE":
            benefit_4_PURE_indicator = True
        elif benefit_4_option == "PURO":
            benefit_4_PURO_indicator = True

    # click the next button to add stupid details like what comes to mind first
    nextbutton = IE.find_element_by_xpath('/html/body/div[3]/div/footer/div/div/div[4]/div/button')
    nextbutton.click()

    # Filling in all the needless details
    WebDriverWait(IE,10).until(EC.visibility_of_all_elements_located((By.XPATH,'/html/body/div[3]/div/div/div/div/div[2]/div[2]/textarea')))
    textarea = IE.find_element_by_xpath('/html/body/div[3]/div/div/div/div/div[2]/div[2]/textarea')
    textarea.send_keys('Sei still und wisse, dass er Gott ist')
    yourneed = IE.find_element_by_xpath('/html/body/div[3]/div/div/div/div/div[6]/div[1]/input')
    yourneed.send_keys(i*10 + 10000)
    importantneed = IE.find_element_by_xpath('/html/body/div[3]/div/div/div/div/div[6]/div[2]/input')
    importantneed.send_keys('Makes me happy...')
    estcost = IE.find_element_by_xpath('/html/body/div[3]/div/div/div/div/div[6]/div[3]/input')
    mysum = str(xlrd.open_workbook(excelfile_source).sheet_by_name("Scenarios").cell_value(rowx=i, colx=13))
    estcost.send_keys(mysum)
    time.sleep(1)

    # click the 'next' button
    WebDriverWait(IE,10).until(EC.element_to_be_clickable((By.XPATH,'/html/body/div[3]/div/footer/div/div/div[4]/div/button/div/div')))
    next = IE.find_element_by_xpath('/html/body/div[3]/div/footer/div/div/div[4]/div/button/div/div')
    next.click()

    # clicking on draw up a budget and filling in details
    try:
        WebDriverWait(IE,20).until(EC.element_to_be_clickable((By.XPATH,'/html/body/div[3]/div/div/div/div/div[2]/div[2]/div/div/div[2]/div/div[2]/button/div')))
        budget = IE.find_element_by_xpath('/html/body/div[3]/div/div/div/div/div[2]/div[2]/div/div/div[2]/div/div[2]/button/div')
        budget.click()
        time.sleep(1)
        your_salary = IE.find_element_by_xpath('/html/body/div[1]/div/div/budget-calculator/div/div[2]/div[2]/div/div/div/div/table/tbody/tr[1]/td[2]/div/div[1]/input')
        your_salary.send_keys(100000)
        partner_salary = IE.find_element_by_xpath('/html/body/div[1]/div/div/budget-calculator/div/div[2]/div[2]/div/div/div/div/table/tbody/tr[2]/td[2]/div/div[1]/input')
        partner_salary.send_keys(100000)
        housing = IE.find_element_by_xpath('/html/body/div[1]/div/div/budget-calculator/div/div[2]/div[4]/div/div/div/div/table/tbody/tr[1]/td[2]/div/div[1]/input')
        housing.send_keys(20000)
        telephone = IE.find_element_by_xpath('/html/body/div[1]/div/div/budget-calculator/div/div[2]/div[4]/div/div/div/div/table/tbody/tr[3]/td[2]/div/div[1]/input')
        telephone.send_keys(10000)
        clothing = IE.find_element_by_xpath('/html/body/div[1]/div/div/budget-calculator/div/div[2]/div[4]/div/div/div/div/table/tbody/tr[4]/td[2]/div/div[1]/input')
        clothing.send_keys(10000)
        children_exp = IE.find_element_by_xpath('/html/body/div[1]/div/div/budget-calculator/div/div[2]/div[4]/div/div/div/div/table/tbody/tr[5]/td[2]/div/div[1]/input')
        children_exp.send_keys(20000)
        debt = IE.find_element_by_xpath('/html/body/div[1]/div/div/budget-calculator/div/div[2]/div[4]/div/div/div/div/table/tbody/tr[6]/td[2]/div/div[1]/input')
        debt.send_keys(1000)
        transport = IE.find_element_by_xpath('/html/body/div[1]/div/div/budget-calculator/div/div[2]/div[4]/div/div/div/div/table/tbody/tr[7]/td[2]/div/div[1]/input')
        transport.send_keys(5000)
        medical = IE.find_element_by_xpath('/html/body/div[1]/div/div/budget-calculator/div/div[2]/div[4]/div/div/div/div/table/tbody/tr[8]/td[2]/div/div[1]/input')
        medical.send_keys(20000)
        time.sleep(1)
        done_click = IE.find_element_by_xpath('/html/body/div[1]/div/div/budget-calculator/div/div[3]/button')
        done_click.click()
        time.sleep(2)

    except(selenium.common.exceptions):
        print("Could not locate all elements.  You should manually continue with this scenario.")

    # Entering employment details
    package = int(xlrd.open_workbook(excelfile_source).sheet_by_name('AS400').cell_value(rowx=i, colx=26))
    print("Package currently: " + str(package))
    enter_employer_details = IE.find_element_by_xpath('/html/body/div[3]/div/div/div/div/div[2]/div[1]/div/div/button')
    enter_employer_details.click()
    time.sleep(1)
    employed = IE.find_element_by_xpath('//*[@id="inputEmploymentTypeButtonGroup"]/button[1]')
    employed.click()
    time.sleep(1)

    if package == 803:
            workplace_name = "MMI - MOMENTUM"
            workplace_textbox = IE.find_element_by_xpath('//*[@id="scrollable-dropdown-menu"]/input')
            workplace_textbox.send_keys(workplace_name)
            time.sleep(2)
            workplace_textbox.send_keys(Keys.ENTER)

            try:
                WebDriverWait(IE, 20).until(EC.element_to_be_clickable((By.CSS_SELECTOR,'body > div.modal.fade.ng-scope.ng-isolate-scope.in > div > div > div > div.modal-body > form > div > button')))
                find_scheme_button = IE.find_element_by_css_selector("body > div.modal.fade.ng-scope.ng-isolate-scope.in > div > div > div > div.modal-body > form > div > button")
                IE.execute_script("document.querySelector('body > div.modal.fade.ng-scope.ng-isolate-scope.in > div > div > div > div.modal-body > form > div > button').click();")
                time.sleep(2)

            except(selenium.common.exceptions.NoSuchElementException,selenium.common.exceptions.TimeoutException):
                print("Could not locate element, please proceed manually")

    elif package == 801:
            workplace_name = "2 MILITARY HOSPITAL"
            Trade_union_name = "CONSOLIDATED WORKERS UNION OF SOUTH AFRICA (COWUSA)"
            IE.execute_script("document.querySelector('body > div.modal.fade.ng-scope.ng-isolate-scope.in > div > div > div > div.modal-body > form > worksite > ng-form > div > div.ng-scope > div:nth-child(1) > div.mmih-m-left_medium > div > div.col-md-10 > div:nth-child(1) > label').click();")

            try:
                Trade_union_textboxs = IE.find_elements_by_tag_name('input')
                for Trade_union_textbox in Trade_union_textboxs:
                    if Trade_union_textbox.get_attribute('name') == "unionNameAndAcro":
                        Trade_union_textbox.send_keys(Trade_union_name)
                        time.sleep(2)
                        Trade_union_option = IE.find_element_by_tag_name('a')
                        IE.execute_script("arguments[0].click();", Trade_union_option)

                workplace_textboxs = IE.find_elements_by_tag_name('input')
                for workplace_textbox in workplace_textboxs:
                    if workplace_textbox.get_attribute('class') == "form-control ng-pristine ng-valid ng-empty ng-touched":
                        workplace_textbox.send_keys(workplace_name)
                        time.sleep(1)
                        workplace_textbox.send_keys(Keys.ENTER)

            except(selenium.common.exceptions.NoSuchElementException, selenium.common.exceptions.TimeoutException):
                print("Could not locate element, please proceed manually from here.")

            try:
                WebDriverWait(IE, 15).until(EC.element_to_be_clickable((By.CSS_SELECTOR,'body > div.modal.fade.ng-scope.ng-isolate-scope.in > div > div > div > div.modal-body > form > div > button')))
                find_scheme_button = IE.find_element_by_css_selector("body > div.modal.fade.ng-scope.ng-isolate-scope.in > div > div > div > div.modal-body > form > div > button")
                IE.execute_script("document.querySelector('body > div.modal.fade.ng-scope.ng-isolate-scope.in > div > div > div > div.modal-body > form > div > button').click();")
                time.sleep(1)
            except(selenium.common.exceptions.NoSuchElementException, selenium.common.exceptions.TimeoutException):
                print("Could not locate element, please proceed manually")

    elif package == 809:
            workplace_name = "KFC - BALLITO"
            workplace_textbox = IE.find_element_by_xpath('//*[@id="scrollable-dropdown-menu"]/input')
            workplace_textbox.send_keys(workplace_name)
            time.sleep(2)
            workplace_textbox.send_keys(Keys.ENTER)

            try:
                WebDriverWait(IE, 20).until(EC.element_to_be_clickable((By.TAG_NAME, 'a')))
                KFC_button = IE.find_element_by_tag_name('a')
                IE.execute_script("arguments[0].click();",KFC_button)

            except(selenium.common.exceptions.NoSuchElementException, selenium.common.exceptions.TimeoutException, selenium.common.exceptions.WebDriverException):
                print("Could not locate element, please proceed manually.")

            try:
                WebDriverWait(IE, 20).until(EC.element_to_be_clickable((By.CSS_SELECTOR,'body > div.modal.fade.ng-scope.ng-isolate-scope.in > div > div > div > div.modal-body > form > div > button')))
                find_scheme_button = IE.find_element_by_css_selector("body > div.modal.fade.ng-scope.ng-isolate-scope.in > div > div > div > div.modal-body > form > div > button")
                IE.execute_script("document.querySelector('body > div.modal.fade.ng-scope.ng-isolate-scope.in > div > div > div > div.modal-body > form > div > button').click();")

            except(selenium.common.exceptions.NoSuchElementException, selenium.common.exceptions.TimeoutException):
                print("Could not locate element, please proceed manually")

    elif package == 818:
            workplace_name = "SANDF - SALDANHA MILITARY ACADEMY"
            workplace_textbox = IE.find_element_by_xpath('//*[@id="scrollable-dropdown-menu"]/input')
            workplace_textbox.send_keys("SANDF - SALDANHA MILITARY ACADEMY")
            workplace_textbox.send_keys(Keys.ENTER)
            time.sleep(1)

            try:
                WebDriverWait(IE, 20).until(EC.element_to_be_clickable((By.CSS_SELECTOR,'body > div.modal.fade.ng-scope.ng-isolate-scope.in > div > div > div > div.modal-body > form > div > button')))
                find_scheme_button = IE.find_element_by_css_selector("body > div.modal.fade.ng-scope.ng-isolate-scope.in > div > div > div > div.modal-body > form > div > button")
                IE.execute_script("document.querySelector('body > div.modal.fade.ng-scope.ng-isolate-scope.in > div > div > div > div.modal-body > form > div > button').click();")
                time.sleep(1)

            except(selenium.common.exceptions.NoSuchElementException, selenium.common.exceptions.TimeoutException):
                print("Could not locate element, please proceed manually")

    try:
        WebDriverWait(IE,20).until(EC.element_to_be_clickable((By.XPATH,'/html/body/div[3]/div/div/div/div/div[3]/div[1]/div[2]/div/div[1]/select')))
        i_am_a_member = Select(IE.find_element_by_xpath('/html/body/div[3]/div/div/div/div/div[3]/div[1]/div[2]/div/div[1]/select'))
        i_am_a_member.select_by_index(1)
        time.sleep(1)

    except(selenium.common.exceptions.NoSuchElementException):
        print("Could not locate element, please proceed manually")
        IE.refresh()


    # Selecting the correct option on for the start date
    try:
        WebDriverWait(IE,20).until(EC.element_to_be_clickable((By.XPATH,'//*[@id="startDate"]/div/div[2]')))
        drop_menu = IE.find_element_by_xpath('//*[@id="startDate"]/div/div[2]')
        drop_menu.click()
        start_date_menu = IE.find_elements_by_tag_name('a')
        for start_date_options in start_date_menu:
            if start_date_options.text == Start_date:
                start_date_options.click()

    except(selenium.common.exceptions.NoSuchElementException, selenium.common.exceptions.TimeoutException):
        print("Could not locate element/s, please proceed manually")


    # Setting default amounts for 'YOURSELF'
    try:
        WebDriverWait(IE,10).until(EC.element_to_be_clickable((By.XPATH,'/html/body/div[3]/div/div/div/div/div[5]/div[2]/insured-lives/div/div[2]/div/form/div/div[1]/div[1]/div[2]/div/date-selector/div/div/input')))
        webelement_DOB = IE.find_element_by_xpath('/html/body/div[3]/div/div/div/div/div[5]/div[2]/insured-lives/div/div[2]/div/form/div/div[1]/div[1]/div[2]/div/date-selector/div/div/input')
        webelement_DOB.clear()
        webelement_DOB.send_keys("01/01/1999")
        time.sleep(1)
        webelement_Need = IE.find_element_by_xpath('/html/body/div[3]/div/div/div/div/div[5]/div[2]/insured-lives/div/div[2]/div/form/div/div[1]/div[2]/div[1]/input')
        webelement_Need.clear()
        webelement_Need.send_keys(0)
        time.sleep(1)

    except(selenium.common.exceptions.ElementNotVisibleException, selenium.common.exceptions.TimeoutException, selenium.common.exceptions.NoSuchElementException):
        print("Default values could not be completed, please proceed manually")
        print()

    try:
        WebDriverWait(IE,20).until(EC.element_to_be_clickable((By.TAG_NAME,'a')))
        option_menu = IE.find_elements_by_tag_name('a')
        for option_menus in option_menu:
            if str(option_menus.text) == new_benefit_1_sum_assured:
                option_menus.click()

    except(selenium.common.exceptions.NoSuchElementException, selenium.common.exceptions.TimeoutException):
        print("Could not locate element, please proceed manually")

    # Setting benefit_1's DOB, Need, planned for need
    if benefit_1_indicator:
        print("benefit 1 indicator true.")
        if benefit_1_MFMM_indicator:
            print("The first benefit selected was: " + benefit_1_option)
            print()
            webelement_DOB = IE.find_element_by_xpath('/html/body/div[3]/div/div/div/div/div[5]/div[2]/insured-lives/div/div[2]/div/form/div/div[1]/div[1]/div[2]/div/date-selector/div/div/input')
            webelement_DOB.clear()
            webelement_DOB.send_keys(First_benefit_DOB)
            webelement_Need = IE.find_element_by_xpath('/html/body/div[3]/div/div/div/div/div[5]/div[2]/insured-lives/div/div[2]/div/form/div/div[1]/div[2]/div[1]/input')
            webelement_Need.clear()
            webelement_Need.send_keys(new_benefit_1_sum_assured)
            time.sleep(1)
            drop_menu = IE.find_element_by_xpath('/html/body/div[3]/div/div/div/div/div[5]/div[2]/insured-lives/div/div[2]/div/form/div/div[2]/div[2]/div/div/div[2]/button')
            drop_menu.click()
            option_menu = IE.find_elements_by_tag_name('a')
            for option_menus in option_menu:
                if str(option_menus.text) == new_benefit_1_sum_assured:
                    option_menus.click()

        if benefit_1_MFEF_indicator:
            print("The first benefit selected was: " + benefit_1_option)
            print()
            webelement_DOB = IE.find_element_by_xpath('/html/body/div[3]/div/div/div/div/div[13]/div[2]/insured-lives/div/div[2]/div/form/div/div[1]/div[1]/div[2]/div/date-selector/div/div/input')
            webelement_DOB.clear()
            webelement_DOB.send_keys(First_benefit_DOB)
            webelement_Need = IE.find_element_by_xpath('/html/body/div[3]/div/div/div/div/div[13]/div[2]/insured-lives/div/div[2]/div/form/div/div[1]/div[2]/div[1]/input')
            webelement_Need.clear()
            webelement_Need.send_keys(new_benefit_1_sum_assured)
            time.sleep(1)
            drop_menu = IE.find_element_by_xpath('/html/body/div[3]/div/div/div/div/div[13]/div[2]/insured-lives/div/div[2]/div/form/div/div[2]/div[2]/div/div/div[2]/button')
            drop_menu.click()
            print("Benefit one sum assured added")
            option_menu = IE.find_elements_by_tag_name('a')
            for option_menus in option_menu:
                if str(option_menus.text) == new_benefit_1_sum_assured:
                    option_menus.click()

        if benefit_1_MFPA_indicator:
            print("The first benefit selected was: " + benefit_1_option)
            print()
            webelement_DOB = IE.find_element_by_xpath('/html/body/div[3]/div/div/div/div/div[11]/div[2]/insured-lives/div/div[2]/div/form/div/div[1]/div[1]/div[2]/div/date-selector/div/div/input')
            webelement_DOB.clear()
            webelement_DOB.send_keys(First_benefit_DOB)
            webelement_Need = IE.find_element_by_xpath('/html/body/div[3]/div/div/div/div/div[11]/div[2]/insured-lives/div/div[2]/div/form/div/div[1]/div[2]/div[1]/input')
            webelement_Need.clear()
            webelement_Need.send_keys(new_benefit_1_sum_assured)
            time.sleep(1)
            drop_menu = IE.find_element_by_xpath('/html/body/div[3]/div/div/div/div/div[11]/div[2]/insured-lives/div/div[2]/div/form/div/div[2]/div[2]/div/div/div[2]/button')
            drop_menu.click()
            print("Benefit one sum assured added")
            option_menu = IE.find_elements_by_tag_name('a')
            for option_menus in option_menu:
                if str(option_menus.text) == new_benefit_1_sum_assured:
                    option_menus.click()

        if benefit_1_MFSP_indicator:
            print("The first benefit selected was: " + benefit_1_option)
            print()
            webelement_DOB = IE.find_element_by_xpath('/html/body/div[3]/div/div/div/div/div[7]/div[2]/insured-lives/div/div[2]/div/form/div/div[1]/div[1]/div[2]/div/date-selector/div/div/input')
            webelement_DOB.clear()
            webelement_DOB.send_keys(First_benefit_DOB)
            webelement_Need = IE.find_element_by_xpath('/html/body/div[3]/div/div/div/div/div[7]/div[2]/insured-lives/div/div[2]/div/form/div/div[1]/div[2]/div[1]/input')
            webelement_Need.clear()
            webelement_Need.send_keys(new_benefit_1_sum_assured)
            time.sleep(1)
            drop_menu = IE.find_element_by_xpath('/html/body/div[3]/div/div/div/div/div[7]/div[2]/insured-lives/div/div[2]/div/form/div/div[2]/div[2]/div/div/div[2]/button')
            drop_menu.click()
            print("Benefit one sum assured added")
            option_menu = IE.find_elements_by_tag_name('a')
            for option_menus in option_menu:
                if str(option_menus.text) == new_benefit_1_sum_assured:
                    option_menus.click()

        if benefit_1_MFUC_indicator:
            print("The first benefit selected was: " + benefit_1_option)
            print()
            webelement_DOB = IE.find_element_by_xpath('/html/body/div[3]/div/div/div/div/div[9]/div[2]/insured-lives/div/div[2]/div/div[1]/div/form/div[1]/div[2]/div/date-selector/div/div/input')
            webelement_DOB.clear()
            webelement_DOB.send_keys(First_benefit_DOB)
            webelement_Need = IE.find_element_by_xpath('/html/body/div[3]/div/div/div/div/div[9]/div[2]/insured-lives/div/div[2]/div/div[1]/div/form/div[2]/div[1]/input')
            webelement_Need.clear()
            webelement_Need.send_keys(new_benefit_1_sum_assured)
            time.sleep(1)
            drop_menu = IE.find_element_by_css_selector("body > div:nth-child(4) > div > div > div > div > div:nth-child(10) > div.col-sm-10.col-md-11.mmih-p-horizontal_extra-small.mmih-p-vertical_extra-small > insured-lives > div > div.mmih-m-around_none.mmih-p-around_none > div > div.mmih-p-around_none.mmih-m-around_none.ng-scope.col-md-3 > div.col-md-4.mmih-p-horizontal_extra-small.ng-scope > div > div > div.col-md-1.mmih-p-around_none.mmih-m-around_none > button")
            IE.execute_script("document.querySelector('body > div:nth-child(4) > div > div > div > div > div:nth-child(10) > div.col-sm-10.col-md-11.mmih-p-horizontal_extra-small.mmih-p-vertical_extra-small > insured-lives > div > div.mmih-m-around_none.mmih-p-around_none > div > div.mmih-p-around_none.mmih-m-around_none.ng-scope.col-md-3 > div.col-md-4.mmih-p-horizontal_extra-small.ng-scope > div > div > div.col-md-1.mmih-p-around_none.mmih-m-around_none > button').click();")
            print("Benefit one sum assured added")
            option_menu = IE.find_elements_by_tag_name('a')
            for option_menus in option_menu:
                if str(option_menus.text) == new_benefit_1_sum_assured:
                    option_menus.click()

    # Setting benefit_2's DOB, Need, planned for need
    if benefit_2_indicator:
        print()

        if benefit_2_PUDT_indicator:
            buttons = IE.find_elements_by_xpath('/html/body/div[3]/div/div/div/div/div[15]/div[2]/div[2]/div[2]/div[1]/benefit/div/div[1]/div/label')
            for button in buttons:
                if button.get_attribute('for') == "benefit200" or "benefit189" or "benefit135":
                    IE.execute_script("arguments[0].click();",button)
                    print("Add on - Death benefit selected")

        if benefit_2_PUDS_indicator:
            buttons = IE.find_elements_by_xpath('/html/body/div[3]/div/div/div/div/div[15]/div[2]/div[2]/div[2]/div[1]/benefit/div/div[1]/div/label')
            for button in buttons:
                if button.get_attribute('for') == "benefit200" or "benefit189" or "benefit135":
                    IE.execute_script("arguments[0].click();", button)
                    print("Add on - Disability benefit selected")

        if benefit_2_PURE_indicator:
            buttons = IE.find_elements_by_xpath('/html/body/div[3]/div/div/div/div/div[15]/div[2]/div[2]/div[2]/div[2]/benefit/div/div[1]/div/label')
            for button in buttons:
                if button.get_attribute('for') == "benefit202" or "benefit137" or "benefit1483" or "benefit1473":
                    IE.execute_script("arguments[0].click();", button)
                    print("Add on - Retirement benefit selected")

        if benefit_2_PURO_indicator:
            buttons = IE.find_elements_by_xpath('/html/body/div[3]/div/div/div/div/div[15]/div[2]/div[2]/div[2]/div[2]/benefit/div/div[1]/div/label')
            for button in buttons:
                if button.get_attribute('for') == "benefit202" or "benefit137" or "benefit1483" or "benefit1473":
                    IE.execute_script("arguments[0].click();", button)
                    print("Add on - Retirement benefit selected")

    # Setting benefit_3's DOB, Need, planned for need
    if benefit_3_indicator:
        if benefit_3_PUDT_indicator:
            if benefit_2_PUDS_indicator:
                print("PUDT/PUDS already selected.")
            else:
                buttons = IE.find_elements_by_xpath('/html/body/div[3]/div/div/div/div/div[15]/div[2]/div[2]/div[2]/div[1]/benefit/div/div[1]/div/label')
                for button in buttons:
                    if button.get_attribute('for') == "benefit200" or "benefit189" or "benefit135":
                        IE.execute_script("arguments[0].click();", button)
                        print("Add on - Death benefit selected")

        if benefit_3_PUDS_indicator:
            if benefit_2_PUDT_indicator:
                print("PUDS/PUDT already selected")
            else:
                buttons = IE.find_elements_by_xpath('/html/body/div[3]/div/div/div/div/div[15]/div[2]/div[2]/div[2]/div[1]/benefit/div/div[1]/div/label')
                for button in buttons:
                    if button.get_attribute('for') == "benefit200" or "benefit189" or "benefit135":
                        IE.execute_script("arguments[0].click();", button)
                        print("Add on - Disability benefit selected")

        if benefit_3_PURE_indicator:
            buttons = IE.find_elements_by_xpath('/html/body/div[3]/div/div/div/div/div[15]/div[2]/div[2]/div[2]/div[2]/benefit/div/div[1]/div/label')
            for button in buttons:
                if button.get_attribute('for') == "benefit202" or "benefit137" or "benefit1483" or "benefit1473":
                    IE.execute_script("arguments[0].click();", button)
                    print("Add on - Retirement benefit selected")

        if benefit_3_PURO_indicator:
            buttons = IE.find_elements_by_xpath('/html/body/div[3]/div/div/div/div/div[15]/div[2]/div[2]/div[2]/div[2]/benefit/div/div[1]/div/label')
            for button in buttons:
                if button.get_attribute('for') == "benefit202" or "benefit137" or "benefit1483" or "benefit1473":
                    IE.execute_script("arguments[0].click();", button)
                    print("Add on - Retirement benefit selected")

    # Setting benefit_4 DOB, Need, planned for need
    if benefit_4_indicator:
        if benefit_4_PUDT_indicator:
            if benefit_3_PUDS_indicator or benefit_2_PUDS_indicator:
                print("PUDS/PUDT already selected..")
            else:
                buttons = IE.find_elements_by_xpath('/html/body/div[3]/div/div/div/div/div[15]/div[2]/div[2]/div[2]/div[1]/benefit/div/div[1]/div/label')
                for button in buttons:
                    if button.get_attribute('for') == "benefit200" or "benefit189" or "benefit135":
                        IE.execute_script("arguments[0].click();", button)
                        print("Add on - Death benefit selected")

        if benefit_4_PUDS_indicator:
            if benefit_3_PUDT_indicator or benefit_2_PUDS_indicator:
                print("PUDS/PUDT already selected...")
            else:
                buttons = IE.find_elements_by_xpath('/html/body/div[3]/div/div/div/div/div[15]/div[2]/div[2]/div[2]/div[1]/benefit/div/div[1]/div/label')
                for button in buttons:
                    if button.get_attribute('for') == "benefit200" or "benefit189" or "benefit135":
                        IE.execute_script("arguments[0].click();", button)
                        print("Add on - Disability benefit selected")

        if benefit_4_PURE_indicator:
            buttons = IE.find_elements_by_xpath('/html/body/div[3]/div/div/div/div/div[15]/div[2]/div[2]/div[2]/div[2]/benefit/div/div[1]/div/label')
            for button in buttons:
                if button.get_attribute('for') == "benefit202" or "benefit137" or "benefit1483" or "benefit1473":
                    IE.execute_script("arguments[0].click();", button)
                    print("Add on - Retirement benefit selected")

        if benefit_4_PURO_indicator:
            buttons = IE.find_elements_by_xpath('/html/body/div[3]/div/div/div/div/div[15]/div[2]/div[2]/div[2]/div[2]/benefit/div/div[1]/div/label')
            for button in buttons:
                if button.get_attribute('for') == "benefit202" or "benefit137" or "benefit1483" or "benefit1473":
                    IE.execute_script("arguments[0].click();", button)
                    print("Add on - Retirement benefit selected")

    # Click the calculate button
    time.sleep(2)
    calculate_button = IE.find_element_by_xpath('/html/body/div[3]/div/footer/div/div/div[3]/div/div/div[1]/button')
    calculate_button.click()
    print()
    print("Calculating...")
    total_premium = IE.find_element_by_xpath('/html/body/div[3]/div/footer/div/div/div[3]/div/div/div[3]/span[2]')

    # A loop to ensure calculation is finished before we proceed
    while total_premium.text == "R 0":
            time.sleep(1)
            if total_premium.text != "R 0":
                print()
                print("Total Premium: " + total_premium.text)
                time.sleep(2)

    # Setting newbus premium for benefit One
    if benefit_1_indicator:
        print()
        print("Now getting benefit 1 new business premium")
        print()
        if benefit_1_MFMM_indicator:
            web_MFMM_prem = IE.find_elements_by_xpath('/html/body/div[3]/div/div/div/div/div[5]/div[2]/insured-lives/div/div[2]/div/form/div/div[2]/div[3]')
            for element1 in web_MFMM_prem:
                mfmm_newbus = str(element1.text)
                print(mfmm_newbus)
                outputworksheet.range(k, 34).value = mfmm_newbus
                print("Done for benefit 1:")
                print()

        if benefit_1_MFSP_indicator:
            web_MFSP_prem = IE.find_elements_by_xpath('/html/body/div[3]/div/div/div/div/div[7]/div[2]/insured-lives/div/div[2]/div/form/div/div[2]/div[3]')
            for element2 in web_MFSP_prem:
                mfsp_newbus = str(element2.text)
                print(str(element2.text))
                outputworksheet.range(k, 34).value = mfsp_newbus
                print("Done for benefit 1:")
                print()

        if benefit_1_MFUC_indicator:
            web_MFUC_prem = IE.find_elements_by_xpath('/html/body/div[3]/div/div/div/div/div[9]/div[2]/insured-lives/div/div[2]/div/div[2]/div[3]')
            for element3 in web_MFUC_prem:
                mfuc_newbus = str(element3.text)
                print(mfuc_newbus)
                outputworksheet.range(k, 34).value = mfuc_newbus
                print("Done for benefit 1:")
                print()

        if benefit_1_MFPA_indicator:
            web_MFPA_prem = IE.find_elements_by_xpath('/html/body/div[3]/div/div/div/div/div[11]/div[2]/insured-lives/div/div[2]/div/form/div/div[2]/div[3]')
            for element4 in web_MFPA_prem:
                mfpa_newbus = str(element4.text)
                print(mfpa_newbus)
                outputworksheet.range(k, 34).value = mfpa_newbus
                print("Done for benefit 1:")
                print()

        if benefit_1_MFEF_indicator:
            web_MFEF_prem = IE.find_elements_by_xpath('/html/body/div[3]/div/div/div/div/div[13]/div[2]/insured-lives/div/div[2]/div/form/div/div[2]/div[3]')
            for element5 in web_MFEF_prem:
                mfef_newbus = str(element5.text)
                print(mfef_newbus)
                outputworksheet.range(k, 34).value = mfef_newbus
                print("Done for benefit 1:")
                print()

    # Setting newbus premium for benefit Two
    if benefit_2_indicator:
        print("Now getting benefit 2 new business premium")
        print()

        if benefit_2_PUDT_indicator:
            web_Death_Disability_prem = IE.find_element_by_xpath('/html/body/div[3]/div/div/div/div/div[15]/div[2]/div[2]/div[2]/div[1]/benefit/div/div[2]/div[3]/div/div[2]')
            death_disability_newbus = str(web_Death_Disability_prem.text)
            print(death_disability_newbus)
            outputworksheet.range(k, 73).value = death_disability_newbus
            print("Done for benefit 2:")
            print()

        if benefit_2_PUDS_indicator:
            web_Death_Disability_prem = IE.find_element_by_xpath('/html/body/div[3]/div/div/div/div/div[15]/div[2]/div[2]/div[2]/div[1]/benefit/div/div[2]/div[3]/div/div[2]')
            death_disability_newbus = str(web_Death_Disability_prem.text)
            print(death_disability_newbus)
            outputworksheet.range(k, 112).value = death_disability_newbus
            print("Done for benefit 2:")
            print()

        if benefit_2_PURE_indicator:
            web_Retirement_prem = IE.find_element_by_xpath('/html/body/div[3]/div/div/div/div/div[15]/div[2]/div[2]/div[2]/div[2]/benefit/div/div[1]/div/label')
            retirement_newbus = str(web_Retirement_prem.text)
            print(retirement_newbus)
            outputworksheet.range(k, 73).value = retirement_newbus
            print("Done for benefit 2:")
            print()

        if benefit_2_PURO_indicator:
            web_Retirement_prem = IE.find_element_by_xpath('/html/body/div[3]/div/div/div/div/div[15]/div[2]/div[2]/div[2]/div[2]/benefit/div/div[1]/div/label')
            retirement_newbus = str(web_Retirement_prem.text)
            print(retirement_newbus)
            outputworksheet.range(k, 73).value = retirement_newbus
            print("Done for benefit 2:")
            print()

    # Setting newbus premium for benefit Three
    if benefit_3_indicator:
        print("Now getting benefit 3 new business premium")
        print()

        if benefit_3_PUDT_indicator:
            web_Death_Disability_prem = IE.find_element_by_xpath('/html/body/div[3]/div/div/div/div/div[15]/div[2]/div[2]/div[2]/div[1]/benefit/div/div[2]/div[3]/div/div[2]')
            death_disability_newbus = str(web_Death_Disability_prem.text)
            print(death_disability_newbus)
            outputworksheet.range(k, 112).value = death_disability_newbus
            print("Done for benefit 3:")
            print()

        if benefit_3_PUDS_indicator:
            web_Death_Disability_prem = IE.find_element_by_xpath('/html/body/div[3]/div/div/div/div/div[15]/div[2]/div[2]/div[2]/div[1]/benefit/div/div[2]/div[3]/div/div[2]')
            death_disability_newbus = str(web_Death_Disability_prem.text)
            print(death_disability_newbus)
            outputworksheet.range(k, 112).value = death_disability_newbus
            print("Done for benefit 3:")
            print()

        if benefit_3_PURE_indicator:
            web_Retirement_prem = IE.find_element_by_xpath('/html/body/div[3]/div/div/div/div/div[15]/div[2]/div[2]/div[2]/div[2]/benefit/div/div[1]/div/label')
            retirement_newbus = str(web_Retirement_prem.text)
            print(retirement_newbus)
            outputworksheet.range(k, 112).value = retirement_newbus
            print("Done for benefit 3:")
            print()

        if benefit_3_PURO_indicator:
            web_Retirement_prem = IE.find_element_by_xpath('/html/body/div[3]/div/div/div/div/div[15]/div[2]/div[2]/div[2]/div[2]/benefit/div/div[1]/div/label')
            retirement_newbus = str(web_Retirement_prem.text)
            print(retirement_newbus)
            outputworksheet.range(k, 112).value = retirement_newbus
            print("Done for benefit 3:")
            print()

    # Setting newbus premium for benefit Four
    if benefit_4_indicator:
        print("Now getting benefit 4 new business premium")
        print()

        if benefit_4_PUDT_indicator:
            web_Death_Disability_prem = IE.find_element_by_xpath('/html/body/div[3]/div/div/div/div/div[15]/div[2]/div[2]/div[2]/div[1]/benefit/div/div[2]/div[3]/div/div[2]')
            death_disability_newbus = str(web_Death_Disability_prem.text)
            print(death_disability_newbus)
            outputworksheet.range(k, 151).value = death_disability_newbus
            print("Done for benefit 4:")
            print()

        if benefit_4_PUDS_indicator:
            web_Death_Disability_prem = IE.find_element_by_xpath('/html/body/div[3]/div/div/div/div/div[15]/div[2]/div[2]/div[2]/div[1]/benefit/div/div[2]/div[3]/div/div[2]')
            death_disability_newbus = str(web_Death_Disability_prem.text)
            print(death_disability_newbus)
            outputworksheet.range(k, 112).value = death_disability_newbus
            print("Done for benefit 4:")
            print()

        if benefit_4_PURE_indicator:
            web_Retirement_prem = IE.find_element_by_xpath('/html/body/div[3]/div/div/div/div/div[15]/div[2]/div[2]/div[2]/div[2]/benefit/div/div[1]/div/label')
            retirement_newbus = str(web_Retirement_prem.text)
            print(retirement_newbus)
            outputworksheet.range(k, 151).value = retirement_newbus
            print("Done for benefit 4:")
            print()

        if benefit_4_PURO_indicator:
            web_Retirement_prem = IE.find_element_by_xpath('/html/body/div[3]/div/div/div/div/div[15]/div[2]/div[2]/div[2]/div[2]/benefit/div/div[1]/div/label')
            retirement_newbus = str(web_Retirement_prem.text)
            print(retirement_newbus)
            outputworksheet.range(k, 151).value = retirement_newbus
            print("Done for benefit 4:")
            print()

    outputworkbook.save()
    time.sleep(3)

    # Click the next button
    next_button = IE.find_element_by_xpath('/html/body/div[3]/div/footer/div/div/div[4]/div/button/div/div')
    next_button.click()
    time.sleep(3)

    # Clicking a series of buttons
    try:
        WebDriverWait(IE,10).until(EC.element_to_be_clickable((By.TAG_NAME,'input')))
        i_am_satisfied_button = IE.find_element_by_css_selector('#section_10 > div:nth-child(2) > div:nth-child(2) > div > div > div > div.col-md-10 > label')
        IE.execute_script("arguments[0].click();",i_am_satisfied_button)
        time.sleep(1)

    except(selenium.common.exceptions):
        print("Could not click radio button.  Proceed manually")

    try:
        WebDriverWait(IE,20).until(EC.url_matches('https://retail-dev.metropolitan.co.za/funeral-planner/review-confirm'))
        WebDriverWait(IE,20).until(EC.element_to_be_clickable((By.XPATH,'//*[@id="section_80"]/div[2]/div[1]/div/div/div[1]/span[2]')))
        IE.execute_script("document.querySelector('#section_40 > div:nth-child(2) > div:nth-child(1) > div > div > div.col-md-2 > span:nth-child(1)').click();")
        IE.execute_script("document.querySelector('#section_40 > div:nth-child(3) > div:nth-child(1) > div > div > div.col-md-2 > span:nth-child(1)').click();")
        IE.execute_script("document.querySelector('#section_40 > div:nth-child(4) > div:nth-child(1) > div > div > div.col-md-2 > span:nth-child(1)').click();")
        IE.execute_script("document.querySelector('#section_40 > div:nth-child(5) > div:nth-child(1) > div > div > div.col-md-2 > span:nth-child(1)').click();")
        IE.execute_script("document.querySelector('#section_70 > div:nth-child(2) > div:nth-child(1) > div > div > div.col-md-2 > span:nth-child(1)').click();")
        IE.execute_script("document.querySelector('#section_70_question_2').click();")
        IE.execute_script("document.querySelector('#section_70 > div:nth-child(5) > div:nth-child(1) > div > div > div.col-md-2 > span:nth-child(1)').click();")
        IE.execute_script("document.querySelector('#section_80 > div:nth-child(3) > div:nth-child(1) > div > div > div.col-md-2 > span:nth-child(2)').click();")
        IE.execute_script("document.querySelector('body > div:nth-child(4) > div > div > div > div > div.row.gutter.button-row.review-confirm-container > div > button.btn.btn--mmih.btn-primary--mmih.review-and-confirm-btn-save').click()")
        print("Saved...")
        time.sleep(3)
        IE.execute_script("document.querySelector('body > div:nth-child(4) > div > div > div > div > div.row.gutter.button-row.review-confirm-container > div > button.btn.btn--mmih.btn-primary--mmih.review-and-confirm-btn-default').click();")
        print("Apply clicked...")
        WebDriverWait(IE, 60).until(EC.url_contains("https://retail-dev.metropolitan.co.za/funeral-planner/payment-details"))

        # click the debit order button and set the reason why
        reason_for_debit_order = Select(IE.find_element_by_xpath('//*[@id="reasonForDebitOrder"]'))
        reason_for_debit_order.select_by_index(1)
        debit_order_buttons = IE.find_elements_by_tag_name('label')
        for debit_order_button in debit_order_buttons:
            if debit_order_button.get_attribute('for') == "rbtn_debitorder":
                IE.execute_script("arguments[0].click();",debit_order_button)

    except(selenium.common.exceptions.TimeoutException):
        print("Timeouterror, please proceed manually")

    # Setting account details for debit order and policy number
    time.sleep(1)
    policy_number = IE.find_element_by_xpath('/html/body/div[3]/div/div/div/div[1]/div/div/div/div/label')
    WebDriverWait(IE,20).until(EC.visibility_of_element_located((By.XPATH,'/html/body/div[3]/div/div/div/div[1]/div/div/div/div/label'))) # Waits until the element in 'policy number' is located and is visible
    policy_number_excel = str(policy_number.text)
    print(policy_number_excel)
    outputworksheet.range(k, 8).value = policy_number_excel
    outputworkbook.save()
    time.sleep(2)

    bank = "UBANK"
    accountnr = "00884937279"
    bank_textbox = IE.find_element_by_xpath('//*[@id="bankInput"]')
    bank_textbox.send_keys(bank)
    bank_textbox.send_keys(Keys.TAB)
    time.sleep(1)
    accountnr_textbox = IE.find_element_by_xpath('//*[@id="accountNumberInput"]')
    accountnr_textbox.send_keys(accountnr)
    time.sleep(1)
    IE.execute_script("document.querySelector('#accountTypeInput > div > button:nth-child(1)').click();")

    # Input the deduction date
    salary_day_textbox = IE.find_element_by_xpath('//*[@id="salaryDateInput"]')
    salary_day_textbox.send_keys(1)
    #salary_day_textbox.send_keys(add_some_days(Start_date))

    # Click the next button
    next_button_2 = IE.find_element_by_xpath('/html/body/div[3]/div/footer/div/div/div[4]/div/button/div/div')
    next_button_2.click()

    # Next page, click married
    try:
        married_status_wait = WebDriverWait(IE,15).until(EC.visibility_of_element_located((By.XPATH,'//*[@id="input-marital-status"]/button[2]')))
        IE.execute_script("document.querySelector('#input-marital-status > button:nth-child(2)').click();")

        # Set preferred comm as e-mail
        IE.execute_script("document.querySelector('#emailBtn').click();")
        time.sleep(1)
        email_address = "user@momentum.co.za"
        email_address_web = IE.find_element_by_xpath('//*[@id="inputEmailAddress"]')
        email_address_web.send_keys(email_address)
        time.sleep(1)

        # Set address type as physical and the rest of the details
        physical_address_s = IE.find_elements_by_tag_name('button')
        for address in physical_address_s:
            if address.get_attribute('ng-click') == "addressDetailsCtrl.changeAddressType('Physical')":
                address.click()
                time.sleep(1)

        street = "256"
        phone = "0829994444"
        suburb = "Centurion"
        town = "Pretoria"
        postal_code = "0185"
        street_web = IE.find_element_by_xpath('//*[@id="inputStreetNumber"]')
        street_web.send_keys(street)
        suburb_web = IE.find_element_by_xpath('//*[@id="inputSuburb"]/input')
        suburb_web.send_keys(suburb)
        town_web = IE.find_element_by_xpath('//*[@id="inputCity"]/input')
        town_web.send_keys(town)
        postal_code_web = IE.find_element_by_xpath('//*[@id="inputPostalCode"]/input')
        postal_code_web.send_keys(postal_code)
        province = Select(IE.find_element_by_xpath('/html/body/div[3]/div/div/div[2]/form/div[4]/address-details/physical-address/div[7]/div/select'))
        province.select_by_index(index=3)
        phone_web = IE.find_element_by_xpath('//*[@id="inputCell"]')
        phone_web.send_keys(phone)
        time.sleep(1)

    except(TimeoutError, selenium.common.exceptions.ElementNotVisibleException):
        print("Could not locate all elements, refreshing page")
        IE.refresh()

    # Click next button
    IE.execute_script("document.querySelector('body > div:nth-child(4) > div > footer > div > div > div:nth-child(4) > div > button > div > div').click();")


    class too_many_options:

        info_button = None
        info_button_2 = None
        info_button_3 = None
        info_button_4 = None
        i1 = False
        i2 = False
        i3 = False
        i4 = False


    # Click first information required button
    WebDriverWait(IE, 60).until(EC.url_contains("https://retail-dev.metropolitan.co.za/funeral-planner/people-on-your-plan"))
    time.sleep(1)

    try:
        i1 = True
        info_required_one = IE.find_element_by_css_selector('#LifePartner-1 > div:nth-child(4) > div.col-md-6.mmih-p-around_none > button > span')

    except(selenium.common.exceptions):
        info_required_one = None
        i1 = False

    if info_required_one != None and i1 == True:
        info_buttons = IE.find_elements_by_css_selector('#LifePartner-1 > div:nth-child(4) > div.col-md-6.mmih-p-around_none > button > span')
        for too_many_options.info_button in info_buttons:
            if too_many_options.info_button.get_attribute('class') == "glyphicon glyphicon-pencil":
                too_many_options.info_button.click()
                time.sleep(1)
                surname_web = IE.find_element_by_xpath('//*[@id="inputSurname"]')
                surname_web.send_keys(surname_workbook)
                time.sleep(1)

                if workbook_gender == "Male":
                    gender_web = IE.find_element_by_css_selector('#inputGender > button:nth-child(1)')
                    gender_web.click()
                    time.sleep(1)
                elif workbook_gender == "Female":
                    gender_web = IE.find_element_by_css_selector('#inputGender > button:nth-child(2)')
                    gender_web.click()
                    time.sleep(1)

                # Adding benef's
                IE.execute_script("document.querySelector('#chk_trustedOne').click();")
                IE.execute_script("document.querySelector('#chk_ownership').click();")
                IE.execute_script("document.querySelector('#chk_beneficiary').click();")
                time.sleep(1)

                # Setting % split
                percentage_split = IE.find_element_by_xpath('//*[@id="share"]')
                percentage_split.send_keys("100")
                percentage_split.send_keys(Keys.TAB)
                time.sleep(1)

                # Click 'save' button
                IE.execute_script("document.querySelector('body > div.modal.insured-lives.fade.ng-scope.ng-isolate-scope.in > div > div > edit-insured-life > form > div.modal-footer.mmih-color-background-catskill-white > button.btn.btn--mmih.btn-primary--mmih.right-button').click();")
                time.sleep(2)

    # Click second info required button
    try:
        i2 = True
        info_required_two = IE.find_elements_by_css_selector('#Child-1 > div:nth-child(4) > div.col-md-6.mmih-p-around_none > button > span') or IE.find_elements_by_css_selector('#Child-2 > div:nth-child(4) > div.col-md-6.mmih-p-around_none > button > span') or IE.find_elements_by_css_selector('#Child-3 > div:nth-child(4) > div.col-md-6.mmih-p-around_none > button > span') or IE.find_elements_by_css_selector('#Child-4 > div:nth-child(4) > div.col-md-6.mmih-p-around_none > button > span')

    except(selenium.common.exceptions):
        i2 = False
        info_required_two = None

    if info_required_two != None and i2 == True:
        info_buttons_two = info_required_two
        for too_many_options.info_button_2 in info_buttons_two:
            if i2 and i1:
                if too_many_options.info_button_2.__hash__() != too_many_options.info_button.__hash__():
                    too_many_options.info_button_2.click()
                    time.sleep(2)
                    surname_web = IE.find_element_by_xpath('//*[@id="inputSurname"]')
                    surname_web.send_keys(surname_workbook)
                    time.sleep(1)

                    if workbook_gender == "Male":
                        IE.execute_script("document.querySelector('#inputGender > button:nth-child(1)').click();")
                    elif workbook_gender == "Female":
                        IE.execute_script("document.querySelector('#inputGender > button:nth-child(2)').click();")

                    # Click 'save' button
                    IE.execute_script("document.querySelector('body > div.modal.insured-lives.fade.ng-scope.ng-isolate-scope.in > div > div > edit-insured-life > form > div.modal-footer.mmih-color-background-catskill-white > button.btn.btn--mmih.btn-primary--mmih.right-button').click();")
                    time.sleep(1)
            elif i2:
                too_many_options.info_button_2.click()
                time.sleep(2)
                surname_web = IE.find_element_by_xpath('//*[@id="inputSurname"]')
                surname_web.send_keys(surname_workbook)
                time.sleep(1)

                if workbook_gender == "Male":
                    IE.execute_script("document.querySelector('#inputGender > button:nth-child(1)').click();")
                elif workbook_gender == "Female":
                    IE.execute_script("document.querySelector('#inputGender > button:nth-child(2)').click();")

                IE.execute_script("document.querySelector('#chk_trustedOne').click();")
                IE.execute_script("document.querySelector('#chk_ownership').click();")
                IE.execute_script("document.querySelector('#chk_beneficiary').click();")
                time.sleep(1)

                # Setting % split
                percentage_split = IE.find_element_by_xpath('//*[@id="share"]')
                percentage_split.send_keys("100")
                percentage_split.send_keys(Keys.TAB)
                time.sleep(1)

                # Click 'save' button
                IE.execute_script("document.querySelector('body > div.modal.insured-lives.fade.ng-scope.ng-isolate-scope.in > div > div > edit-insured-life > form > div.modal-footer.mmih-color-background-catskill-white > button.btn.btn--mmih.btn-primary--mmih.right-button').click();")
                time.sleep(1)

    # Click third 'info required' button
    try:
        i3 = True
        info_required_three = IE.find_elements_by_css_selector('#Parent-2 > div:nth-child(4) > div.col-md-6.mmih-p-around_none > button') or IE.find_elements_by_css_selector('#Parent-3 > div:nth-child(4) > div.col-md-6.mmih-p-around_none > button') or IE.find_elements_by_css_selector('#Parent-1 > div:nth-child(4) > div.col-md-6.mmih-p-around_none > button > span')

    except(selenium.common.exceptions):
        i3 = False
        info_required_three = None

    if info_required_three != None and i3 == True:
        info_buttons_three = info_required_three
        for too_many_options.info_button_3 in info_buttons_three:
            if i3 and i2 and i1:
                if too_many_options.info_button_3.__hash__() != too_many_options.info_button.__hash__() and too_many_options.info_button_3.__hash__() != too_many_options.info_button_2.__hash__():
                    too_many_options.info_button_3.click()

                    time.sleep(2)
                    surname_web = IE.find_element_by_xpath('//*[@id="inputSurname"]')
                    surname_web.send_keys(surname_workbook)
                    time.sleep(1)

                    if workbook_gender == "Male":
                        IE.execute_script("document.querySelector('#inputGender > button:nth-child(1)').click();")
                    elif workbook_gender == "Female":
                        IE.execute_script("document.querySelector('#inputGender > button:nth-child(2)').click();")

                        # Click 'save' button
                        IE.execute_script("document.querySelector('body > div.modal.insured-lives.fade.ng-scope.ng-isolate-scope.in > div > div > edit-insured-life > form > div.modal-footer.mmih-color-background-catskill-white > button.btn.btn--mmih.btn-primary--mmih.right-button').click();")
                        time.sleep(1)

            elif i3 and i2:
                if too_many_options.info_button_3.__hash__() != too_many_options.info_button_2.__hash__():
                    too_many_options.info_button_3.click()

                    time.sleep(2)
                    surname_web = IE.find_element_by_xpath('//*[@id="inputSurname"]')
                    surname_web.send_keys(surname_workbook)
                    time.sleep(1)

                    if workbook_gender == "Male":
                        IE.execute_script("document.querySelector('#inputGender > button:nth-child(1)').click();")
                    elif workbook_gender == "Female":
                        IE.execute_script("document.querySelector('#inputGender > button:nth-child(2)').click();")

                        # Adding benef's
                        IE.execute_script("document.querySelector('#chk_trustedOne').click();")
                        IE.execute_script("document.querySelector('#chk_ownership').click();")
                        IE.execute_script("document.querySelector('#chk_beneficiary').click();")
                        time.sleep(1)

                        # Setting % split
                        #percentage_split = IE.find_element_by_xpath('//*[@id="share"]')
                        #percentage_split.send_keys("50")
                        #percentage_split.send_keys(Keys.TAB)
                        time.sleep(1)

                        # Click 'save' button
                        IE.execute_script("document.querySelector('body > div.modal.insured-lives.fade.ng-scope.ng-isolate-scope.in > div > div > edit-insured-life > form > div.modal-footer.mmih-color-background-catskill-white > button.btn.btn--mmih.btn-primary--mmih.right-button').click();")
                        time.sleep(1)

            elif i3 and i1:
                if too_many_options.info_button_3.__hash__() != too_many_options.info_button.__hash__():
                    too_many_options.info_button_3.click()

                    time.sleep(2)
                    surname_web = IE.find_element_by_xpath('//*[@id="inputSurname"]')
                    surname_web.send_keys(surname_workbook)
                    time.sleep(1)

                    if workbook_gender == "Male":
                        IE.execute_script("document.querySelector('#inputGender > button:nth-child(1)').click();")
                    elif workbook_gender == "Female":
                        IE.execute_script("document.querySelector('#inputGender > button:nth-child(2)').click();")

                        # Click 'save' button
                        IE.execute_script("document.querySelector('body > div.modal.insured-lives.fade.ng-scope.ng-isolate-scope.in > div > div > edit-insured-life > form > div.modal-footer.mmih-color-background-catskill-white > button.btn.btn--mmih.btn-primary--mmih.right-button').click();")
                        time.sleep(1)

            elif i3:
                too_many_options.info_button_3.click()
                time.sleep(2)
                surname_web = IE.find_element_by_xpath('//*[@id="inputSurname"]')
                surname_web.send_keys(surname_workbook)
                time.sleep(1)

                if workbook_gender == "Male":
                    IE.execute_script("document.querySelector('#inputGender > button:nth-child(1)').click();")
                elif workbook_gender == "Female":
                    IE.execute_script("document.querySelector('#inputGender > button:nth-child(2)').click();")

                    # Adding benef's
                    IE.execute_script("document.querySelector('#chk_trustedOne').click();")
                    IE.execute_script("document.querySelector('#chk_ownership').click();")
                    IE.execute_script("document.querySelector('#chk_beneficiary').click();")
                    time.sleep(1)

                    # Setting % split
                    percentage_split = IE.find_element_by_xpath('//*[@id="share"]')
                    percentage_split.send_keys("100")
                    percentage_split.send_keys(Keys.TAB)
                    time.sleep(1)

                    # Click 'save' button
                    IE.execute_script("document.querySelector('body > div.modal.insured-lives.fade.ng-scope.ng-isolate-scope.in > div > div > edit-insured-life > form > div.modal-footer.mmih-color-background-catskill-white > button.btn.btn--mmih.btn-primary--mmih.right-button').click();")
                    time.sleep(1)

    # Click fourth 'info required' button
    try:
        i4 = True
        info_required_four = IE.find_elements_by_css_selector('#ExtendedFamily-2 > div:nth-child(4) > div.col-md-6.mmih-p-around_none > button') or IE.find_elements_by_css_selector('#ExtendedFamily-3 > div:nth-child(4) > div.col-md-6.mmih-p-around_none > button > span')

    except:
        i4 = False
        info_required_four = None

    if info_required_four != None and i4 == True:
        info_buttons_four = info_required_four
        for too_many_options.info_button_4 in info_buttons_four:
            if i4 and i3 and i2 and i1:
                if too_many_options.info_button_4.__hash__() != too_many_options.info_button.__hash__() and too_many_options.info_button_4.__hash__() != too_many_options.info_button_2.__hash__() and too_many_options.info_button_4.__hash__() != too_many_options.info_button_3.__hash__():
                    IE.execute_script("arguments[0].click();", too_many_options.info_button_4)
                    time.sleep(2)
                    surname_web = IE.find_element_by_xpath('//*[@id="inputSurname"]')
                    surname_web.send_keys(surname_workbook)
                    time.sleep(1)

                    if workbook_gender == "Male":
                        IE.execute_script("document.querySelector('#inputGender > button:nth-child(1)').click();")
                    elif workbook_gender == "Female":
                        IE.execute_script("document.querySelector('#inputGender > button:nth-child(2)').click();")

                relationship = Select(IE.find_element_by_xpath('/html/body/div[1]/div/div/edit-insured-life/form/personal-info/div[8]/div/div/div/select'))
                relationship.select_by_index(1)

                # Click 'save' button
                IE.execute_script("document.querySelector('body > div.modal.insured-lives.fade.ng-scope.ng-isolate-scope.in > div > div > edit-insured-life > form > div.modal-footer.mmih-color-background-catskill-white > button.btn.btn--mmih.btn-primary--mmih.right-button').click();")
                time.sleep(1)

            elif i4 and i3 and i2:
                if too_many_options.info_button_4.__hash__() != too_many_options.info_button_2.__hash__() and too_many_options.info_button_4.__hash__() != too_many_options.info_button_3.__hash__():
                    IE.execute_script("arguments[0].click();", too_many_options.info_button_4)
                    time.sleep(2)
                    surname_web = IE.find_element_by_xpath('//*[@id="inputSurname"]')
                    surname_web.send_keys(surname_workbook)
                    time.sleep(1)

                    if workbook_gender == "Male":
                        IE.execute_script("document.querySelector('#inputGender > button:nth-child(1)').click();")
                    elif workbook_gender == "Female":
                        IE.execute_script("document.querySelector('#inputGender > button:nth-child(2)').click();")

                relationship = Select(IE.find_element_by_xpath('/html/body/div[1]/div/div/edit-insured-life/form/personal-info/div[8]/div/div/div/select'))
                relationship.select_by_index(1)

                # Adding benef's
                IE.execute_script("document.querySelector('#chk_trustedOne').click();")
                IE.execute_script("document.querySelector('#chk_ownership').click();")
                IE.execute_script("document.querySelector('#chk_beneficiary').click();")
                time.sleep(1)

                # Click 'save' button
                IE.execute_script("document.querySelector('body > div.modal.insured-lives.fade.ng-scope.ng-isolate-scope.in > div > div > edit-insured-life > form > div.modal-footer.mmih-color-background-catskill-white > button.btn.btn--mmih.btn-primary--mmih.right-button').click();")
                time.sleep(1)

            elif i4 and i3 and i1:
                if too_many_options.info_button_4.__hash__() != too_many_options.info_button_3.__hash__() and too_many_options.info_button_4.__hash__() != too_many_options.info_button.__hash__():
                    IE.execute_script("arguments[0].click();", too_many_options.info_button_4)
                    time.sleep(2)
                    surname_web = IE.find_element_by_xpath('//*[@id="inputSurname"]')
                    surname_web.send_keys(surname_workbook)
                    time.sleep(1)

                    if workbook_gender == "Male":
                        IE.execute_script("document.querySelector('#inputGender > button:nth-child(1)').click();")
                    elif workbook_gender == "Female":
                        IE.execute_script("document.querySelector('#inputGender > button:nth-child(2)').click();")

                relationship = Select(IE.find_element_by_xpath('/html/body/div[1]/div/div/edit-insured-life/form/personal-info/div[8]/div/div/div/select'))
                relationship.select_by_index(1)

                # Click 'save' button
                time.sleep(1)
                save_button = IE.find_element_by_xpath('/html/body/div[1]/div/div/edit-insured-life/form/div[11]/button[2]')
                save_button.click()

            elif i4 and i2 and i1:
                if too_many_options.info_button_4.__hash__() != too_many_options.info_button_2.__hash__() and too_many_options.info_button_4.__hash__() != too_many_options.info_button.__hash__():
                    IE.execute_script("arguments[0].click();", too_many_options.info_button_4)
                    time.sleep(2)
                    surname_web = IE.find_element_by_xpath('//*[@id="inputSurname"]')
                    surname_web.send_keys(surname_workbook)
                    time.sleep(1)

                    if workbook_gender == "Male":
                        IE.execute_script("document.querySelector('#inputGender > button:nth-child(1)').click();")
                    elif workbook_gender == "Female":
                        IE.execute_script("document.querySelector('#inputGender > button:nth-child(2)').click();")

                relationship = Select(IE.find_element_by_xpath('/html/body/div[1]/div/div/edit-insured-life/form/personal-info/div[8]/div/div/div/select'))
                relationship.select_by_index(1)

                # Click 'save' button
                IE.execute_script("document.querySelector('body > div.modal.insured-lives.fade.ng-scope.ng-isolate-scope.in > div > div > edit-insured-life > form > div.modal-footer.mmih-color-background-catskill-white > button.btn.btn--mmih.btn-primary--mmih.right-button').click();")
                time.sleep(1)

            elif i4 and i1:
                if too_many_options.info_button_4.__hash__() != too_many_options.info_button.__hash__():
                    IE.execute_script("arguments[0].click();", too_many_options.info_button_4)
                    time.sleep(2)
                    surname_web = IE.find_element_by_xpath('//*[@id="inputSurname"]')
                    surname_web.send_keys(surname_workbook)
                    time.sleep(1)

                    if workbook_gender == "Male":
                        IE.execute_script("document.querySelector('#inputGender > button:nth-child(1)').click();")
                    elif workbook_gender == "Female":
                        IE.execute_script("document.querySelector('#inputGender > button:nth-child(2)').click();")

                relationship = Select(IE.find_element_by_xpath('/html/body/div[1]/div/div/edit-insured-life/form/personal-info/div[8]/div/div/div/select'))
                relationship.select_by_index(1)

                # Click 'save' button
                IE.execute_script("document.querySelector('body > div.modal.insured-lives.fade.ng-scope.ng-isolate-scope.in > div > div > edit-insured-life > form > div.modal-footer.mmih-color-background-catskill-white > button.btn.btn--mmih.btn-primary--mmih.right-button').click();")
                time.sleep(1)

            elif i4 and i2:
                if too_many_options.info_button_4.__hash__() != too_many_options.info_button_2.__hash__():
                    IE.execute_script("arguments[0].click();", too_many_options.info_button_4)
                    time.sleep(2)
                    surname_web = IE.find_element_by_xpath('//*[@id="inputSurname"]')
                    surname_web.send_keys(surname_workbook)
                    time.sleep(1)

                    if workbook_gender == "Male":
                        IE.execute_script("document.querySelector('#inputGender > button:nth-child(1)').click();")
                    elif workbook_gender == "Female":
                        IE.execute_script("document.querySelector('#inputGender > button:nth-child(2)').click();")

                relationship = Select(IE.find_element_by_xpath('/html/body/div[1]/div/div/edit-insured-life/form/personal-info/div[8]/div/div/div/select'))
                relationship.select_by_index(1)

                # Click 'save' button
                IE.execute_script("document.querySelector('body > div.modal.insured-lives.fade.ng-scope.ng-isolate-scope.in > div > div > edit-insured-life > form > div.modal-footer.mmih-color-background-catskill-white > button.btn.btn--mmih.btn-primary--mmih.right-button').click();")
                time.sleep(1)

            elif i4 and i3:
                if too_many_options.info_button_4.__hash__() != too_many_options.info_button_3.__hash__():
                    IE.execute_script("arguments[0].click();", too_many_options.info_button_4)
                    time.sleep(2)
                    surname_web = IE.find_element_by_xpath('//*[@id="inputSurname"]')
                    surname_web.send_keys(surname_workbook)
                    time.sleep(1)

                    if workbook_gender == "Male":
                        IE.execute_script("document.querySelector('#inputGender > button:nth-child(1)').click();")
                    elif workbook_gender == "Female":
                        IE.execute_script("document.querySelector('#inputGender > button:nth-child(2)').click();")

                relationship = Select(IE.find_element_by_xpath('/html/body/div[1]/div/div/edit-insured-life/form/personal-info/div[8]/div/div/div/select'))
                relationship.select_by_index(1)

                # Click 'save' button
                IE.execute_script("document.querySelector('body > div.modal.insured-lives.fade.ng-scope.ng-isolate-scope.in > div > div > edit-insured-life > form > div.modal-footer.mmih-color-background-catskill-white > button.btn.btn--mmih.btn-primary--mmih.right-button').click();")
                time.sleep(1)

            elif i4:
                IE.execute_script("arguments[0].click();", too_many_options.info_button_4)
                time.sleep(2)
                surname_web = IE.find_element_by_xpath('//*[@id="inputSurname"]')
                surname_web.send_keys(surname_workbook)
                time.sleep(1)

                if workbook_gender == "Male":
                    IE.execute_script("document.querySelector('#inputGender > button:nth-child(1)').click();")
                elif workbook_gender == "Female":
                    IE.execute_script("document.querySelector('#inputGender > button:nth-child(2)').click();")

            relationship = Select(IE.find_element_by_xpath('/html/body/div[1]/div/div/edit-insured-life/form/personal-info/div[8]/div/div/div/select'))
            relationship.select_by_index(1)

            # Adding benef's
            IE.execute_script("document.querySelector('#chk_trustedOne').click();")
            IE.execute_script("document.querySelector('#chk_ownership').click();")
            IE.execute_script("document.querySelector('#chk_beneficiary').click();")
            time.sleep(1)

            # Setting % split
            percentage_split = IE.find_element_by_xpath('//*[@id="share"]')
            percentage_split.send_keys("100")
            percentage_split.send_keys(Keys.TAB)
            time.sleep(1)

            # Click 'save' button
            IE.execute_script("document.querySelector('body > div.modal.insured-lives.fade.ng-scope.ng-isolate-scope.in > div > div > edit-insured-life > form > div.modal-footer.mmih-color-background-catskill-white > button.btn.btn--mmih.btn-primary--mmih.right-button').click();")
            time.sleep(1)

    # Calling class to execute
    try:
        too_many_options()

    except(selenium.common.exceptions.NoSuchElementException):
        time.sleep(3)

    # Click 'next' button
    next = IE.find_element_by_xpath('/html/body/div[3]/div/footer/div/div/div[4]/div/button')
    next.click()

    # Next page
    WebDriverWait(IE, 20).until(EC.url_contains("https://retail-dev.metropolitan.co.za/funeral-planner/declaration"))
    time.sleep(2)

    # click you have not received the user guide
    no = IE.find_element_by_xpath('//*[@id="section_20"]/div[2]/div[1]/div/div/div[1]/span[2]')
    no.click()

    # click 'no' for marketing option one
    no = IE.find_element_by_xpath('//*[@id="section_90"]/div[2]/div[1]/div/div/div[1]/span[2]')
    no.click()

    # click 'no' for marketing option two
    no = IE.find_element_by_xpath('//*[@id="section_90"]/div[3]/div[1]/div/div/div[1]/span[2]')
    no.click()

    # click 'paper' option
    WebDriverWait(IE, 10).until(EC.element_to_be_clickable((By.XPATH,'/html/body/div[3]/div/div/div/div/div/div/div/div/div/div[2]/div[2]/button')))
    IE.execute_script("document.querySelector('body > div:nth-child(4) > div > div > div > div > div > div > div > div > div > div.col-md-10 > div:nth-child(6) > button').click();")

    # click 'next'
    next = IE.find_element_by_xpath('/html/body/div[3]/div/footer/div/div/div[4]/div/button')
    next.click()

    # Next page
    WebDriverWait(IE, 30).until(EC.url_contains("https://retail-dev.metropolitan.co.za/funeral-planner/advisor"))
    time.sleep(2)

    # click both 'yes' options for financial advisor
    yes_1 = IE.find_element_by_css_selector('#section_50 > div:nth-child(2) > div:nth-child(1) > div > div > div.col-md-2 > span:nth-child(1)')
    yes_1.click()
    yes_2 = IE.find_element_by_css_selector('#section_50 > div:nth-child(3) > div:nth-child(1) > div > div > div.col-md-2 > span:nth-child(1)')
    yes_2.click()

    # click 'next'
    nex = IE.find_element_by_css_selector('body > div:nth-child(4) > div > footer > div > div > div:nth-child(4) > div > button')
    nex.click()

    # Next page
    WebDriverWait(IE, 20).until(EC.url_contains("https://retail-dev.metropolitan.co.za/funeral-planner/finalise-application"))
    time.sleep(2)

    # click 'upload' button
    upload_button = IE.find_element_by_css_selector('#uploadDocBtn-0')
    upload_button.click()
    time.sleep(1)
    pyautogui.click(x=227, y=148, clicks=1)
    time.sleep(1)
    pyautogui.click(x=782, y=507, clicks=1)
    # path = "C:\\Users\\EcBerry\\Desktop\\Myriad info\\Metropolitain\\VP\\Funeral_Ratebook_17254v1_112018_web (2).pdf"
    # print(pyautogui.position())

    # click 'submit application'
    WebDriverWait(IE, 80).until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'body > div:nth-child(4) > div > footer > div > div > div:nth-child(4) > div > button')))
    submit = IE.find_element_by_css_selector('body > div:nth-child(4) > div > footer > div > div > div:nth-child(4) > div > button')
    IE.execute_script("arguments[0].click();", submit)

    # Waiting for the 'close' button to become clickable
    WebDriverWait(IE, 300).until(EC.element_to_be_clickable((By.CSS_SELECTOR,'body > div.modal.fade.ng-scope.ng-isolate-scope.in > div > div > div > div > div > div > button')))
    close = IE.find_element_by_css_selector('body > div.modal.fade.ng-scope.ng-isolate-scope.in > div > div > div > div > div > div > button')
    close.click()
    time.sleep(2)
    IE.quit()
    time.sleep(4)

