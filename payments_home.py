
from selenium import webdriver
import xlgc_home
import email_data
from selenium.webdriver.common.keys import Keys
from msedge.selenium_tools import Edge, EdgeOptions
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.ie.options import Options
from fuzzywuzzy import fuzz
from tkinter import messagebox
import os, time, getpass, datetime
import xlwings
import glob, traceback
import PyPDF2, re
import pandas
from pathlib import Path
print(os.path.join(Path(__file__).parent))
class web_control:
    def __init__(self, browser):
        # os.environ["HTTP_PROXY"] = ""
        # os.environ["HTTPS_PROXY"] = ""
        # prefs ={"browser.startup.page": 1,
        #         "browser.startup.homepage": "http://www.seleniumhq.org"}

        if browser == 'edge':
            edge_driver_path = os.path.join(os.path.abspath(os.getcwd()), 'web_drivers', 'msedgedriver')
            self.Edge_Options = EdgeOptions()
            self.Edge_Options.add_experimental_option("excludeSwitches", ["enable-automation"])
            self.Edge_Options.add_experimental_option('useAutomationExtension', False)
            self.Edge_Options.use_chromium = True
            self.Edge_Options.add_argument("--start-maximized")
            self.Edge_Options.add_argument('no-sandbox')
            # self.Edge_Options.add_experimental_option('prefs',prefs)
            # self.Edge_Options.to_capabilities()
            self.Edge_Options.add_argument("user-data-dir=C:/Users/" + getpass.getuser() + "/AppData/Local/Microsoft/Edge/User Data")
            self.Edge_Options.add_argument("headless")
            self.Edge_Options.add_argument("--remote-debugging-port=9222")
            # self.Edge_Options.add_argument("disable-gpu")
            try:
                self.driver = Edge(edge_driver_path, options=self.Edge_Options)

            except:
                messagebox.showinfo('Bottomline Payments Tool',
                                    'Please close Edge and then run the tool.')
            self.driver.set_page_load_timeout(60)

        elif browser == 'chrome':
            chrome_driver_path = os.path.join(os.path.abspath(os.getcwd()), 'web_drivers', 'chromedriver')
            chromeOptions = webdriver.ChromeOptions()
            chromeOptions.add_argument("--start-maximized")
            # chromeOptions.add_argument("--remote-debugging-port=9222")
            # chromeOptions.add_argument("--headless")
            chromeOptions.add_argument("--disable-gpu")
            self.driver = webdriver.Chrome(chrome_driver_path, options=chromeOptions)

        elif browser == 'ie':
            ie_driver_path = os.path.join(os.path.abspath(os.getcwd()),'web_drivers','IEDriverServer')
            ieOptions = Options()
            ieOptions.add_argument("--start-maximized")
            self.driver = webdriver.Ie(ie_driver_path, options=ieOptions)

    def ElementFinder(self, element, action="", text=""):
        self.htmlElement = None
        self.waiting = WebDriverWait(self.driver, 30).until(EC.presence_of_element_located((By.XPATH, element)))
        self.wait = WebDriverWait(self.driver, 30).until(EC.visibility_of_element_located((By.XPATH, element)))
        self.htmlElement = self.driver.find_element_by_xpath(element)
        if action == "click":
            self.driver.execute_script("arguments[0].click();", self.htmlElement)
        if action == "sendKeys":
            self.htmlElement.clear()
            self.htmlElement.send_keys(text)

    def element_focus(self, element):
        self.waiting = WebDriverWait(self.driver, 30).until(EC.presence_of_element_located((By.XPATH, element)))
        hover = ActionChains(self.driver).move_to_element(self.driver.find_element_by_xpath(element))
        hover.perform()

class read_data:
    def __init__(self):
        # global variables declaration
        # documentum live
        # self.documentum_url = ''
        # documentum uat
        self.documentum_url = ''
        # xlgc copylive
        # self.xlgc_url = ''
        # xlgc live
        # self.xlgc_url = ''
        # xlgc uat
        self.xlgc_url = ''
        self.excel = xlwings.Book(os.path.join(os.path.abspath(os.getcwd()), 'Bottmoline Payments Tool.xlsm'))
        self.excel.save()
        self.details_sht = self.excel.sheets['Invoice details']
        self.tool_sht = self.excel.sheets['Tool']
        self.banking_sht = self.excel.sheets['Vendor Banking Details']
        self.expert_sht = self.excel.sheets['Expert List']
        self.last_row = self.details_sht.range('C' + str(self.details_sht.cells.last_cell.row)).end('up').row
        self.expert_last_row = self.expert_sht.range('A' + str(self.expert_sht.cells.last_cell.row)).end('up').row
        self.banking_last_row = self.banking_sht.range('A' + str(self.banking_sht.cells.last_cell.row)).end('up').row
        self.option_selected = int(self.tool_sht.shapes("Drop Down 3").api.ControlFormat.ListIndex)
        self.col = {'run_date': 'A', 'subject': 'B','claim': 'C','vendor': 'D','ttp': 'E',
               'inv': 'F','btt': 'G', 'claim_prof':'H', 'documentum':'I','ref': 'J','address': 'K',
                'bottomline_status':'L', 'URL':'M'}
        self.all_folder_sequence = []
        self.all_prefix = []
        self.lob_trigger = []
        for tool_row in range(8,11):
            self.all_folder_sequence.append([str(self.tool_sht.range(f'E{tool_row}').value), "Inbox"])
            self.all_prefix.append(str(self.tool_sht.range(f'F{tool_row}').value))
            self.lob_trigger.append(str(self.tool_sht.range(f'G{tool_row}').value))

        self.today = datetime.datetime.today()
        self.today_strf = self.today.strftime('%m/%d/%Y')
        self.today_strp = datetime.datetime.strptime(self.today_strf,'%m/%d/%Y')

class documentum:
    def __init__(self, driver, ElementFinder, expert_sht, details_sht, expert_last_row, col, today_strf, all_prefix):
        self.driver = driver
        self.ElementFinder = ElementFinder
        self.expert_sht = expert_sht
        self.details_sht = details_sht
        self.expert_last_row = expert_last_row
        self.col = col
        self.today_strf = today_strf
        self.all_prefix = all_prefix

    def open_documentum(self, documentum_url):
        star_config = pandas.read_csv('C:/XL_Apps/config.csv')
        login_id = star_config.iloc[0, 1]
        password = star_config.iloc[1, 1]
        self.driver.get(documentum_url)
        # login to documentum
        self.ElementFinder('//input[@id="j_username-inputEl"]','sendKeys', login_id)
        self.ElementFinder('//input[@id="j_password-inputEl"]','sendKeys', password)
        time.sleep(3)
        self.ElementFinder('//span[@id="signin-btn-btnInnerEl" and text()="Sign In"]', 'click')
        # driver.find_element_by_xpath('//span[@id="signin-btn-btnInnerEl" and text()="Sign In"]').click()
        # tasks > All tasks
        try:
            wait = WebDriverWait(self.driver, 60).until(EC.visibility_of_element_located((
                By.XPATH,'//div[contains(@class,"workqueue_dropdown_list")]/descendant::span[contains(text(),"Select Work Queue:")]')))
        except:
            self.ElementFinder('//span[@id="signin-btn-btnInnerEl" and text()="Sign In"]', 'click')
            wait = WebDriverWait(self.driver, 60).until(EC.visibility_of_element_located((
                By.XPATH,'//div[contains(@class,"workqueue_dropdown_list")]/descendant::span[contains(text(),"Select Work Queue:")]')))

        self.ElementFinder('//span[text()="Tasks"]', 'click')
        self.ElementFinder('//span[text()="All Tasks"]','click')
        self.documentum_table_load()
        # select work queue > payments queue
        self.ElementFinder('(//div[contains(@class,"task_list_box")]/descendant::div[contains(@class,"x-form-arrow-trigger-default")])[1]','click')
        self.ElementFinder('//ul[@class="x-list-plain"]/div[text()="payments queue"]', 'click')
        # select location queue > Americas
        self.ElementFinder('(//div[contains(@class,"task_list_box")]/descendant::div[contains(@class,"x-form-arrow-trigger-default")])[3]','click')
        self.ElementFinder('//ul[@class="x-list-plain"]/div[text()="Americas"]', 'click')
        self.documentum_table_load()

    def documentum_table_load(self, wait_time=30):
        try:
            wait = WebDriverWait(self.driver, wait_time).until(EC.visibility_of_element_located((By.XPATH,'//div[contains(@class,"x-masked")]')))
            doc_table = self.driver.find_element_by_xpath('//div[contains(@class,"x-masked")]')
            while True:
                if not 'x-masked' in doc_table.get_attribute('class'):
                    break
        except:
            pass

    def documentum_data_extraction(self, lob_trigger, outlook, pdf_data):

        # Extract document names from the page
        counter = self.details_sht.range('C' + str(self.details_sht.cells.last_cell.row)).end('up').row + 1
        lob_name_dict = {0: f'contains(text(),"{self.all_prefix[0]}")',
                         1: f'contains(text(),"{self.all_prefix[1]}")',
                         2: f'contains(text(),"{self.all_prefix[2]}")'}
        prefix_xpath = [f'{lob_name_dict[trigger]}' for trigger in range(len(lob_trigger)) if lob_trigger[trigger] == 'Yes']
        prefix_xpath = ' or '.join(prefix_xpath)
        while True:
            self.documentum_table_load(10)
            all_documents = self.driver.find_elements_by_xpath(
                f'//div[@class="x-grid-item-container"]/table[./descendant::a[contains(text(),"Invoice#") and ({prefix_xpath})]]')
            if all_documents.__len__() > 0:
                try:
                    all_document_names = [document.find_element_by_xpath(
                                        f'./descendant::a[contains(text(),"Invoice#") and ({prefix_xpath})]').get_attribute('innerText')
                                          for document in all_documents]
                except:
                    self.driver.refresh()
                    self.documentum_table_load(10)
                    all_documents = self.driver.find_elements_by_xpath(
                        f'//div[@class="x-grid-item-container"]/table[./descendant::a[contains(text(),"Invoice#") and ({prefix_xpath})]]')
                    all_document_names = [document.find_element_by_xpath(
                        f'./descendant::a[contains(text(),"Invoice#") and ({prefix_xpath})]').get_attribute('innerText')
                                          for document in all_documents]

                for document_name in all_document_names:
                    self.invoice_num = ''
                    self.invoice_num = re.search(r'(Invoice#.*)(,Client)', document_name)
                    self.prefix = (self.all_prefix[0] if self.all_prefix[0] in document_name else
                                   (self.all_prefix[1] if self.all_prefix[1] in document_name else
                                   (self.all_prefix[2] if self.all_prefix[2] in document_name else None)))
                    if self.invoice_num:
                        self.invoice_num = self.invoice_num.group(1).strip().replace('Invoice# ','Invoice#:')
                        self.details_sht.range('{0}{1}'.format(self.col['subject'], counter)).value = document_name
                        # extract data from mail and its attachment.
                        outlook.get_email_data(self.invoice_num, self.prefix,  counter)
                        pdf_data.pdf_reader(counter)
                        counter += 1
            # Go to next page
            if self.driver.find_element_by_xpath(
                    '(//a[./descendant::span[contains(@class,"page-next")]])[2]').get_attribute(
                    'aria-disabled') == 'false':
                self.ElementFinder('(//a[./descendant::span[contains(@class,"page-next")]])[2]', 'click')
                self.documentum_table_load()
            else:
                break


    def documentum_data_entry(self, row_values, row):
        # run for line items in excel file.
        if row_values['claim'] != 'None' \
        and re.search(r'Match' ,row_values['documentum_status']) != None:

            invoice_num = re.search(r'(Invoice#.*)(,Client)', row_values['subject'])
            invoice_num = invoice_num.group(1)
            prefix = (self.all_prefix[0] if self.all_prefix[0] in row_values['subject'] else
                       (self.all_prefix[1] if self.all_prefix[1] in row_values['subject'] else
                        (self.all_prefix[2] if self.all_prefix[2] in row_values['subject'] else None)))
            # go to first page

            while True:
                if self.driver.find_element_by_xpath('(//a[./descendant::span[contains(@class,"page-prev")]])[2]').get_attribute('aria-disabled') == 'false':
                    self.ElementFinder('(//a[./descendant::span[contains(@class,"page-prev")]])[2]', 'click')
                    self.documentum_table_load()
                else:
                    break
            # find document name in all pages
            while True:
                self.documentum_table_load(10)
                all_documents = self.driver.find_elements_by_xpath(
                f'//div[@class="x-grid-item-container"]/table[./descendant::a[contains(text(),"{invoice_num}") and contains(text(),"{prefix}")]]')
                if all_documents.__len__() == 0:
                    if self.driver.find_element_by_xpath('(//a[./descendant::span[contains(@class,"page-next")]])[2]').get_attribute('aria-disabled') == 'false':
                        self.ElementFinder('(//a[./descendant::span[contains(@class,"page-next")]])[2]', 'click')
                        self.documentum_table_load()
                    else:
                        break
                else:
                    break

            if all_documents.__len__() >= 1:
                element = all_documents[0].find_element_by_xpath('./descendant::td[2]')
                self.driver.execute_script("arguments[0].scrollIntoView(true);", element)
                time.sleep(5)
                element = self.driver.find_element_by_xpath(f'//div[@class="x-grid-item-container"]/table[./descendant::a[contains(text(),"{invoice_num}") and contains(text(),"{prefix}")]]/descendant::td[2]')
                action = ActionChains(self.driver)
                action.move_to_element(element)
                action.context_click(element).perform()
                # click on acquire
                wait = WebDriverWait(self.driver, 5).until(EC.visibility_of_element_located((
                    By.XPATH,'//div[contains(@class,"context_menu_xcp_acquire_task")]/a')))
                self.driver.find_element_by_xpath('//div[contains(@class,"context_menu_xcp_acquire_task")]/a').click()

                expert_type = 'Adjuster'# driver.find_element_by_xpath('//div[contains(@class,"context_menu_task_gotopage")]/a').click()
                # action.context_click(all_documents[0]).send_keys(Keys.ARROW_DOWN).send_keys(Keys.ARROW_DOWN).send_keys(Keys.RETURN).perform()
                # action.context_click(all_documents[0]).send_keys(Keys.ARROW_DOWN).send_keys(Keys.RETURN).perform()
                # check if vendor is expert or adjuster
                for exp_row in range(2, self.expert_last_row + 1):
                    if fuzz.ratio(row_values['vendor'].lower(), str(self.expert_sht.range(f'A{exp_row}').value).lower()) > 70:
                    # if row_values['vendor'].lower().split()[0] in str(self.expert_sht.range(f'A{exp_row}').value).lower():
                        expert_type = 'Expert'
                        break
                # upload data to documentum
                title = '{0} - Payment of USD {1} {2} to {3} USD {4} Adjuster to Bottomline'.format(
                    self.today_strf, row_values['ttp'], expert_type, row_values['vendor'], row_values['btt'])
                if str(row_values['claim'])[:4] != '000':
                    claim = '000' + str(row_values['claim'])
                wait = WebDriverWait(self.driver, 30).until(EC.visibility_of_element_located((
                    By.XPATH, f'//div[contains(@id,"value_display") and contains(text(),"{invoice_num}")]')))
                print('acquired')
                time.sleep(5)
                self.driver.execute_script("arguments[0].scrollIntoView(true);", self.driver.find_element_by_xpath(
                    '//div[contains(@class,"GQ_ACTION_DROPDOWN_LIST")]/descendant::div[contains(@class,"x-form-arrow-trigger")]'))
                self.ElementFinder('//div[contains(@class,"GQ_ACTION_DROPDOWN_LIST")]/descendant::div[contains(@class,"x-form-arrow-trigger")]','click')
                self.ElementFinder('//div[text()="Index and do not create a task in XLGC"]','click')
                self.ElementFinder('//div[contains(@class,"CLAIM_TEXT_INPUT")]/descendant::input','sendKeys', claim)
                self.ElementFinder('//div[contains(@class,"CATEGORY_DROPDOWN_LIST")]/descendant::div[contains(@class,"x-form-arrow-trigger")]','click')
                self.ElementFinder('//div[contains(@class,"x-boundlist-item") and text()="Financials"]','click')
                try:
                    self.ElementFinder('//div[contains(@class,"SUBCAT_DROPDOWN_LIST")]/descendant::div[contains(@class,"x-form-arrow-trigger")]','click')
                    self.ElementFinder('//div[contains(@class,"x-boundlist-item") and text()="Payments"]','click')
                except:
                    self.ElementFinder('//div[contains(@class,"CATEGORY_DROPDOWN_LIST")]/descendant::div[contains(@class,"x-form-arrow-trigger")]','click')
                    self.ElementFinder('//div[contains(@class,"x-boundlist-item") and text()="Financials"]', 'click')
                    self.ElementFinder('//div[contains(@class,"SUBCAT_DROPDOWN_LIST")]/descendant::div[contains(@class,"x-form-arrow-trigger")]','click')
                    self.ElementFinder('//div[contains(@class,"x-boundlist-item") and text()="Payments"]', 'click')
                self.ElementFinder('//div[contains(@class,"TITLE_TEXT_INPUT")]/descendant::input','sendKeys', title)
                # click on done or close button
                try:
                    wait = WebDriverWait(self.driver, 5).until(EC.visibility_of_element_located((By.XPATH, '//a[@aria-disabled="false" and contains(@class,"done_button")]')))
                except:
                    pass

                # self.ElementFinder('//span[text()="Close" and contains(@class,"x-btn-inner")]', 'click')
                # self.details_sht.range('{0}{1}'.format(self.col['documentum'], row)).value = "Done"
                # print('Done')
                if self.driver.find_element_by_xpath('//a[contains(@class,"done_button")]').get_attribute('aria-disabled') == 'false':
                    self.ElementFinder('//span[text()="Done" and contains(@class,"x-btn-inner")]', 'click')
                    self.details_sht.range('{0}{1}'.format(self.col['documentum'], row)).value = "Done"
                else:
                    self.details_sht.range('{0}{1}'.format(self.col['documentum'], row)).value = 'Acquired but could not complete.'
                    self.ElementFinder('//span[text()="Close" and contains(@class,"x-btn-inner")]', 'click')
                self.documentum_table_load()
            else:
                self.details_sht.range('{0}{1}'.format(self.col['documentum'], row)).value = "Error"

class pdf_extractor:

    def __init__(self, details_sht, col):
        self.details_sht = details_sht
        self.col = col

    def pdf_reader(self, row):

        if self.details_sht.range('{0}{1}'.format(self.col['claim'],row)).value != None:
            claim = int(float(self.details_sht.range('{0}{1}'.format(self.col['claim'],row)).value))
            vendor_pdf_name = os.path.join(os.path.abspath(os.getcwd()), 'Invoice', 'sfi*{0}*.pdf'.format(claim))
            bt_pdf_name = os.path.join(os.path.abspath(os.getcwd()), 'Invoice', 'bar*{0}*.pdf'.format(claim))

            # extract Reference number from bottomline invoice
            pdf_list = glob.glob(vendor_pdf_name)
            if pdf_list.__len__() > 0:
                pdfFileObj = open(pdf_list[0], 'rb')
                pdfReader = PyPDF2.PdfFileReader(pdfFileObj)
                pageObj = pdfReader.getPage(0)
                reference = re.search(r'\d{13,15}',pageObj.extractText())
                if reference:
                    self.details_sht.range('{0}{1}'.format(self.col['ref'],row)).value = reference.group()
                pdfFileObj.close()

            # extract address from vendor invoice
            pdf_list = glob.glob(bt_pdf_name)
            if pdf_list.__len__() > 0:
                pdfFileObj = open(pdf_list[0], 'rb')
                pdfReader = PyPDF2.PdfFileReader(pdfFileObj)
                pageObj = pdfReader.getPage(0)
                firm_address = re.search(r'(Firm Address:)((.|\n)*)(Firm Phone:)',pageObj.extractText())
                if firm_address:
                    self.details_sht.range('{0}{1}'.format(self.col['address'],row)).value = firm_address.group(2)
                pdfFileObj.close()

def initialize():

    data = read_data()

    web = web_control('chrome')

    if data.option_selected == 1 or data.option_selected == 2:
        documentum_object = documentum(web.driver, web.ElementFinder, data.expert_sht,
                              data.details_sht, data.expert_last_row, data.col, data.today_strf, data.all_prefix)
        try:
            documentum_object.open_documentum(data.documentum_url)
        except:
            web.driver.quit()
            messagebox.showerror('Bottomline payments tool',
                                'Cannot sign in to documentum, please run the tool again.')
            return

    if data.option_selected == 1:
        outlook = email_data.email_extract(data.all_folder_sequence, data.all_prefix, data.lob_trigger,
                                           data.details_sht, data.col, data.today_strf)
        pdf_data = pdf_extractor(data.details_sht, data.col)
        documentum_object.documentum_data_extraction(data.lob_trigger, outlook, pdf_data)
        data.excel.save()
        data.last_row = data.details_sht.range('C' + str(data.details_sht.cells.last_cell.row)).end('up').row

    if data.option_selected == 1 or data.option_selected == 3:
        xlgc_object = xlgc_home.xlgc(web.driver, web.ElementFinder, web.element_focus, data.details_sht, data.col, data.today_strp,
                     data.today_strf, data.banking_sht, data.banking_last_row, data.expert_sht, data.expert_last_row)
        xlgc_object.open_xlgc(data.xlgc_url)

    for row in range(2, data.last_row + 1):
        try:
            row_values = {'subject': data.details_sht.range('{0}{1}'.format(data.col['subject'], row)).value,
                          'claim': int(float(data.details_sht.range('{0}{1}'.format(data.col['claim'], row)).value)),
                          'vendor': str(data.details_sht.range('{0}{1}'.format(data.col['vendor'], row)).value),
                          'ttp': float(data.details_sht.range('{0}{1}'.format(data.col['ttp'], row)).value),
                          'inv': data.details_sht.range('{0}{1}'.format(data.col['inv'], row)).value,
                          'btt': float(data.details_sht.range('{0}{1}'.format(data.col['btt'], row)).value),
                          'claim_prof': str(data.details_sht.range('{0}{1}'.format(data.col['claim_prof'], row)).value),
                          'ref': data.details_sht.range('{0}{1}'.format(data.col['ref'], row)).value,
                          'address': str(data.details_sht.range('{0}{1}'.format(data.col['address'], row)).value),
                          'documentum_status': str(data.details_sht.range('{0}{1}'.format(data.col['documentum'], row)).value),
                          'bottomline_status': str(data.details_sht.range('{0}{1}'.format(data.col['bottomline_status'], row)).value)}
        except:
            data.details_sht.range('{0}{1}'.format(data.col['bottomline_status'], row)).value = 'Missing Data'
            data.details_sht.range('{0}{1}'.format(data.col['documentum'], row)).value = 'Missing Data'
            continue

        if data.option_selected == 2:
            try:
                documentum_object.documentum_data_entry(row_values=row_values, row=row)
            except:
                try:
                    data.details_sht.range('{0}{1}'.format(data.col['documentum'], row)).value = 'Error'
                    web.driver.get(data.documentum_url)
                    web.ElementFinder('//div[contains(@class,"workqueue_dropdown_list")]/descendant::span[contains(text(),"Select Work Queue:")]')
                    documentum_object.documentum_table_load()
                except:
                    messagebox.showerror('Bottomline payments tool', traceback.format_exc())
                    return

        if data.option_selected == 3:
            try:
                xlgc_object.xlgc_data_entry(row_values, row)
            except:
                try:
                    web.driver.delete_all_cookies()
                    data.details_sht.range('{0}{1}'.format(data.col['bottomline_status'], row)).value = 'Error'
                    xlgc_object.xlgc_restart()
                    xlgc_object.open_xlgc(data.xlgc_url)
                except:
                    messagebox.showerror('Bottomline payments tool', traceback.format_exc())
                    return
        if data.option_selected == 1:
            try:
                xlgc_object.claim_professional(row_values, row)
            except:
                try:
                    web.driver.delete_all_cookies()
                    data.details_sht.range('{0}{1}'.format(data.col['documentum'], row)).value = 'Error'
                    xlgc_object.xlgc_restart()
                    xlgc_object.open_xlgc(data.xlgc_url)
                except:
                    messagebox.showerror('Bottomline payments tool', traceback.format_exc())
                    return

    web.driver.quit()
    messagebox.showinfo('Bottomline payments tool', 'Complete.')

if __name__ == '__main__':
    initialize()

