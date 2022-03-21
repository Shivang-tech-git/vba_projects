from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from datetime import datetime
import time, re
from fuzzywuzzy import fuzz
class frames_workflow:

    def __init__(self, driver, event, main_df, info, element_finder):
        self.driver = driver
        self.event = event
        self.main_df = main_df
        self.info = info
        self.element_finder = element_finder
        self.property_services = ''

    def event_refresh(self, ue_df):
        self.ue_df = ue_df
        # --------------- navigate to outstanding movements -------------------
        self.element_finder('//td[@class="mainmenu" and text()="Claims"]', 'click')
        self.element_finder('//div[@id="claims.outstandingMovements"]/a[text()="Outstanding movements"]', 'click')
        # --------------- For each row in pcs_df dataframe --------------------
        for i, row in self.ue_df.iterrows():
            self.coding_process(df=self.ue_df, i=i, row=row)
            # -------------- update pcs_df with new values from ue_pcs_df --------------
            self.main_df.loc[self.main_df['BPR'] == self.ue_df.at[i, 'BPR'], 'Status'] = self.ue_df.at[i, 'Status']
            self.main_df.loc[self.main_df['BPR'] == self.ue_df.at[i, 'BPR'], 'Event'] = self.ue_df.at[i, 'Event']
            self.main_df.loc[self.main_df['BPR'] == self.ue_df.at[i, 'BPR'], 'Loss type'] = self.ue_df.at[i, 'Loss type']
            self.main_df.to_csv('{0} Output {1}.CSV'.format(self.event, datetime.today().strftime('%d-%m-%Y')), index=False)
        self.driver.quit()

    def event_coding(self, blank_df):
        self.blank_df = blank_df
        # --------------- navigate to outstanding movements -------------------
        self.element_finder('//td[@class="mainmenu" and text()="Claims"]','click')
        self.element_finder('//div[@id="claims.outstandingMovements"]/a[text()="Outstanding movements"]','click')
        # --------------- For each row in main_df, find BPR --------------------
        for i, row in blank_df.iterrows():
            self.coding_process(df=self.blank_df, i=i, row=row)
            self.main_df.loc[self.main_df['BPR'] == self.blank_df.at[i, 'BPR'], 'Status'] = self.blank_df.at[i, 'Status']
            self.main_df.loc[self.main_df['BPR'] == self.blank_df.at[i, 'BPR'], 'Event'] = self.blank_df.at[i, 'Event']
            self.main_df.loc[self.main_df['BPR'] == self.blank_df.at[i, 'BPR'], 'Loss type'] = self.blank_df.at[i, 'Loss type']
            self.main_df.to_csv('{0} Output {1}.CSV'.format(self.event, datetime.today().strftime('%d-%m-%Y')), index=False)
        self.err_df = self.main_df.loc[self.main_df['Status'] == 'Undefined Error']
        # -------------- run again for undefined error -----------------
        if self.err_df.shape[0] > 0:
            for i, row in self.err_df.iterrows():
                # -------------- update main_df with new values from err_df --------------
                self.coding_process(df=self.err_df, i=i, row=row)
                self.main_df.loc[self.main_df['BPR'] == self.err_df.at[i, 'BPR'], 'Status'] = self.err_df.at[i, 'Status']
                self.main_df.loc[self.main_df['BPR'] == self.err_df.at[i, 'BPR'], 'Event'] = self.err_df.at[i, 'Event']
                self.main_df.loc[self.main_df['BPR'] == self.err_df.at[i, 'BPR'], 'Loss type'] = self.err_df.at[i, 'Loss type']
                self.main_df.to_csv('{0} Output {1}.CSV'.format(self.event, datetime.today().strftime('%d-%m-%Y')), index=False)
        self.driver.quit()

    def coding_process(self, df, i, row):
        try:
            if row['BPR'] != '':
                try:
                    BPR = int(float(row['BPR']))
                except:
                    BPR = row['BPR']
                self.element_finder('//input[@type = "text" and @name = "bpr"]', 'sendKeys', BPR)
                self.element_finder('//input[@type = "submit" and @value = "Search"]', 'click')
                assign_event_exist = False
                search_code = []
                # ----------------------- check for the bpr text and cat code text ----------------------------------
                time.sleep(5)
                self.element_finder('//td[./input[@id="csrfToken"]]')
                BPR_text = self.driver.find_element_by_xpath('//td[./input[@id="csrfToken"]]').get_attribute('innerText')
                if BPR_text.find('There are no items for event assignment') > 0:
                    df.at[i, 'Status'] = 'There are no items for event assignment'
                    return
                else:
                    assign_event_exist = True
                # ------------------- click on assign event-------------------------------
                if assign_event_exist == True:
                    self.element_finder('//table[@id="viewBean"]/descendant::input[@value="Assign event"]', 'click')
                    waiting = WebDriverWait(self.driver, 20).until(
                        EC.presence_of_element_located((By.XPATH, '//table[@class="popupTable"]')))
                    # --------------- CAT ---------------------------------------
                    if self.event == 'CAT':
                        search_code = ['','','','','',row['Advised cat.']]
                        df.at[i, 'Event'] = 'CAT'
                    # --------------- PCS ---------------------------------------
                    elif self.event == 'PCS':
                        self.pcs_event_code_process()
                        if (self.property_services == '') and (df.at[i, 'Description'] == 'Mandatory Review. Event Code VARS'):
                            search_code = ['vars', '', '', '', '', '']
                            df.at[i, 'Event'] = 'VARS'

                        elif self.property_services == '':
                            actuals_process_result = self.actuals_event_code_process(df,i,row)
                            if actuals_process_result == 'missing info':
                                search_code = ['vars', '', '', '', '', '']
                                df.at[i, 'Event'] = 'VARS'
                            elif actuals_process_result == 'manual intervention':
                                df.at[i, 'Status'] = 'Manual Intervention'
                                return
                            else:
                                df.at[i, 'Event'] = 'Actuals'
                                return

                        else:
                            df.at[i, 'Event'] = 'PCS'
                            search_code = ['', 'PCS {}'.format(self.property_services), '', '', '', '']
                    # --------------- Actuals ---------------------------------------
                    elif self.event == 'Actuals':
                        self.loss_type = ''
                        actuals_process_result = self.actuals_event_code_process(df, i, row)
                        if actuals_process_result == 'missing info':
                            search_code = ['vars', '', '', '', '', '']
                            df.at[i, 'Event'] = 'VARS'
                        elif actuals_process_result == 'manual intervention':
                            df.at[i, 'Status'] = 'Manual Intervention'
                            return
                        else:
                            df.at[i, 'Event'] = 'Actuals'
                            df.at[i, 'Loss type'] = self.loss_type
                            return
                    # --------------- VARS ---------------------------------------
                    elif self.event == 'VARS':
                        df.at[i, 'Event'] = 'VARS'
                        search_code = ['vars', '', '', '', '', '']
                    # ------------------------------ search for specific event code results ----------------------------------------
                    self.element_finder('//table[@class="popupTable"]/descendant::input[@value=" ... "]', 'click')
                    self.element_finder('//div[@class="popupBody"]/descendant::input[@name="eventCode"]', 'sendKeys', search_code[0])
                    self.element_finder('//div[@class="popupBody"]/descendant::input[@name="eventName"]', 'sendKeys', search_code[1])
                    self.element_finder('//div[@class="popupBody"]/descendant::input[@name="startDate"]', 'sendKeys', search_code[2])
                    self.element_finder('//div[@class="popupBody"]/descendant::input[@name="narrative"]', 'sendKeys', search_code[3])
                    self.element_finder('//div[@class="popupBody"]/descendant::input[@name="days"]', 'sendKeys', search_code[4])
                    self.element_finder('//div[@class="popupBody"]/descendant::input[@name="catCode"]', 'sendKeys', search_code[5])
                    # ------------------------------ cases > no results found or multiple cat codes or one cat code ----------------
                    self.element_finder(
                        '//table[@class="popupTable"]/descendant::input[@name="linkPressed"and @value="Search"]', 'click')

                    waiting = WebDriverWait(self.driver, 20).until(EC.presence_of_element_located(
                        (By.XPATH, '//div[@class="popupBody"]/descendant::td[@class="tabledata"]')))
                    if self.driver.find_element_by_xpath(
                            '//div[@class="popupBody"]/descendant::td[@class="tabledata"]').get_attribute(
                            'innerText') == 'No results found':
                        df.at[i, 'Status'] = 'No results found'
                        self.element_finder(
                            '(//table[@class="popupTable"]/descendant::input[@value="Cancel" and @type="submit"])[2]',
                            'click')
                        self.element_finder(
                            '(//table[@class="popupTable"]/descendant::input[@value="Cancel" and @type="submit"])[1]',
                            'click')
                    else:
                        try:
                            html_element = self.driver.find_element_by_xpath(
                                '(//div[@class="popupBody"]/descendant::table[@class="tablerule"]/descendant::tr[./td[@class="tabledata"]])[2]')
                            df.at[i, 'Status'] = 'Multiple events exist'
                            self.element_finder(
                                '(//table[@class="popupTable"]/descendant::input[@value="Cancel" and @type="submit"])[2]',
                                'click')
                            self.element_finder(
                                '(//table[@class="popupTable"]/descendant::input[@value="Cancel" and @type="submit"])[1]',
                                'click')
                        except:
                            self.element_finder('//table[@class="popupTable"]/descendant::input[@value="   OK   "]', 'click')
                            time.sleep(5)
                            self.element_finder(
                                ('//table[@class="popupTable"]/descendant::input[@value="Done" and @type="submit"]'),
                                'click')
                            df.at[i, 'Status'] = 'Successfully Allocated'

        #  if any row throws unexpected error, then start frames again and start from the next row.
        except:
            df.at[i, 'Status'] = 'Undefined Error'
            try:
                self.driver.get('')
                self.element_finder('//td[@class="mainmenu" and text()="Claims"]', 'click')
                self.element_finder('//div[@id="claims.outstandingMovements"]/a[text()="Outstanding movements"]', 'click')
            except:
                # self.driver.quit()
                self.driver.get('')
                self.element_finder('//td[@class="mainmenu" and text()="Claims"]', 'click')
                self.element_finder('//div[@id="claims.outstandingMovements"]/a[text()="Outstanding movements"]', 'click')

    def pcs_event_code_process(self):
        self.element_finder(
            '//table[@class="popupTable"]/descendant::td[./preceding-sibling::td[contains(text(),"Claim ID :")]]/a',
            'click')
        self.driver.switch_to.window(self.driver.window_handles[1])
        self.property_services = ''
        try:
            self.element_finder('//a[text()="Original bureau message"]', 'click')
            self.element_finder('//td[text()[contains(.,"Property services: ")]]')
            if self.driver.find_element_by_xpath(
                    '//td[text()[contains(.,"Property services: ")]]/span[6]').get_attribute('innerText') != '':
                self.property_services = self.driver.find_element_by_xpath(
                    '//td[text()[contains(.,"Property services: ")]]/span[6]').get_attribute('innerText')
        except:
            pass
        self.driver.close()
        self.driver.switch_to.window(self.driver.window_handles[0])

    def actuals_event_code_process(self, df, i, row):

        cost_center = str(row['CostCentre'])
        # loss_type = self.driver.find_element_by_xpath('//table[@class="popupTable"]/descendant::td[./preceding-sibling::td[contains(text(), "Loss type:")]][1]'
        #                                               ).get_attribute('innerHTML')
        claim_narrative = self.driver.find_element_by_xpath(
            '//table[@class="popupTable"]/descendant::td[./preceding-sibling::td[contains(text(), "Claim narrative :")]][1]'
                                                        ).get_attribute('innerHTML')
        loss_start_date = self.driver.find_element_by_xpath(
            '//table[@class="popupTable"]/descendant::td[./preceding-sibling::td[contains(text(), "Loss start date :")]][1]'
                                                        ).get_attribute('innerText')
        claim_start_date = self.driver.find_element_by_xpath(
            '//table[@class="popupTable"]/descendant::td[./preceding-sibling::td[contains(text(), "Claim start date :")]][1]'
                                                        ).get_attribute('innerText')
        advised_insured = self.driver.find_element_by_xpath(
            '//table[@class="popupTable"]/descendant::td[./preceding-sibling::td[contains(text(), "Advised insured :")]][1]'
                                                        ).get_attribute('innerText')

        if (loss_start_date == '' and claim_start_date == '') or claim_narrative == '':
            return 'missing info'

        for keyword in self.info.manual_keywords:
            if re.search(r"\b{}\b".format(keyword), claim_narrative.replace('\n',' ')):
                self.element_finder(
                    ('//table[@class="popupTable"]/descendant::input[@value="Cancel" and @type="submit"]'), 'click')
                return 'manual intervention'

        insured_name = ''
        # bloodstock
        if self.event == 'Actuals' and cost_center in ['EK','NX','UE']:
            if re.search('[^0-9](20\d{2} (.*?)\n)', claim_narrative):
                insured_name = re.search('[^0-9](20\d{2} (.*?)\n)', claim_narrative).group(1)
                insured_name = insured_name[:35]
        # hull
        elif self.event == 'Actuals' and cost_center in ['BH', 'TE', 'TX', 'TY']:
            if re.search('((VSL|VESSEL)(.*?))\\n', claim_narrative):
                insured_name = re.search('((VSL|VESSEL)(.*?))\\n', claim_narrative).group(1)

        if insured_name == '':
            if re.search('O/I (.*?)\\n',claim_narrative):
                insured_name = re.search('O/I (.*?)\\n',claim_narrative).group(1)
            elif re.search('INS (.*?)\\n',claim_narrative):
                insured_name = re.search('INS (.*?)\\n',claim_narrative).group(1)
            elif re.search('R/I (.*?)\\n',claim_narrative):
                insured_name = re.search('R/I (.*?)\\n',claim_narrative).group(1)
            elif advised_insured != '':
                insured_name = advised_insured
            else:
                return 'missing info'

        if loss_start_date != '':
            start_date = datetime.strptime(loss_start_date, '%d/%m/%Y')
        elif claim_start_date != '':
            start_date = datetime.strptime(claim_start_date, '%d/%m/%Y')

        insured_found = False
        self.element_finder('//table[@class="popupTable"]/descendant::input[@value=" ... "]', 'click')
        self.element_finder('//div[@class="popupBody"]/descendant::input[@name="searchData.maxNumberOfRows"]', 'sendKeys', '5099')
        self.element_finder('//table[@class="popupTable"]/descendant::input[@name="linkPressed"and @value="Search"]', 'click')
        waiting = WebDriverWait(self.driver, 20).until(EC.presence_of_element_located(
            (By.XPATH, '//div[@class="popupBody"]/descendant::td[@class="tabledata"]')))

        ins_rows = self.driver.find_elements_by_xpath(
            '//div[@class="popupBody"]/descendant::table[@class="tablerule"]/descendant::tr[./td[@class="tabledata"]]')
        if ins_rows[0].find_element_by_xpath('./td').get_attribute('innerText') != 'No results found':
            for ins_row in ins_rows:
                date_row_value = ins_row.find_element_by_xpath('./td[3]').get_attribute('innerText')
                date_row_value = datetime.strptime(date_row_value, '%d/%m/%y')
                ins_row_value = ins_row.find_element_by_xpath('./td[4]').get_attribute('innerText')
                lob_row_value = ins_row.find_element_by_xpath('./td[5]').get_attribute('innerText')
                if fuzz.ratio(ins_row_value.lower(), insured_name.lower()) > 80 and \
                        (lob_row_value.lower() == self.info.cost_center_to_lob[cost_center].lower()) and \
                        ((start_date - date_row_value).days == 0):
                    ins_row.find_element_by_xpath('./td[1]/input').click()
                    print(f'Existing insured name: {ins_row_value} Searched insured name: {insured_name}')
                    self.element_finder('//table[@class="popupTable"]/descendant::input[@value="   OK   "]', 'click')
                    time.sleep(5)
                    self.loss_type = self.driver.find_element_by_xpath('//table[@class="popupTable"]/descendant::td[./preceding-sibling::td[contains(text(), "Loss type:")]][1]/span'
                                                                  ).get_attribute('innerHTML')
                    self.element_finder(
                        ('//table[@class="popupTable"]/descendant::input[@value="Done" and @type="submit"]'),'click')
                    df.at[i, 'Status'] = f'Successfully Allocated to existing event, ' \
                                         f'Existing: {ins_row_value}, ' \
                                         f'Searched: {insured_name}'
                    insured_found = True
                    break

        if insured_found == False:

            if self.event == 'Actuals' and cost_center in ['EK', 'NX', 'UE', 'BH', 'TE', 'TX', 'TY']:
                self.element_finder('//table[@class="popupTable"]/descendant::input[@value="Cancel"]', 'click')
                df.at[i, 'Status'] = f'Allocate manually. No existing event found for ' \
                                     f'Insured name: {insured_name}'
                return

            else:
                self.element_finder('//table[@class="popupTable"]/descendant::input[@value="New..."]', 'click')
                waiting = WebDriverWait(self.driver, 20).until(
                            EC.presence_of_element_located((By.XPATH, '//table[@class="popupTable"]/descendant::td[text()="Create new event"]')))
                self.element_finder('//div[@class="popupBody"]/descendant::input[@name="eventName"]', 'sendKeys',
                                    insured_name)
                if self.driver.find_element_by_xpath('//div[@class="popupBody"]/descendant::input[@name="catNotApplicable"]').get_attribute('checked') != 'true':
                    self.element_finder('//div[@class="popupBody"]/descendant::input[@name="catNotApplicable"]', 'click')
                    self.element_finder('//div[@class="popupBody"]/descendant::input[@value="Not Applicable"]')
                loss_dropdown = Select(self.driver.find_element_by_xpath(
                    '//div[@class="popupBody"]/descendant::select[@name="defaultLossType"]'))
                try:
                    loss_dropdown.select_by_visible_text(self.info.cost_center_to_lob[cost_center])
                except:
                    df.at[i, 'Status'] = f'Cost center not found in the provided list.'
                    self.element_finder('//table[@class="popupTable"]/descendant::input[@value="Cancel"]', 'click')
                    return
                self.element_finder('//div[@class="popupFooter"]/descendant::input[@value=" OK "]', 'click')
                try:
                    waiting = WebDriverWait(self.driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, '//table[@class="popupTable"]/descendant::td[text()="Possible duplicates"]')))
                except:
                    self.loss_type = self.driver.find_element_by_xpath(
                        '//table[@class="popupTable"]/descendant::td[./preceding-sibling::td[contains(text(), "Loss type:")]][1]/span').get_attribute('innerHTML')
                    self.element_finder(('//table[@class="popupTable"]/descendant::input[@value="Done" and @type="submit"]'), 'click')
                    df.at[i, 'Status'] = f'Successfully Allocated to new event, ' \
                                         f'Insured name: {insured_name}'
                    return
                # duplicate_rows = self.driver.find_elements_by_xpath(
                # '//table[@class="popupTable" and ./descendant::td[text()="Possible duplicates"]]/descendant::table[@class="tablerule"]/descendant::tr[./td[@class="tabledata"]]')
                # for duplicate_row in duplicate_rows:
                #     duplicate_row_value = duplicate_row.find_element_by_xpath('./td[4]').get_attribute('innerText')
                #     dup_date_row_value = duplicate_row.find_element_by_xpath('./td[3]').get_attribute('innerText')
                #     dup_date_row_value = datetime.strptime(dup_date_row_value, '%Y/%m/%d %H:%M:%S')
                #     if fuzz.ratio(duplicate_row_value.lower(), insured_name.lower()) > 80 and \
                #         ((start_date - dup_date_row_value).days == 0):
                #         df.at[i, 'Status'] = f'Need to check Possible duplicate found, ' \
                #                              f'Existing: {duplicate_row_value}, ' \
                #                              f'Searched: {insured_name}'
                #         self.element_finder('//table[@class="popupTable"]/descendant::input[@value="Cancel"]', 'click')
                #         return

                self.element_finder('//table[@class="popupTable"]/descendant::input[@value="Ignore"]', 'click')
                time.sleep(5)
                self.loss_type = self.driver.find_element_by_xpath('//table[@class="popupTable"]/descendant::td[./preceding-sibling::td[contains(text(), "Loss type:")]][1]/span'
                    ).get_attribute('innerHTML')
                self.element_finder(
                    ('//table[@class="popupTable"]/descendant::input[@value="Done" and @type="submit"]'), 'click')
                df.at[i, 'Status'] = f'Successfully Allocated to new event, ' \
                                     f'Insured name: {insured_name}'


