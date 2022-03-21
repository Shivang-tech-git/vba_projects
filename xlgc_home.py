import math

from selenium.webdriver.common.action_chains import ActionChains
from fuzzywuzzy import fuzz
import datetime, re, time, os
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By

class xlgc:
    def __init__(self, driver, ElementFinder, element_focus, details_sht, col, today_strp, today_strf, banking_sht, banking_last_row,
                 expert_sht, expert_last_row):
        self.driver = driver
        self.ElementFinder = ElementFinder
        self.element_focus = element_focus
        self.details_sht = details_sht
        self.col = col
        self.today_strp = today_strp
        self.today_strf = today_strf
        self.banking_sht = banking_sht
        self.banking_last_row = banking_last_row
        self.expert_sht = expert_sht
        self.expert_last_row = expert_last_row
        diff = max(1, (datetime.datetime.today().weekday() + 6) % 7 - 3)
        self.previous_wd_strf = (datetime.datetime.today() - datetime.timedelta(days=diff)).strftime('%m/%d/%Y')

    def open_xlgc(self, xlgc_url):
        try:
            self.driver.get(xlgc_url)
            self.element_focus('//a[@id="lnkDesktopFindClaim"]')
        except:
            self.driver.refresh()
            self.element_focus('//a[@id="lnkDesktopFindClaim"]')
        self.ElementFinder('//a[@id="lnkDesktopFindClaim"]', 'click')
        self.driver.switch_to.window(self.driver.window_handles[1])

    def search_claim(self, row_values, row):
        self.ElementFinder('//input[contains(@id,"MainContent_txtClaimRefNo")]', 'sendKeys', row_values['claim'])
        self.element_focus('//a[@title="Search - ALT+s"]')
        self.ElementFinder('//a[@title="Search - ALT+s"]', 'click')
        self.ElementFinder('//b[text()="Claim Search Results"]')
        try:
            self.driver.find_element_by_xpath('//table[@id="tblResult"]/descendant::td[@class="clsStatus"]/span')
        except:
            self.driver.find_element_by_xpath('//span[contains(text(),"No Results were found.")]')
            self.element_focus('//a[@title="Modify Search - ALT+m"]')
            self.ElementFinder('//a[@title="Modify Search - ALT+m"]', 'click')
            return 'wrong_claim'

    def xlgc_data_entry(self, row_values, row):
        if row_values['bottomline_status'] != 'Done' \
        and row_values['documentum_status'] == 'Done':
            # initialize variables
            self.row_status = ''
            self.claim_status = ''
            # search claim number in xlgc
            if self.search_claim(row_values, row) == 'wrong_claim':
                self.details_sht.range('{0}{1}'.format(self.col['bottomline_status'], row)).value = 'Incorrect claim number'
                return
            if self.driver.find_element_by_xpath('//table[@id="tblResult"]/descendant::td[@class="clsStatus"]/span').get_attribute('title') == 'Closed':
                self.claim_status = 'Closed'
            self.ElementFinder('//table[@id="tblResult"]/descendant::td[@class="clsClaimNum"]/a','click')
            self.driver.switch_to.window(self.driver.window_handles[2])
            if self.check_reserves(row_values,row) == 'not_available':
                if self.add_reserves(row_values, row) == 'insufficient_authority':
                    return
            if self.check_duplicates(row_values,row) == 'yes':
                return
            # uncomment this code to enable vendor payment
            # if row_values['vendor_status'] != 'Done':
            #     self.post_vendor_payment = True
            #     self.post_payment(row_values, row)
            #     self.details_sht.range('{0}{1}'.format(self.col['vendor_status'], row)).value = self.row_status
            if row_values['bottomline_status'] != 'Done':
                self.post_vendor_payment = False
                # uncomment this to create file notes
                # if self.file_note(row_values, row) == 'not_created':
                #     return
                try:
                    self.post_payment(row_values, row)
                    self.details_sht.range('{0}{1}'.format(self.col['bottomline_status'], row)).value = self.row_status
                except:
                    try:
                        self.row_status = self.driver.find_element_by_xpath(
                            '//span[contains(text(),"The following errors have occurred:")]').get_attribute('innerText')
                    except:
                        self.row_status = self.driver.find_element_by_xpath(
                            '//h3[contains(text(),"An unexpected error has occurred in XLGC")]').get_attribute('innerText')
            self.close_xlgc()

    def check_reserves(self, row_values, row):
        # check available reserves
        reserves = str(self.driver.find_element_by_xpath(
            '(//table[@id="ReserveData"]/descendant::tr[./descendant::span[text()="Non-Legal (Chargeable)"]]/td)[4]').get_attribute(
            'innerText')).replace(',','')

        # if (float(reserves) - row_values['ttp'] + row_values['btt']) < 0:
        if (float(reserves) - row_values['btt']) < 0:
            # self.details_sht.range('{0}{1}'.format(self.col['bottomline_status'], row)).value = 'Insufficient Reserves'
            # self.close_xlgc()
            return 'not_available'
        else:
            return 'available'

    def check_duplicates(self, row_values,row):
        # click on financials > view posted payments
        self.ElementFinder('//tr[contains(@id,"HeaderMenu")]/descendant::a[text()="Financials"]')
        hover = ActionChains(self.driver).move_to_element(
            self.driver.find_element_by_xpath('//tr[contains(@id,"HeaderMenu")]/descendant::a[text()="Financials"]'))
        hover.perform()
        self.ElementFinder('//a[text()="View Posted Payments"]', 'click')
        # check duplicate posted payments within last 1 year
        posted_payments = self.driver.find_elements_by_xpath('//div[@class="gridbackground"]/descendant::tr[@class="oddrow" or @class="evenrow"]')
        for payment in posted_payments:
            try:
                posted_date = payment.find_element_by_xpath('./td[3]/span').get_attribute('innerText')
                posted_date = datetime.datetime.strptime(posted_date, '%m/%d/%Y')
                date_diff = self.today_strp - posted_date
                vendor = str(payment.find_element_by_xpath('./td[4]/span').get_attribute('innerText'))
                amount = str(payment.find_element_by_xpath('./td[5]/span').get_attribute('innerText')).replace(' USD','').replace(',','')
                if float(amount) - row_values['btt'] == 0:
                    if fuzz.ratio(vendor.lower(), 'Bottomline Technologies (de), Inc.'.lower()) > 70:
                    # if vendor.lower().strip() == 'Bottomline Technologies (de), Inc.'.lower():
                        if date_diff.days < 367:
                            self.details_sht.range('{0}{1}'.format(self.col['bottomline_status'], row)).value = 'Duplicate payment'
                            self.close_xlgc()
                            return 'yes'
            except:
                pass
        return 'no'

    def check_participants(self, row_values,row):

        address_match = False
        address_list = []

        participants = self.driver.find_elements_by_xpath(
            '//span[text()="Check Participant to select"]/ancestor::table[1]/descendant::tr')

        # select participant name for bottomline
        if self.post_vendor_payment == False:
            for participant in participants:
                try:
                    participant_name = participant.find_element_by_xpath('./td[2]').get_attribute('innerText')
                except:
                    continue
                if fuzz.ratio('Bottomline Technologies (de), Inc.'.lower(), participant_name.split('-')[0].lower()) > 70:
                # if 'Bottomline Technologies (de), Inc.'.lower() == participant_name.split('-')[0].lower().strip():
                    address_match = True
                    break

        if self.post_vendor_payment == True:
            # select participant name for vendor
            for participant in participants:
                try:
                    participant_name = participant.find_element_by_xpath('./td[2]').get_attribute('innerText')
                except:
                    continue
                if fuzz.ratio(row_values['vendor'].lower(), participant_name.split('-')[0].lower()) > 70:
                # if row_values['vendor'].lower() == participant_name.split('-')[0].lower().strip():

                    self.driver.execute_script("arguments[0].click();", participant.find_element_by_xpath('./td[2]/a'))
                    try:
                        self.ElementFinder('//span[contains(text(),"Address:")]/ancestor::tr[1]')
                    except:
                        self.ElementFinder('//a[@title="Back - ALT+b"]', 'click')
                        self.ElementFinder('//a[@title="Back - ALT+b"]', 'click')
                        break
                    # create address list in case multiple address exist.
                    address_text = self.driver.find_element_by_xpath(
                        '//span[contains(text(),"Address:")]/ancestor::th/following-sibling::td').get_attribute(
                        'innerText')
                    address_rows = self.driver.find_elements_by_xpath(
                        '//span[contains(text(),"Address:")]/ancestor::tr[1]/following-sibling::tr')
                    for address_row in address_rows:
                        if re.search("Address:", address_row.find_element_by_xpath('./th').get_attribute("innerText")):
                            address_list.append(address_text)
                            address_text = ''
                        address_text = f"{address_text} {address_row.find_element_by_xpath('./td').get_attribute('innerText')}"
                    address_list.append(address_text)

                    # if any address matches then return to post payment.
                    for address in address_list:
                        address = address.lower().replace('united states','')
                        if fuzz.token_sort_ratio(row_values['address'].lower().replace('united states',''), address) > 80:
                        # if row_values['address'].lower().replace('united states', '') == address:
                            address_match = True
                    self.ElementFinder('//a[@title="Back - ALT+b"]', 'click')
                    self.ElementFinder(f'//a[contains(text(),"{participant_name}")]/ancestor::tr[1]/td[1]/input')
                    if address_match == True:
                        break

        if address_match == True:
            self.driver.execute_script("arguments[0].click();", self.driver.find_element_by_xpath(
                f'//a[contains(text(),"{participant_name}")]/ancestor::tr[1]/td[1]/input'))
            self.ElementFinder('//a[@title="Continue - ALT+c"]', 'click')
            return

        self.row_status = 'Participant not matching'
        return 'not_found'

    def post_payment(self, row_values, row):
        # self.driver.find_element_by_xpath('//span[contains(text(),"The following errors have occurred:")]').get_attribute('innerText')
        file_note_exist = False
        expert_type = 'Adjuster'
        payment_method = ''
        if self.post_vendor_payment == True:
            invoice_no = str(int(row_values['inv'])) if type(row_values['inv']) == float else str(row_values['inv'])
            inv_amt = str(row_values['ttp'])
            comment = f'inv: {invoice_no}'
            vendor_name = row_values['vendor']
            # check if vendor is expert or adjuster
            for exp_row in range(2, self.expert_last_row + 1):
                if fuzz.ratio(row_values['vendor'].lower(), str(self.expert_sht.range(f'A{exp_row}').value).lower()) > 70:
                # if row_values['vendor'].lower().split()[0] in str(self.expert_sht.range(f'A{exp_row}').value).lower():
                    expert_type = 'Expert'
                    break
        elif self.post_vendor_payment == False:
            invoice_no = str(int(row_values['ref'])) if type(row_values['ref']) == float else str(row_values['ref'])
            inv_amt = str(row_values['btt'])
            comment = f'Ref: {invoice_no}'
            vendor_name = 'Bottomline Technologies (de), Inc.'

        # post vendor payment
        self.ElementFinder('//tr[contains(@id,"HeaderMenu")]/descendant::a[text()="Financials"]', 'click')
        self.ElementFinder('//a[text()="Create Payment"]', 'click')
        self.ElementFinder('//span[contains(text(),"Payee Name:")]/ancestor::tr[1]/descendant::a[@title="Select - ALT+s"]', 'click')
        self.ElementFinder('//span[text()="Check Participant to select"]/ancestor::table[1]/descendant::tr')
        if self.check_participants(row_values, row) == 'not_found':
            self.add_participant(row_values, row)
        self.ElementFinder('//select[contains(@id,"PaymentMethod")]')
        # check if vendor exists in vendor banking details sheet
        for bank_row in range(2, self.banking_last_row + 1):
            if fuzz.ratio(vendor_name.lower(), str(self.banking_sht.range(f'A{bank_row}').value).lower()) > 70:
            # if vendor_name.lower() == str(self.banking_sht.range(f'A{bank_row}').value).lower():
                payment_method = 'ACH'
                sort = self.banking_sht.range(f'B{bank_row}').value
                account = self.banking_sht.range(f'D{bank_row}').value
                sort = str(int(sort)) if type(sort) == float else str(sort)
                account = str(int(account)) if type(account) == float else str(account)
                break
        # select payment method as ACH if vendor is available in vendor banking details otherwise select ACQ
        if payment_method == 'ACH' or self.post_vendor_payment == False:
            bank_details_exist = False
            payment_method_dd = Select(self.driver.find_element_by_xpath('//select[contains(@id,"PaymentMethod")]'))
            payment_method_dd.select_by_value('ACH')
            WebDriverWait(self.driver, 30).until(EC.invisibility_of_element_located((By.XPATH,
                '//div[contains(@id,"UpdateStatus") and @aria-hidden="false"]')))
            # select bank account
            self.ElementFinder('//a[contains(@id,"SelectBankAccount")]', 'click')
            self.ElementFinder('//input[contains(@id,"rdoSelect")]/ancestor::tr[1]')
            bank_detail_rows = self.driver.find_elements_by_xpath('//input[contains(@id,"rdoSelect")]/ancestor::tr[1]')
            for bank_detail_row in bank_detail_rows:
                if re.search(r'{0}'.format(sort), bank_detail_row.find_element_by_xpath('./td[2]').get_attribute('innerText')) and \
                re.search(r'{0}'.format(account), bank_detail_row.find_element_by_xpath('./td[3]').get_attribute('innerText')):
                    self.driver.execute_script("arguments[0].click();", bank_detail_row.find_element_by_xpath(
                        './descendant::input[@type="radio"]'))
                    self.ElementFinder('//a[@title="Continue - ALT+c"]', 'click')
                    bank_details_exist = True
                    break
            if bank_details_exist == False:
                self.ElementFinder('//a[@title="Back - ALT+b"]', 'click')
                self.row_status = 'Bank details not matching'
                return
        else:
            payment_method_dd = Select(self.driver.find_element_by_xpath('//select[contains(@id,"PaymentMethod")]'))
            payment_method_dd.select_by_value('ACQ')
            WebDriverWait(self.driver, 30).until(EC.invisibility_of_element_located((By.XPATH,
                        '//div[contains(@id,"UpdateStatus") and @aria-hidden="false"]')))
        # Fill Reference number
        self.ElementFinder('//input[contains(@id,"RefrenceInvoiceNumber")]', 'sendKeys', invoice_no[-10:])
        self.ElementFinder('//input[contains(@id,"InvoiceAmt")]', 'sendKeys', '')
        WebDriverWait(self.driver, 30).until(EC.invisibility_of_element_located((By.XPATH,
            '//div[contains(@id,"UpdateStatus") and @aria-hidden="false"]')))
        # Fill Invoice Date
        try:
            WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.XPATH, '//input[contains(@id,"ReceivedDateforInvoice")]')))
            self.ElementFinder('//input[contains(@id,"ReceivedDateforInvoice")]', 'sendKeys', str(self.previous_wd_strf))
        except:
            pass
        # Fill Invoice amount
        self.ElementFinder('//input[contains(@id,"InvoiceAmt")]', 'sendKeys', inv_amt)
        # fill comments
        self.ElementFinder('//textarea[contains(@id,"txtComments")]', 'sendKeys', comment)
        # Add file note
        self.ElementFinder('//a[contains(@id,"btnAddFileNote")]', 'click')
        # Select file note
        self.ElementFinder('//table[@id="HTMLSectionCopyID"]/descendant::td[@class="tdDescription"]')
        file_notes = self.driver.find_elements_by_xpath('//table[@id="HTMLSectionCopyID"]/descendant::td[@class="tdDescription"]')
        for file_note in file_notes:
            file_note_text = file_note.get_attribute('innerText')
            if str(row_values['ttp']).replace(',','') in file_note_text.replace(',','') and \
            str(row_values['btt']).replace(',','') in file_note_text.replace(',','') and \
            fuzz.partial_ratio(row_values['vendor'].lower(),file_note_text.lower()) > 90:
                self.driver.execute_script("arguments[0].click();", file_note.find_element_by_xpath(
                './ancestor::tr/td[1]/input[@type="checkbox"]'))
                file_note_exist = True
        if file_note_exist == False:
            self.row_status = 'File note was not found'
            return
        self.ElementFinder('//a[@title="Continue - ALT+c"]', 'click')
        self.ElementFinder('//a[@title="Continue - ALT+c"]', 'click')
        # select payment type as expense payment
        self.ElementFinder('//select[contains(@id,"PaymentType")]')
        payment_type_dd = Select(self.driver.find_element_by_xpath('//select[contains(@id,"PaymentType")]'))
        payment_type_dd.select_by_visible_text('Expense Payment')
        WebDriverWait(self.driver, 30).until(EC.invisibility_of_element_located((By.XPATH,
                '//div[contains(@id,"UpdateStatus") and @aria-hidden="false"]')))
        # select payment method
        payment_method_dd = Select(self.driver.find_element_by_xpath('//select[contains(@id,"ServiceBenefitType")]'))
        if expert_type == 'Adjuster':
            payment_method_dd.select_by_visible_text('Adjuster Fees')
        elif expert_type == 'Expert':
            payment_method_dd.select_by_visible_text('Expert Fees')
        WebDriverWait(self.driver, 30).until(EC.invisibility_of_element_located((By.XPATH,
            '//div[contains(@id,"UpdateStatus") and @aria-hidden="false"]')))
        # Fill Invoice amount
        self.ElementFinder('//input[contains(@id,"ApprovedAmount")]', 'sendKeys', inv_amt)
        # save
        self.ElementFinder('//a[@title="Save - ALT+s"]', 'click')
        WebDriverWait(self.driver, 30).until(EC.invisibility_of_element_located((By.XPATH,
            '//span[text()="There are no Regular Payment Items to display for this claim."]')))
        WebDriverWait(self.driver, 30).until(EC.invisibility_of_element_located((By.XPATH,
            '//div[contains(@id,"UpdateStatus") and @aria-hidden="false"]')))
        # continue
        self.ElementFinder('//a[@title="Continue - ALT+c"]', 'click')
        # Finish later
        # self.ElementFinder('//a[@title="Finish Later - ALT+f"]', 'click')
        self.ElementFinder('//a[@title="Post - ALT+p"]')
        self.details_sht.range('{0}{1}'.format(self.col['URL'], row)).value = self.driver.current_url
        self.ElementFinder('//a[@title="Post - ALT+p"]', 'click')
        try:
            # accept alert
            WebDriverWait(self.driver, 10).until(EC.alert_is_present(), 'Click `OK` to continue or `Cancel` to return to Payment Summary')
            alert = self.driver.switch_to.alert
            alert.accept()
            print('pop up clicked')
        except:
            pass
        self.row_status = 'Done'

    def file_note(self, row_values, row):

        client_inv_id = re.search(r'(ClientInvoiceID#)(.*)(.msg)', row_values['subject'])
        client_inv_id = client_inv_id.group(2).replace(' ', '')
        if os.path.isfile(os.path.join(os.path.abspath(os.getcwd()), 'Emails', f'{client_inv_id}.msg')):
            mail_path = os.path.join(os.path.abspath(os.getcwd()), 'Emails', f'{client_inv_id}.msg')
        else:
            self.details_sht.range('{0}{1}'.format(self.col['bottomline_status'], row)).value =\
                'File note attachment was not found in Emails folder.'
            self.close_xlgc()
            return 'not_created'
        expert_type = 'Adjuster'
        for exp_row in range(2, self.expert_last_row + 1):
            if fuzz.ratio(row_values['vendor'].lower(), str(self.expert_sht.range(f'A{exp_row}').value).lower()) > 70:
            # if row_values['vendor'].lower().split()[0] in str(self.expert_sht.range(f'A{exp_row}').value).lower():
                expert_type = 'Expert'
                break
        title = '{0} - Payment of USD {1} {2} to {3} USD {4} Adjuster to Bottomline'.format(
            self.today_strf, row_values['ttp'], expert_type, row_values['vendor'], row_values['btt'])
        # click on file note header
        self.ElementFinder('//tr[contains(@id,"HeaderMenu")]/descendant::a[text()="File Notes"]', 'click',)
        # select financials in category
        self.ElementFinder('//select[contains(@id,"cboCategory")]')
        category_dd = Select(self.driver.find_element_by_xpath('//select[contains(@id,"cboCategory")]'))
        category_dd.select_by_visible_text('Financials')
        WebDriverWait(self.driver, 30).until(EC.invisibility_of_element_located((By.XPATH, '//div[contains(@id,"UpdateStatus") and @aria-hidden="false"]')))
        # select payments in sub category
        sub_category_dd = Select(self.driver.find_element_by_xpath('//select[contains(@id,"cboSubCategory")]'))
        sub_category_dd.select_by_visible_text('Payments')
        WebDriverWait(self.driver, 30).until(EC.invisibility_of_element_located((By.XPATH, '//div[contains(@id,"UpdateStatus") and @aria-hidden="false"]')))
        # enter description
        self.ElementFinder('//textarea[contains(@id,"txtLiveDescript")]', 'sendKeys', title)
        # click on attachment header
        self.ElementFinder('//a[contains(@title,"Distribution/Attachment - ALT+a")]', 'click')
        WebDriverWait(self.driver, 30).until(EC.invisibility_of_element_located((By.XPATH, '//div[contains(@id,"UpdateStatus") and @aria-hidden="false"]')))
        # attach email
        self.ElementFinder('//span[contains(text(),"Drop or Paste item into area below")]')
        attachment_tag = self.driver.find_element_by_xpath('//div[contains(@id,"AttachmentsDragDrop")]/descendant::input[@id="file"]')
        attachment_tag.send_keys(mail_path)
        # click on save attachments
        self.ElementFinder('//a[contains(@id,"btnSaveAttachment")]', 'click')
        WebDriverWait(self.driver, 30).until(EC.invisibility_of_element_located((By.XPATH, '//div[contains(@id,"UpdateStatus") and @aria-hidden="false"]')))
        self.ElementFinder('//textarea[contains(@id,"txtAttachmentDescription")]')
        # click on save
        self.ElementFinder('//a[@title="Save - ALT+s"]', 'click')
        WebDriverWait(self.driver, 30).until(EC.invisibility_of_element_located((By.XPATH, '//div[contains(@id,"UpdateStatus") and @aria-hidden="false"]')))

    def claim_professional(self, row_values, row):
        if self.search_claim(row_values, row) == 'wrong_claim':
            self.details_sht.range('{0}{1}'.format(self.col['documentum'], row)).value = 'Incorrect claim number'
            return
        claim_prof = self.driver.find_element_by_xpath('//table[@id="tblResult"]/descendant::a[contains(@title,"Claim Owner:")]').get_attribute('innerText')
        if re.search(r'^([A-Z][a-z]+\s)([A-Z][a-z]+)', str(claim_prof)):
            prof_last_name = re.search('^([A-Z][a-z]+\s)([A-Z][a-z]+)', str(claim_prof)).group(2)
            if str(row_values['claim_prof']).lower().find(prof_last_name.lower()) != -1:
                self.details_sht.range('{0}{1}'.format(self.col['documentum'], row)).value = f'Match: {claim_prof}'
            else:
                self.details_sht.range('{0}{1}'.format(self.col['documentum'], row)).value = f'Different: {claim_prof}'
        self.element_focus('//a[@title="Modify Search - ALT+m"]')
        self.ElementFinder('//a[@title="Modify Search - ALT+m"]', 'click')

    def add_participant(self, row_values, row):

        self.ElementFinder('//a[@title="Add a Vendor"]','click')
        self.ElementFinder('//select[contains(@id,"cboSearchBy")]')
        search_by = Select(self.driver.find_element_by_xpath('//select[contains(@id,"cboSearchBy")]'))
        search_by.select_by_visible_text('Company Name')
        self.loading()
        self.ElementFinder('//input[contains(@id,"txtCompanyName")]', 'sendKeys', 'bottomline')
        self.ElementFinder('//a[@title="Search - ALT+s"]', 'click')
        self.ElementFinder('//a[./ancestor::td/following-sibling::td/span[text()="Bottomline Technologies (de), Inc."]]', 'click')
        self.ElementFinder('//input[@name="rgpSelectVendorLocPerformer" and ./ancestor::td/following-sibling::td/span[text()="Bottomline Technologies (de), Inc."]]', 'click')
        self.ElementFinder('//a[contains(@id,"btnAttachBottom")]', 'click')
        self.loading()
        self.ElementFinder('//select[contains(@id,"lstInvRoles")]')
        roles = Select(self.driver.find_element_by_xpath('//select[contains(@id,"lstInvRoles")]'))
        roles.select_by_visible_text('Audit Service')
        try:
            aob = Select(self.driver.find_element_by_xpath('//select[contains(@id,"cboAOBType")]'))
            aob.select_by_visible_text('No')
        except:
            pass
        self.ElementFinder('//a[@title="Save - ALT+s"]', 'click')
        self.loading()
        # select vendor once it's added
        self.ElementFinder('//tr[contains(@id,"HeaderMenu")]/descendant::a[text()="Financials"]', 'click')
        self.ElementFinder('//a[text()="Create Payment"]', 'click')
        self.ElementFinder('//span[contains(text(),"Payee Name:")]/ancestor::tr[1]/descendant::a[@title="Select - ALT+s"]', 'click')
        self.ElementFinder('//span[text()="Check Participant to select"]/ancestor::table[1]/descendant::tr')
        self.check_participants(row_values, row)

    def add_reserves(self, row_values, row):

        self.ElementFinder('//tr[contains(@id,"HeaderMenu")]/descendant::a[text()="Financials"]', 'click')
        try:
            if self.claim_status == 'Closed':
                self.ElementFinder('//a[text()="Re-open Expense Financials"]', 'click')
            else:
                self.ElementFinder('//a[text()="Change Expense Reserve"]', 'click')
        except:
            self.details_sht.range('{0}{1}'.format(self.col['bottomline_status'],row)).value = 'Unable to add reserves'
            self.close_xlgc()
            return 'insufficient_authority'

        self.loading()
        self.ElementFinder('//input[./ancestor::td/following-sibling::td/span[text()="Non-Legal (Chargeable)"]]', 'click')
        self.ElementFinder('//a[@title="Edit - ALT+e"]', 'click')
        self.ElementFinder('//input[contains(@id,"tcuAmount")]', 'sendKeys', math.ceil(row_values['btt']))
        self.ElementFinder('//a[@title="Save - ALT+s"]', 'click')
        self.loading()
        self.ElementFinder('//a[@title="Post - ALT+p"]', 'click')
        try:
            # accept alert
            WebDriverWait(self.driver, 20).until(EC.alert_is_present(), 'Would you like to continue?')
            alert = self.driver.switch_to.alert
            alert.accept()
            print('pop up clicked')
        except:
            pass
        # Insufficient authority
        try:
            self.ElementFinder('//a[text()="Create Payment"]')
        except:
            self.ElementFinder('//span[contains(text(),"Insufficient Authority")]')
            self.ElementFinder('//a[@title="Continue - ALT+c"]', 'click')
            self.loading()
            self.ElementFinder('//a[contains(@id,"btnSelectGrantor")]', 'click')
            self.loading()
            self.ElementFinder('//input[./ancestor::td/following-sibling::td/span[text()="Claim Owner"]]', 'click')
            self.ElementFinder('//a[@title="Continue - ALT+c"]', 'click')
            self.loading()
            self.ElementFinder('//a[@title="Submit - ALT+s"]', 'click')
            self.loading()
            self.details_sht.range('{0}{1}'.format(self.col['bottomline_status'], row)).value = 'Sent for approval because of insufficient authority'
            self.close_xlgc()
            return 'insufficient_authority'


    def loading(self):
        WebDriverWait(self.driver, 30).until(EC.invisibility_of_element_located((By.XPATH, '//div[contains(@id,"UpdateStatus") and @aria-hidden="false"]')))

    def close_xlgc(self):
        # close the window
        self.driver.close()
        self.driver.switch_to.window(self.driver.window_handles[1])
        self.element_focus('//a[@title="Modify Search - ALT+m"]')
        self.ElementFinder('//a[@title="Modify Search - ALT+m"]', 'click')

    def xlgc_restart(self):

        num_handle = self.driver.window_handles.__len__()
        if num_handle > 0:
            for num in range(1, num_handle):
                self.driver.switch_to.window(self.driver.window_handles[1])
                self.driver.close()
        self.driver.switch_to.window(self.driver.window_handles[0])


