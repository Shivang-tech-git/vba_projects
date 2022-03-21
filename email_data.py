import win32com.client, os, datetime, shutil
import re
from datetime import datetime, timedelta

class email_extract:

    def __init__(self, all_folder_sequence, all_prefix, lob_trigger, details_sht, col, today_strf):
        self.details_sht = details_sht
        self.col = col
        self.today_strf = today_strf
        self.all_prefix = all_prefix
        outlook = win32com.client.Dispatch('outlook.application')
        mapi = outlook.GetNamespace('MAPI')
        self.all_required_folders = [None] * len(all_folder_sequence)
        for folder_sequence in range(len(all_folder_sequence)):
            if lob_trigger[folder_sequence] == 'Yes':
                self.all_required_folders[folder_sequence] = mapi
                for folder in all_folder_sequence[folder_sequence]:
                    for index in range(1, self.all_required_folders[folder_sequence].Folders.Count + 1):
                        # print(self.all_required_folders[folder_sequence].Folders.Item(index).Name)
                        if str(self.all_required_folders[folder_sequence].Folders.Item(index).Name).strip().lower() == str(folder).strip().lower():
                            self.all_required_folders[folder_sequence] = self.all_required_folders[folder_sequence].Folders.Item(index)
                            break

        # filter inbox for recieve date as one week
        received_dt = datetime.now() - timedelta(days=7)
        received_dt = received_dt.strftime('%m/%d/%Y %H:%M %p')
        self.messages = [None] * len(self.all_required_folders)
        for required_folder in range(len(self.all_required_folders)):
            if not self.all_required_folders[required_folder] is None:
                self.messages[required_folder] = self.all_required_folders[required_folder].Items.Restrict("[ReceivedTime] >= '" + received_dt + "'")

        # create folder for Emails and attachments
        folder_names = ['Invoice']
        for folder_name in folder_names:
            if os.path.exists(os.path.join(os.path.abspath(os.getcwd()), folder_name)):
                shutil.rmtree(os.path.join(os.path.abspath(os.getcwd()), folder_name))
            os.makedirs(os.path.join(os.path.abspath(os.getcwd()), folder_name))

    def get_email_data(self, invoice_num, prefix, row):
        # filter for invoice number in subject.
        mail_found = False
        for msg_counter in range(len(self.messages)):
            if prefix == self.all_prefix[msg_counter] and self.messages[msg_counter] is not None:
                self.sub_filter = self.messages[msg_counter].Restrict(f"@SQL=(urn:schemas:httpmail:subject LIKE '%{invoice_num}%')")
                mail_found = True
                break
        if mail_found == False:
            return
        # if claim number is found in body then find all other details
        for mail in list(self.sub_filter):
            self.claim = re.search(r'(Claim\s#: )(\d+)', mail.body)
            # self.claim = re.search(r'(Claim\s#: )(.*)', mail.body)
            if self.claim:
                # mail.SaveAs(os.path.join(os.path.abspath(os.getcwd()), 'Emails', f'{client_invoice_id}.msg'))
                for attachment in mail.Attachments:
                    if 'sfi' in attachment.FileName or 'bar' in attachment.FileName:
                        attachment.SaveAsFile(os.path.join(os.path.abspath(os.getcwd()), 'Invoice', attachment.FileName))
                self.claim = self.claim.group(2)
                self.vendor = re.search(r'(Vendor Name: )(.*)', mail.body).group(2).replace('\r','')
                self.total_to_pay = re.search(r'(Total To Pay: )(.*)', mail.body).group(2).replace('\r','')
                self.total_to_pay = self.total_to_pay.replace('$','').replace(',','')
                self.invoice = re.search(r'(Invoice #: )(.*)', mail.body).group(2).replace('\r','')
                self.bt_total = re.search(r'(BT Total: )(.*)', mail.body).group(2).replace('\r','')
                self.bt_total = self.bt_total.replace('$','').replace(',','')
                self.claim_prof = re.search(r'(Claim Professional: )(.*)', mail.body).group(2).replace('\r','')
                # give output to details sheet

                self.details_sht.range('{0}{1}'.format(self.col['run_date'], row)).value = self.today_strf
                self.details_sht.range('{0}{1}'.format(self.col['claim'], row)).value = self.claim
                self.details_sht.range('{0}{1}'.format(self.col['vendor'], row)).value = self.vendor
                self.details_sht.range('{0}{1}'.format(self.col['ttp'], row)).value = self.total_to_pay
                self.details_sht.range('{0}{1}'.format(self.col['inv'], row)).value = self.invoice
                self.details_sht.range('{0}{1}'.format(self.col['btt'], row)).value = self.bt_total
                self.details_sht.range('{0}{1}'.format(self.col['claim_prof'], row)).value = self.claim_prof
                break


