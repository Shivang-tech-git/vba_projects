from selenium import webdriver
from webdriver_manager.microsoft import EdgeChromiumDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from msedge.selenium_tools import Edge, EdgeOptions
from selenium.webdriver.ie.options import Options
from tkinter import messagebox
from tkinter import filedialog
import main_dataframe
import all_information
from pathlib import Path
from datetime import datetime
import os, getpass
import pandas
import tkinter as tk
from tkinter import ttk
import xlwings
import traceback

#---------- XPath ------ one element ----------------

class web_control:
    def __init__(self):

        # os.environ["HTTP_PROXY"] = ""
        # os.environ["HTTPS_PROXY"] = ""
        self.Edge_Options = EdgeOptions()
        self.Edge_Options.add_experimental_option("excludeSwitches", ["enable-automation"])
        self.Edge_Options.add_experimental_option('useAutomationExtension', False)
        self.Edge_Options.use_chromium = True
        self.Edge_Options.add_argument('no-sandbox')
        self.Edge_Options.add_argument("user-data-dir=C:/Users/" + getpass.getuser() + "/AppData/Local/Microsoft/Edge/User Data")
        self.Edge_Options.add_argument("headless")
        self.Edge_Options.add_argument("disable-gpu")
        edge_driver_path = os.path.join(os.path.abspath(os.getcwd()), 'web_drivers', 'msedgedriver')
        try:
            # self.driver = Edge(EdgeChromiumDriverManager().install(), options=Edge_Options)
            self.driver = Edge(executable_path=edge_driver_path, options=self.Edge_Options)
        except:
            messagebox.showinfo('Frames event coding tool',
                                'Please close Edge and then run the tool.')

        # chrome
        # chromedriver_path = os.path.join(os.path.abspath(os.getcwd()), 'web_drivers', 'chromedriver')
        # chromeOptions = webdriver.ChromeOptions()
        # chromeOptions.add_argument("--start-maximized")
        # chromeOptions.add_experimental_option("excludeSwitches", ["enable-automation"])
        # chromeOptions.add_argument("disable-infobars")
        # self.driver = webdriver.Chrome(chromedriver_path, options=chromeOptions)

        # ie
        # ie_driver_path = os.path.join(os.path.abspath(os.getcwd()), 'web_drivers', 'IEDriverServer')
        # ieOptions = Options()
        # ieOptions.add_argument("--start-maximized")
        # self.driver = webdriver.Ie(ie_driver_path, options=ieOptions)

        self.driver.set_page_load_timeout(60)
        self.driver.get('')

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

class frames_home():
    filename = None
    def __init__(self, event):
        if event == 'Property event coding':
            self.selected_event = {'event':'PCS',
                              'output_filename':'PCS Output {}.CSV'.format(datetime.today().strftime('%d-%m-%Y')),
                              'message':'PCS Allocation Completed, Please check the PCS output file to see status.',
                                'sht_name':'Property Claims'}
        elif event == 'CAT event coding':
            self.selected_event = {'event':'CAT',
                              'output_filename':'CAT Output {}.CSV'.format(datetime.today().strftime('%d-%m-%Y')),
                              'message':'CAT Allocation Completed, Please check the CAT output file to see status.',
                                'sht_name':'CAT'}
        elif event == 'Various event coding':
            self.selected_event = {'event':'VARS',
                              'output_filename':'VARS Output {}.CSV'.format(datetime.today().strftime('%d-%m-%Y')),
                              'message':'VARS Allocation Completed, Please check the VARS output file to see status.',
                                'sht_name':'Various'}
        elif event == 'Actuals event coding':
            self.selected_event = {'event':'Actuals',
                              'output_filename':'Actuals Output {}.CSV'.format(datetime.today().strftime('%d-%m-%Y')),
                              'message':'Actuals Allocation Completed, Please check the Actuals output file to see status.',
                                'sht_name':'Actual Events'}
    @classmethod
    def allocation_file(cls):
        frames_home.filename = filedialog.askopenfilename(initialdir="/",
                                              title="Select allocation File")

    def run_event(self):
        # if output file checkbox is not selected
        if output_selected.get() == 0:
            # Check if allocation file exists
            self.excel = xlwings.Book(frames_home.filename)
            self.sht = self.excel.sheets[self.selected_event['sht_name']]
            self.main_df = self.sht.range('A1').options(pandas.DataFrame, header=1, index=False, expand='table').value
            self.main_df['Event'] = ''
            self.main_df['Status'] = ''
            self.main_df['Loss type'] = ''
            self.main_df[['Event', 'Status', 'Loss type']] = self.main_df[['Event', 'Status', 'Loss type']].astype(str)
            self.main_df.to_csv(self.selected_event['output_filename'], index=False)
            self.blank_df = self.main_df
        # if output file checkbox is selected
        elif output_selected.get() == 1:
            self.main_df = pandas.read_csv(self.selected_event['output_filename'])
            self.blank_df = self.main_df.loc[self.main_df['Status'].isnull()]

    def run_refresh(self):
        self.main_df = pandas.read_csv(self.selected_event['output_filename'])
        self.ue_df = self.main_df.loc[self.main_df['Status'] == 'Undefined Error']
        self.ue_df[['Event', 'Status', 'Loss type']] = self.ue_df[['Event', 'Status', 'Loss type']].astype(str)

def run_instance():
    data = frames_home(event_selected.get())
    try:
        data.run_event()
    except:
        messagebox.showinfo('Frames event coding tool',
                            'Please select allocation file and then run the tool.')
        return
    web = web_control()
    run = main_dataframe.frames_workflow(web.driver,
                                         data.selected_event['event'],
                                         data.main_df,
                                         all_information.event_info,
                                         web.ElementFinder)
    try:
        run.event_coding(data.blank_df)
        messagebox.showinfo('Frames event coding tool',
                            data.selected_event['message'])
    except Exception as msg:
        web.driver.quit()
        messagebox.showerror('Frames event coding tool', traceback.format_exc())

def refresh_instance():
    data = frames_home(event_selected.get())
    try:
        data.run_refresh()
    except:
        messagebox.showerror('Frames event coding tool',
                             'Error: Ouptut file does not exist.')
        return
    web = web_control()
    refresh = main_dataframe.frames_workflow(web.driver,
                                             data.selected_event['event'],
                                             data.main_df,
                                             all_information.event_info,
                                             web.ElementFinder)
    try:
        refresh.event_refresh(data.ue_df)
        messagebox.showinfo('Frames event coding tool',
                            data.selected_event['message'])
    except Exception as msg:
        web.driver.quit()
        messagebox.showerror('Frames event coding tool', traceback.format_exc())

if __name__ == '__main__':
    window = tk.Tk()
    output_selected = tk.IntVar()
    event_selected = tk.StringVar()
    window.title('AXA XL')
    # ------------------------------- Title ----------------------------
    greeting = tk.Label(text="Event coding tool", font=('Arial',12), fg='yellow', bg='green', width=60)
    frm_open = tk.Frame(master=window, width=60)
    # ------------------------------ Browse allocation file --------------------------
    lbl_open = tk.Label(master=frm_open, text='Select allocation file',font=('Arial',10),
                        width=30, borderwidth=2, relief="ridge").grid(row=0, column=0, pady=5)
    btn_open = tk.Button(master=frm_open, text='Browse',font=('Arial',10), width=15, bg='lavender',
                         command=frames_home.allocation_file).grid(row=0,column=1,pady=5)
    use_output_file = tk.Checkbutton(master=frm_open, text='Use Output file',font=('Arial',10), width=15, bg='lavender',
                         variable=output_selected).grid(row=0,column=2,pady=5)
    # ------------------------------ combobox ----------------------------------------\
    frm_combobox = tk.Frame(master=window, width=60)
    lbl_combo = tk.Label(master=frm_combobox, text='Select event code',font=('Arial',10),
                           width=30, borderwidth=2, relief="ridge").grid(row=1, column=0, pady=5)
    combobox_events = ttk.Combobox(master=frm_combobox, width = 40, textvariable = event_selected)
    combobox_events.grid(row=1, column=1, pady=5)
    combobox_events['values'] = ('CAT event coding','Property event coding',
                                 'Various event coding', 'Actuals event coding')
    combobox_events.current(1)

    btn_run = tk.Button(master=frm_combobox,font=('Arial',10), text='Run', width=10, bg='floral white',
                        command=run_instance).grid(row=2, column=0, pady=5)
    btn_refresh = tk.Button(master=frm_combobox,font=('Arial',10), text='Refresh', width=10, bg='misty rose',
                        command=refresh_instance).grid(row=3, column=0, pady=5)
    greeting.pack()
    frm_open.pack(fill=tk.BOTH)
    frm_combobox.pack(fill=tk.BOTH)
    window.mainloop()
