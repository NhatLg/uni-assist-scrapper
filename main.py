import requests
import pandas as pd
import numpy as np
import selenium
from selenium.webdriver.common.by import By
import json
from bs4 import BeautifulSoup
from lxml import html
import openpyxl
import os
from os import walk
import glob
import time


os.chdir(r"C:\Users\Admin\Jottacloud\Automate Applications")

# ***************************************************************************
# -------------------------CONFIGURATION-------------------------------------
# ***************************************************************************

current_semester = 'SS2023' # name for the saved excel file, YOU MUST MANUALLY CREATE A FOLDER WITH THE SAME NAME IN THE DIRECTORY BELOW
transcript_save_location = f"C:/Users/Admin/Jottacloud/Ebgo/{current_semester}"

extra_columns = ["econ_credits", "math_credits", "credit_system_eval", "admitted", "reason_reject",
                 "admit_conditions", "admit_type"] # aside from the columns pulled from uni-assist, I want some extra columns in the excel file (so that I can fill in the data)

save_file_name = f'last_dta_ids_{current_semester}.json'  # keep track of what has been pulled and what's not

login_url = 'https://ww2.uni-assist.de/portal/index.php'  # link to the login site of uni-assist
request_url = 'https://ww2.uni-assist.de/portal/index.php?go=doz'  # link to the LATEST semester

payload = {
    "login": "Luong",
    "pass": "`Reg~2020!#",
}


dict_reject = {
    'econ': '...Not enough ECTS were obtained from the fields of economics totaling at least 60 credits.',
    'math': '...Not enough ECTS were obtained from the fields of mathematics or statistics totaling at least 18 credits.',
    'unrelated': '..Acquired degree not suitable for further studies in economics.'
}


dict_req_courses = {
    'econo': "Introduction to Econometrics (6CP) or Introductory Econometrics (6CP),\n",
    'behav': "Introduction to Behavioral Economics (6 CP),\n",
    'gtheo': "Microeconomics using Calculus (6CP) OR Introduction to Game Theory (6 CP),\n",
    'pecon': "Public Economics (6 CP) or European Economics (6CP),\n",
    'elece': "any course from the Bachelor's program Wirtschaftswissenschaften "
             "(6CP, the course can be chosen individually but must refer to economics),\n",
    'elecf': "any course from the Bachelor's program Wirtschaftswissenschaften "
             "(6CP, the course can be chosen individually).\n",
}


try:
    old_df = pd.read_excel(f'application_data{current_semester}.xlsx', engine='openpyxl', index_col=0)
    old_df = old_df.sort_index()
    old_df = old_df.reset_index(drop=True)
    old_df['transcript_id'].astype("int64")
except FileNotFoundError:
    old_df = pd.DataFrame()
# df_new.set_index('index', inplace=True)
# df_new.drop(columns=['level_0'], inplace=True)


# Check if there are transcripts that are not downloaded in the previous run although the data is recorded
# (internet connection problem)
filenames = next(walk(transcript_save_location), (None, None, []))[2]  # [] if no file
downloaded_transcript_ids = []
for i in filenames:
    try:
        downloaded_transcript_ids.append(int(i.split(".")[0]))
    except ValueError:
        print(f"Require manual check for this file, possible that the certificate is not uploaded to uni-assist: {i}")

if not old_df.empty:
    listed_trscript_ids = list(old_df["transcript_id"])
    old_ids = list(old_df["app_ids"])
    old_bew_nrs = list(old_df["bew_nr"])
    missing_trscript_ids = [x for x in listed_trscript_ids if x not in downloaded_transcript_ids]
    old_df_sub = old_df[['transcript_link', 'transcript_id', 'bew_nr']]
    df_missing = pd.merge(pd.DataFrame({'trscript_id': missing_trscript_ids}), old_df_sub,
                       left_on='trscript_id',
                       right_on='transcript_id',
                       how='inner')
    list_miss_links = list(df_missing["transcript_link"])
else:
    missing_trscript_ids = []
    old_ids = []
    old_bew_nrs = []


if missing_trscript_ids:
    print(f"Detect missing transcripts of records in folder: {transcript_save_location}")
    print("Missing transcripts can be result of internet disconnection in the previous run. "
          "Records are shown in the excel file but the transcripts were not succesfully downloaded")
    is_downloading_missing_trscripts = input("Do you want to download missing transcripts?y/n")
    if is_downloading_missing_trscripts == 'y':
        test_df = pd.merge(pd.DataFrame({'trscript_id': missing_trscript_ids}), old_df_sub, left_index=True, right_on='transcript_id', how='right')
        print("Added to downloading list")
    elif is_downloading_missing_trscripts == 'n':
        is_downloading_missing_trscripts = 0
        print("Skipped")
    else:
        print("Can not understand instruction. Skipped. You can re-run the code to download this later")
elif not missing_trscript_ids:
    is_downloading_missing_trscripts = 0

#df includes everything currently on uni-assist
#df_new includes only new applicants that you have not saved in your local files
#old_df includes only applicants that you have saved in your local files

def click_dropdown(driver, element_name, option_text):
    """
    The function click on the "semester wechseln" dropdown menu on uni-assist website. This function requires Selenium
    and a browswer driver to use. Tested with Chrome latest driver.
    :param driver:
    :param element_name:
    :param option_text:
    :return:
    """
    el = driver.find_element(By.NAME, element_name)
    for option in el.find_elements(By.TAG_NAME, 'option'):
        if option.text == option_text:
            option.click()
            break

#login with username and password, then with request pull all info from the current semester evaluation main table
with requests.Session() as session:
    post = session.post(login_url, data=payload)
    r = session.get(request_url)

tables = pd.read_html(r.text)  # pandas library has a function that read ALL tables on a websites and put it in an array
df = tables[2]  # the main table (list all applicants) is in the index 2

soup = BeautifulSoup(r.text, 'lxml')


bewerten_links = [] # contain all links to detailed application page of individual participant
"""
all applicants id for the current semester so far (including those have been processed) 
applicaiton id is the id embeded in the link to an individual application profile, 
e.g: https://ww2.uni-assist.de/portal/index.php?go=doz&do=anza&dozeid=984090
984090 is the application id stored in "app_ids"
this app_ids are not documents id
"""
app_ids = []
circle_img = [] # all applicants processing status 'g' is green (processed) 'r' is red (not processed) 'y' is yellow (not finished)
bew_nr = []

for a in soup.find_all('a', href=True, title='bewerten'):
    current_link = 'https://ww2.uni-assist.de/portal/' + a['href']
    current_id = current_link.split("dozeid=", 1)[1]
    # get the src of the image right after "Bewerten", then get the 5th letter of the src string (g for green, r for red)
    red_green_img = a.find_next_sibling('img')['src'][4]
    # append link, id, and img to a list
    circle_img.append(red_green_img)
    bewerten_links.append(current_link)
    app_ids.append(current_id)


for a in soup.find_all('a', href=True, title= 'Bewerberdetails anzeigen'):
    href = a['href']
    bew_num = href[23:30]
    bew_nr.append(bew_num)


is_processed_app = [1 if i == 'g' else 0 for i in circle_img]

""" 
df contains ALL APPLICATIONS whether processed or not
Use is_processed column to check if the application is processed
"""
df['bew_nr'] = bew_nr
df['bewerten_links'] = bewerten_links
df['is_processed'] = is_processed_app
new = df['bewerten_links'].str.split(pat="dozeid=", n=1, expand=True)
df['app_ids'] = new[1].astype(int)
df['bew_nr'] = df['bew_nr'].astype(int)
bew_nr = list(map(int, bew_nr)) # convert to integer
dict_is_processed = dict(zip(bew_nr, is_processed_app))

# Select only new bew_nr
new_bew_nrs = [x for x in bew_nr if x not in old_bew_nrs]


# Removed rows that have the old ids (those that have been downloaded before)
df_new = df[df['bew_nr'].isin(new_bew_nrs)]

time.sleep(4)
df2 = pd.DataFrame()
df_new.reset_index(inplace=True)
if not df_new.empty:
    with requests.Session() as session:
        post = session.post(login_url, data=payload)
        # With the detailed link of each applicant:
        for i in range(0, len(df_new['bewerten_links'])):
            request_url = df_new['bewerten_links'][i]
            r = session.get(request_url)
            #Get the detailed table of each applicants
            applicant_table = pd.read_html(r.text)
            applicant_info = applicant_table[2]

            #Transform the table before concat to large df
            applicant_info = applicant_info.transpose()
            applicant_info.columns = applicant_info.iloc[0]
            applicant_info.drop(applicant_info.index[0], inplace=True)

            #Get transcript.pdf links
            soup = BeautifulSoup(r.text, 'lxml')
            zeugnisse_element = soup.find_all('td', text="Zeugnisse ")
            next_element = zeugnisse_element[0].next_sibling.next_sibling
            applicant_info['transcript_link'] = 'https://ww2.uni-assist.de/portal/' + next_element.find('a')['href']

            df2 = pd.concat([df2, applicant_info])

    df2 = df2.reset_index(drop=True)
    df_new = pd.concat([df_new, df2], axis=1)

    regex_date = r"(\d{1,2}[/. ](?:\d{1,2}|January|Jan)[/. ]\d{2}(?:\d{2})?)"
    df_new = df_new.loc[:, ~df_new.columns.duplicated()]
    df_new['uniassist_date'] = df_new['Antrag'].str.extract(regex_date)
    df_new['uniassist_date'] = pd.to_datetime(df_new['uniassist_date'], dayfirst=True).dt.date

    df_new['transcript_id'] = df_new['transcript_link'].str.slice(-10, stop=None, step=1)
    df_new['ready_upload'] = 0

elif df_new.empty:
    print('Nothing new in uni-assist')

# Append some empty column into this dataframe before saving
if old_df.empty:    # First time run, new file
    for i in extra_columns:
        df_new[i] = ""
    df_new = df_new.sort_values('uniassist_date', ascending=True)
    df['is_processed'] = df['bew_nr']
    df_new['is_processed'].replace(dict_is_processed, inplace=True)
    df_new.to_excel(f'application_data{current_semester}.xlsx')
elif not df_new.empty:              # old_file
    old_id = list(old_df['app_ids'].astype(str))
    df_new = df_new[~df_new['app_ids'].isin(old_id)]
    df_new.drop(columns=['index'], inplace=True)
    df = pd.concat([old_df, df_new], axis=0, sort=False)
    df['is_processed'] = is_processed_app
    df = df.sort_values('uniassist_date', ascending=True)
    df['is_processed'] = df['bew_nr']
    df['is_processed'].replace(dict_is_processed, inplace=True)
    df.to_excel(f'application_data{current_semester}.xlsx')


#download all transcripts that is in df_new
from requests import Session
s = Session()

def _get_downloaded_trans():
    """

    :return: list of pdf_transcript
    """
    transcripts_pdf_paths = glob.glob(transcript_save_location + "/*.pdf")
    exist_transcript = [path[-14:-4] for path in transcripts_pdf_paths]
    return exist_transcript


def pdfDownload(url):
    s.post(url, payload)
    file_id = url[-10:]
    pdf_location = transcript_save_location + f"/{file_id}.pdf"
    response = s.get(url)
    expdf = response.content
    egpdf = open(pdf_location, 'wb')
    egpdf.write(expdf)
    egpdf.close()

if is_downloading_missing_trscripts and missing_trscript_ids:
    print("DOWNLOADING MISSING FILES")
    for individual_url in list_miss_links:
            pdfDownload(individual_url)
            print(f'downloaded 1 file {individual_url[-10:]}')
    print('Finish downloading')

if not df_new.empty:
    count = 0
    existing_trans = _get_downloaded_trans()
    for individual_url in df_new['transcript_link']:
        count += 1
        if individual_url[-10:] not in existing_trans:
            pdfDownload(individual_url)
            print(f'downloaded 1 file {individual_url[-10:]}')
            if count == len(df_new['transcript_link']):
                print('Finish downloading')
        elif individual_url[-10:] in existing_trans:
            print(f"File {individual_url} already downloaded from previous run, skipped")
else:
    print('nothing to download')

#################################
# PUTTING RESULT ONLINE #
#################################
from selenium import webdriver
from selenium.webdriver.common.keys import Keys

import xlrd
df_evaluated = pd.read_excel(f'application_data{current_semester}.xlsx', engine='openpyxl', index_col=0)

df['is_processed'] = df['bew_nr']
df_evaluated['is_processed'].replace(dict_is_processed, inplace=True)
df_upload = df_evaluated[(df_evaluated['is_processed'] == 0) & (df_evaluated['ready_upload'] == 1)]
# list_test = ['197053904', '197054120']df_test = df_evaluated[20:27]
df_upload = df_upload.reset_index(drop=True)

#====================
def option_admit(i):
    switcher = {
        1: '//*[@id="content"]/form/table[2]/tbody/tr[2]/td/input[1]',  # yes admit
        2: '//*[@id="content"]/form/table[2]/tbody/tr[2]/td/input[2]',  # no not admit
        3: '//*[@id="content"]/form/table[2]/tbody/tr[2]/td/input[3]',  # admit with conditions
    }
    return switcher.get(i, "Invalid")

if not df_upload.empty:
    driver = webdriver.Chrome()

    #login using selenium
    driver.get("https://ww2.uni-assist.de/portal/index.php")
    driver.find_element(By.ID, "login").send_keys("Luong")
    driver.find_element(By.ID, "pass").send_keys(payload['pass'])
    driver.find_element(By.XPATH,'//*[@id="content"]/form/div/input[3]').click()
    print("first check")
    page_source = driver.page_source

    for index, admit_true in df_upload['admitted'].iteritems():
        # go to a specific participant page
        applicant_link = df_upload.loc[index, 'bewerten_links']
        print(applicant_link)
        print("second_check")
        driver.get(applicant_link)
        print("third_check")
        if admit_true:
            # select admit in 1st question
            driver.find_element(By.XPATH, '//*[@id="content"]/form/table[2]/tbody/tr[2]/td/input[1]').click()
            driver.find_element(By.XPATH, '//*[@id="content"]/form/div/input').click()

            # admission type in 2nd question
            if df_upload.loc[index, 'admit_type'] == 1:
                admit_decision = option_admit(3)
                driver.find_element(By.XPATH,admit_decision).click()
                driver.find_element(By.XPATH, '//*[@id="content"]/form/div/input').click()

                # specify admit conditions in 3rd question:

                ad_conditions = df_upload.loc[index, 'admit_conditions'].split(',')
                written_conds = ""
                for i in ad_conditions:
                    written_conds += dict_req_courses[i]
                # delete old entry (could improve this, only run this if the previous question has been answered)
                driver.find_element(By.NAME,"memo_c").send_keys(Keys.CONTROL + "a");
                driver.find_element(By.NAME, "memo_c").send_keys(Keys.DELETE);
                # add new entry
                driver.find_element(By.NAME, "memo_c").send_keys(written_conds)
                driver.find_element(By.XPATH, '//*[@id="content"]/form/div/input').click()

                # skip last question, use for specifying why reject, close the case:
                driver.find_element(By.NAME, 'anzeigen_submit').click()

            elif df_upload.loc[index, 'admit_type'] == 0:
                admit_decision = option_admit(1)
                driver.find_element(By.XPATH, admit_decision).click()
                driver.find_element(By.XPATH, '//*[@id="content"]/form/div/input').click()

                # skip the conditions specification since this student is admitted without conditions:
                # delete old entry
                driver.find_element(By.NAME, "memo_c").send_keys(Keys.CONTROL + "a");
                driver.find_element(By.NAME, "memo_c").send_keys(Keys.DELETE);
                # click next
                driver.find_element(By.XPATH, '//*[@id="content"]/form/div/input').click()

                # skip the last question, use for specifying why reject, close the case:
                # delete old entry
                driver.find_element(By.NAME, "memo_c").send_keys(Keys.CONTROL + "a");
                driver.find_element(By.NAME, "memo_c").send_keys(Keys.DELETE);
                # click next
                driver.find_element(By.XPATH, '//*[@id="content"]/form/div/input').click()

            else:
                print('some error must have occured, admitted but unspecified what type of admission (with our without'
                      'conditions')

        elif admit_true == 0:
            # select not admit in the 1st question:
            driver.find_element(By.XPATH, '//*[@id="content"]/form/table[2]/tbody/tr[2]/td/input[2]').click()
            driver.find_element(By.XPATH, '//*[@id="content"]/form/div/input').click()

            # answer 2nd question, not admit:
            admit_decision = option_admit(2)
            driver.find_element(By.XPATH, admit_decision).click()
            driver.find_element(By.XPATH, '//*[@id="content"]/form/div/input').click()

            # skip the 3rd question, conditions when admit
            # delete old entry
            driver.find_element(By.NAME, "memo_c").send_keys(Keys.CONTROL + "a");
            driver.find_element(By.NAME, "memo_c").send_keys(Keys.DELETE);
            # go to next question
            driver.find_element(By.XPATH, '//*[@id="content"]/form/div/input').click()

            # specifying why reject, 4th question, close the case:
            # delete old entry
            driver.find_element(By.NAME, "memo_c").send_keys(Keys.CONTROL + "a");
            driver.find_element(By.NAME, "memo_c").send_keys(Keys.DELETE);
            # add new entry
            reason = df_upload.loc[index, 'reason_reject']
            driver.find_element(By.NAME, "memo_c").send_keys("Ablehnung:\n" + dict_reject[reason])
            driver.find_element(By.XPATH, '//*[@id="content"]/form/div/input').click()



