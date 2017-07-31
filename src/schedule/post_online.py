'''
Created on Jul 30, 2017

@author: venturf2
'''
import os
import time
import yaml
from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException

with open("schedule/creds.yml", 'r') as cred_file:
    creds = yaml.load(cred_file)

os.environ['path'] += r'E:\workspace\Toastmasters\drivers;'
driver = webdriver.Firefox()
# driver.implicitly_wait(10)
driver.get('http://ventanavoices.toastmastersclubs.org/agenda.html')
login_btn = driver.find_element_by_id("memberlogin")
login_btn.click()
user_in = driver.find_element_by_id("memberselect")
user_in.send_keys(creds['user'])
pass_in = driver.find_element_by_id("mpass")
pass_in.send_keys(creds['pass'])
pass_in.send_keys(Keys.RETURN)


def create_agenda(file_path):
    month_agenda = {}
    with open(file_path) as agenda_file:
        meeting = False
        for line in agenda_file:
            if line == '\n':
                continue
            if line.startswith('Meeting Date:'):
                meeting_date = line.split('Meeting Date: ')[1]\
                    .replace('\n', '')
                month_agenda[meeting_date] = {}
                meeting = True
                continue
            if meeting:
                role, member_info = line.split(':')
                member = member_info.split('  ')[0].lstrip()\
                    .replace('\t', '').replace('\n', '').rstrip()
                if member == '':
                    member = 'unassigned'
                # Naming differences b/w Club Scheduler & FTH
                if 'Vote' in role:
                    role = 'ballot counter'
                if 'Topic' in role:
                    role = 'table topics'
                month_agenda[meeting_date][role.lower()] = member.lower()
    return month_agenda


file_path = r'E:\Ventana Drive\Ventana (Work Stuff)\Toastmasters'\
             '\VentanaVoices\VP Ed. scheduling\TMI\ClubScheduler'\
             '\Schedules\Schedule 08-03-2017.txt'
month_agenda = create_agenda(file_path)


select_pos = {}


def create_member_list(select_item):
    print("Creating Members")
    global select_pos
    options = select_item.find_elements_by_tag_name('option')
    for i, opt in enumerate(options):
        if i == 0:
            select_pos['unassigned'] = i
            continue
        member = opt.text
        member = member.split('] ')[1].split(',')[0].lower()
        select_pos[member] = i
    print("Found Members: " + str(select_pos))


for meeting_date, agenda in month_agenda.items():
    # meeting_date should be agenda date in format YYYY-MM-DD
    # agenda should be dictionary with role: member_name
    time.sleep(5)
    driver.find_element_by_link_text("Create New").click()
    time.sleep(1)
    # Creates new agenda
    driver.find_element_by_css_selector("i.fa.fa-plus").click()
    # Select the template
    time.sleep(1)
    Select(driver.find_element_by_id("template")).\
        select_by_visible_text("Template 4: Ventana Voices 3:30-4:30")
    # OK the pop up that comes up
    time.sleep(1)
    driver.find_element_by_xpath("(//button[@type='button'])[42]").click()
    time.sleep(10)
    # Set the date of meeting
#     date_pick = driver.find_element_by_name('meetingdate')
#     date_id = date_pick.get_attribute('id')
#     driver.execute_script("document.getElementById('" + date_id +
#                           "').setAttribute('value', '" + meeting_date + "')")
    time.sleep(1)
    # Save changes
    driver.find_element_by_xpath("(//button[@type='button'])[39]").click()
    time.sleep(5)
    # Change to the meeting role tab
    driver.find_element_by_id("ui-id-15").click()
    time.sleep(5)
    # Find all rows of meeting agenda
    all_rows = driver.\
        find_elements_by_xpath("//table[@id='rostertablema2']/tbody/tr")
#     print("Found rows: " + str(len(all_rows)))
    for row in all_rows:
        # row_id = row.get_attribute('id')
        # Need to find the columns to find the role and select input
        all_cols = row.find_elements_by_xpath(".//td")
#         print("Found cols: " + str(len(all_cols)))
        for col in all_cols:
            try:
                select_item = col.find_element_by_tag_name("select")
                # Only one column should have a select input, that's the one
                my_col = col
                break
            except NoSuchElementException:
                pass
        if len(select_pos) == 0:
            # if not previously created need to populate member list
            create_member_list(select_item)
        # The positions is bolded
        role = my_col.find_element_by_tag_name('b').text.lower()
        print("Found Role: " + role)
        # From the list of members find the select list index
        if "final remarks" in role:
            role = 'toastmaster'
        if role == 'presiding officer':
            member = 'rebecca bowermaster'
        else:
            member = agenda[role]
        sel_index = select_pos[member]
        Select(select_item).select_by_index(sel_index)
    # Save our work
    driver.find_element_by_xpath("(//button[@type='button'])[39]").click()
    time.sleep(5)
    # Close the Agenda Editor
    driver.find_element_by_xpath("(//button[@type='button'])[40]").click()
    time.sleep(5)

driver.quit()


if __name__ == '__main__':
    pass
