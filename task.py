

from pyvirtualdisplay import Display
from selenium import webdriver
import time
from bs4 import BeautifulSoup
from openpyxl import Workbook
from selenium.webdriver.support.ui import Select


def main():
    display = Display(visible=0, size=(1024, 768))
    display.start()

    ################# Agency Excel
    browser = webdriver.Firefox(service_log_path='/dev/null')
    browser.get('https://itdashboard.gov/')

    time.sleep(1)

    browser.find_element_by_xpath('//a[@href="#home-dive-in"]').click()
    time.sleep(1)

    agencies_html = browser.page_source

    soup = BeautifulSoup(agencies_html, 'html.parser')
    agencies = soup.find_all(attrs={"class": "h4 w200"})
    spendings = soup.find_all(attrs={"class": "h1 w900"})

    agency_workbook = Workbook()
    dest_filename = 'output/Agencies.xlsx'
    ws1 = agency_workbook.active
    ws1.title = "Agencies"
    ws1['A1'] = "Agencies"
    ws1['B1'] = "Spendings"
    agency_iteration = 2

    for row in range(len(spendings)):
        
        ws1['A'+str(agency_iteration)] = agencies[agency_iteration-2].text
        ws1['B'+str(agency_iteration)] = spendings[agency_iteration-2].text
        print('------ agencies----'+str(agency_iteration))
        agency_iteration = agency_iteration + 1
        

    agency_workbook.save(filename = dest_filename)

    ################# Investment Excel
    agency = soup.find(attrs={"class": "row top-gutter-20"})

    investment_browser = webdriver.Firefox(service_log_path='/dev/null')
    investment_browser.get('https://itdashboard.gov'+str(agency.a['href']))
    time.sleep(10)
    el = investment_browser.find_element_by_id('investments-table-object_length')
    select = Select(el.find_element_by_name('investments-table-object_length'))
    select.select_by_value('-1')

    time.sleep(20)
    agency_page_html = investment_browser.page_source
    investment_soup = BeautifulSoup(agency_page_html, 'html.parser')
    investment_soup.find('table')

    investment_workbook = Workbook()

    investment_dest_filename1 = 'output/Individual_Investments.xlsx'
    ws2 = investment_workbook.active

    ws2['A1'] = investment_soup.find_all('table')[1].find('thead').find_all('tr')[1].find_all('th')[0].text
    ws2['B1'] = investment_soup.find_all('table')[1].find('thead').find_all('tr')[1].find_all('th')[1].text
    ws2['C1'] = investment_soup.find_all('table')[1].find('thead').find_all('tr')[1].find_all('th')[2].text
    ws2['D1'] = investment_soup.find_all('table')[1].find('thead').find_all('tr')[1].find_all('th')[3].text
    ws2['E1'] = investment_soup.find_all('table')[1].find('thead').find_all('tr')[1].find_all('th')[4].text
    ws2['F1'] = investment_soup.find_all('table')[1].find('thead').find_all('tr')[1].find_all('th')[5].text
    ws2['G1'] = investment_soup.find_all('table')[1].find('thead').find_all('tr')[1].find_all('th')[6].text

    investment_iteration = 2
    for el in range(len(investment_soup.find_all('table')[2].find('tbody').find_all('tr'))):

        try:
            
            ws2['A'+str(investment_iteration)] = investment_soup.find_all('table')[2].find('tbody').find_all('tr')[investment_iteration-2].find_all('td')[0].text
            ws2['B'+str(investment_iteration)] = investment_soup.find_all('table')[2].find('tbody').find_all('tr')[investment_iteration-2].find_all('td')[1].text
            ws2['C'+str(investment_iteration)] = investment_soup.find_all('table')[2].find('tbody').find_all('tr')[investment_iteration-2].find_all('td')[2].text
            ws2['D'+str(investment_iteration)] = investment_soup.find_all('table')[2].find('tbody').find_all('tr')[investment_iteration-2].find_all('td')[3].text
            ws2['E'+str(investment_iteration)] = investment_soup.find_all('table')[2].find('tbody').find_all('tr')[investment_iteration-2].find_all('td')[4].text
            ws2['F'+str(investment_iteration)] = investment_soup.find_all('table')[2].find('tbody').find_all('tr')[investment_iteration-2].find_all('td')[5].text
            ws2['G'+str(investment_iteration)] = investment_soup.find_all('table')[2].find('tbody').find_all('tr')[investment_iteration-2].find_all('td')[6].text
            print('------investments----'+str(investment_iteration))
            investment_iteration = investment_iteration + 1
        except:
            continue
        
    investment_workbook.save(filename = investment_dest_filename1)

    #Download PDF

    pdf_profile = webdriver.FirefoxProfile()
    pdf_profile.set_preference("browser.download.dir", 'output/')
    pdf_file_browser = webdriver.Firefox(service_log_path='/dev/null', firefox_profile=pdf_profile)

    pdf_file_browser.get('https://itdashboard.gov'+str(investment_soup.find_all('table')[2].find('tbody').find_all('tr')[0].find_all('td')[0].a['href']))
    time.sleep(10)
    pdf_element = pdf_file_browser.find_element_by_id('business-case-pdf')
    pdf_element.find_element_by_tag_name('a').click()
    time.sleep(5)

    pdf_file_browser.close()
    investment_browser.close()
    browser.close()


if __name__ == "__main__":
    main()