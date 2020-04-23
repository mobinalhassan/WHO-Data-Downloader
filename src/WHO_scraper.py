from telnetlib import EC
from time import sleep
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.ui import Select
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait
from src.countries_list import country_list, years_from, years_to
import os
os.makedirs("../data/", exist_ok=True)
os.makedirs("../data2/", exist_ok=True)
os.makedirs("../data3/", exist_ok=True)


options = Options()
options.add_argument("start-maximized")
options.add_argument("--disable-extensions")
# options.add_argument('--headless')
options.add_argument('window-size=1920x1080')
options.add_argument('--no-sandbox')
options.add_argument("--hide-scrollbars")
options.add_argument("disable-infobars")
options.add_argument('--disable-dev-shm-usage')

prefs = {"profile.default_content_settings.popups": 0,
         "download.default_directory": r"/home/mobin/PycharmProjects/WHO_scraper/data3/",
         # IMPORTANT - ENDING SLASH V IMPORTANT
         "directory_upgrade": True}

options.add_experimental_option('prefs', prefs)


class WhoFileDownloader:
    def __init__(self):
        self.pro_url = 'https://apps.who.int/flumart/Default?ReportNo=12'
        self.pro_links_list = []
        self.driver = webdriver.Chrome(options=options)

    def start(self):
        print('start')
        self.make_selection()

    def set_cookies(self):
        try:
            sleep(5)
            self.driver.implicitly_wait(10)
            cookie = {'name': 'foo', 'value': 'bar'}
            self.driver.add_cookie(cookie)
            print(f"Cookies set ==> {cookie}")
        except Exception as error:
            print(f"Failed in setting cookies ==> {error}")

    def filter_selection(self, country, year_from, year_to, week_from, week_to):
        try:
            # Deselecting All
            sleep(.5)
            select = WebDriverWait(self.driver, 20).until(EC.presence_of_element_located((By.ID, "lstSearchBy")))
            Select(select).deselect_all()
            self.driver.implicitly_wait(10)

            # Selecting Required country
            sleep(.5)
            select = WebDriverWait(self.driver, 20).until(EC.presence_of_element_located((By.ID, "lstSearchBy")))
            Select(select).select_by_visible_text(country)
            print(f'Country selected ==> {country}')
            self.driver.implicitly_wait(10)

            # Selecting Year From
            select = WebDriverWait(self.driver, 20).until(EC.presence_of_element_located((By.ID, "ctl_list_YearFrom")))
            Select(select).select_by_visible_text(year_from)
            print(f'year from selected ==> {year_from}')
            self.driver.implicitly_wait(10)

            # Selecting Week From
            select = WebDriverWait(self.driver, 20).until(EC.presence_of_element_located((By.ID, "ctl_list_WeekFrom")))
            Select(select).select_by_visible_text(week_from)
            print(f'Week from selected ==> {week_from}')
            self.driver.implicitly_wait(10)

            # Selecting Year To
            select = WebDriverWait(self.driver, 20).until(EC.presence_of_element_located((By.ID, "ctl_list_YearTo")))
            Select(select).select_by_visible_text(year_to)
            print(f'year To selected ==> {year_to}')
            self.driver.implicitly_wait(10)

            # Selecting Week To
            select = WebDriverWait(self.driver, 20).until(EC.presence_of_element_located((By.ID, "ctl_list_WeekTo")))
            Select(select).select_by_visible_text(week_to)
            print(f'Week To selected ==> {week_to}')
            self.driver.implicitly_wait(10)

            # clicking report Button
            sleep(1)
            self.driver.find_element_by_css_selector("#ctl_ViewReport").click()
            print('Display Report Btn clicked')

        except Exception as e:
            print('Error in Selection of Finler: ' + str(e))

    def save_record_file(self):

        try:
            sleep(5)
            try:
                WebDriverWait(self.driver, 15).until(EC.invisibility_of_element_located((By.XPATH,
                                                                                         '//*[@id="ctl_ReportViewer_AsyncWait_Wait"]')))
                print(f'Waiting...')
            except TimeoutException:
                print(f"Loader hide")

            sleep(5)
            element = self.driver.find_element_by_css_selector("a#ctl_ReportViewer_ctl05_ctl04_ctl00_ButtonLink")
            self.driver.execute_script("arguments[0].click();", element)
            print('Save File Btn clicked')
            sleep(5)
            element = self.driver.find_element_by_xpath('//a[@title="Excel"]')
            self.driver.execute_script("arguments[0].click();", element)
            print('Excel File select to Download clicked')
        except Exception as e:
            print(f'Error in clicking this path  BTN : ' + str(e))
        sleep(2)

    def make_selection(self):
        try:
            print(self.pro_url)
            self.driver.get(self.pro_url)
            self.set_cookies()
            sleep(1)
            for country in country_list:
                for year_from, year_to in zip(years_from, years_to):
                    week_to = '53'
                    if year_to == '2020':
                        week_to = '14'
                    # print(f'Country ==> {country} Start Year ==> {year_from} End Year ==> {year_to} Start Week ==> {week_from} End week ==> {week_to}')
                    self.filter_selection(country=country, year_from=year_from, year_to=year_to, week_from='1',
                                          week_to=week_to)
                    self.save_record_file()

            self.driver.quit()
        except Exception as error:
            print(f'Quitting form get pro link function ==> {error}')
            self.driver.quit()


def main():
    who_file_downloader = WhoFileDownloader()
    who_file_downloader.start()


if __name__ == "__main__":
    main()
