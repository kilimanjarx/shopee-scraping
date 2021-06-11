from selenium import webdriver
from selenium.webdriver.support.wait import WebDriverWait
import pandas as pd 
import time
import datetime
import smtplib, ssl
from email.message import EmailMessage

class LazyScraper:

    keyword = ""
    url = "https://shopee.com.my/shocking_sale"

    def read_keyword(self):
        print("Reading keyword..")
        kwd = open("keyword.txt", "r")
        self.keyword = kwd.readline()

    def result_path(self):
        hour = datetime.datetime.now()
        self.filename = 'result_{}.xlsx'.format(hour.strftime("%I"))
        path = r'C:\insert\your\path\shopee-web-scraper\result\{}'.format(self.filename)
        return path
    
    # notify through email
    def send_email(self):
        SUBJECT = "Shopee Scraper Result"
        SENDER_EMAIL = "your email"
        APP_PASSWORD = "your app password"
        RECIPIENT_EMAIL = "your recipient email"
        msg = EmailMessage()
        msg['Subject'] = SUBJECT
        msg['From'] = SENDER_EMAIL
        msg['To'] = RECIPIENT_EMAIL
        msg.set_content("Your scraping result on {}".format(self.keyword))
        context = ssl.create_default_context()

        with open(self.result_path(), 'rb') as f:
            file_data = f.read()
        msg.add_attachment(file_data, maintype="application", subtype="xlsx", filename=self.filename)
        try:
            print("Sending email...")
            smtpserver = smtplib.SMTP("smtp.gmail.com", 587)
            smtpserver.starttls(context=context)
            smtpserver.ehlo()
            smtpserver.login(SENDER_EMAIL, APP_PASSWORD)
            smtpserver.send_message(msg)
            smtpserver.quit()
        except Exception as e:                
            print('Error sending email. Details: {} - {}'.format(e.__class__, e))
        
    def run_search(self,driver):
        self.read_keyword()
        driver.get(self.url)
        WebDriverWait(driver, 20)

        # click on first language selection
        lang = driver.find_elements_by_class_name("language-selection__list-item")
        lang[0].click()
        
        print("Searching for keyword {} at {}...".format(self.keyword,self.url))
        last_height = driver.execute_script("return document.body.scrollHeight")
        while True:
            # Scroll down to bottom
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            # Wait to load page
            time.sleep(2)
            # Calculate new scroll height and compare with last scroll height
            new_height = driver.execute_script("return document.body.scrollHeight")
            if new_height == last_height:
                break
            last_height = new_height

        containers = driver.find_elements_by_xpath('//div[@class="flash-sale-item-card flash-sale-item-card--landing-page flash-sale-item-card--MY"]')
        found = [items for items in containers if "{}".format(self.keyword) in items.find_element_by_xpath('.//div[@class="flash-sale-item-card__item-name-box"]').text]
        listings = []
        for item in found:
            dic = {}
            dic["name"] = item.find_element_by_xpath('.//div[@class="flash-sale-item-card__item-name-box"]').text
            dic["price"] = item.find_element_by_xpath('.//div[@class="flash-sale-item-card__current-price flash-sale-item-card__current-price--landing-page"]').text
            dic["link"] = item.find_element_by_xpath('.//a[@class="flash-sale-item-card-link"]').get_attribute("href")
            listings.append(dic)
        print("Found {} item(s)".format(len(listings)))
        driver.quit()
        return listings

    def main(self):
        options = webdriver.ChromeOptions()
        options.add_argument('--ignore-certificate-errors')
        options.add_argument('--incognito')
        options.add_argument('--headless')
        options.add_experimental_option("excludeSwitches", ["enable-logging"])
        driver = webdriver.Chrome("C:/insert/your/path/shopee-web-scraper/drivers/chromedriver", options=options)
        listings = self.run_search(driver)
        if len(listings) > 0:
            # convert to dataframe and excel 
            df = pd.DataFrame(listings)
            print(df)
            df.to_excel(self.result_path(), index = False)
            self.send_email()
            print("Done!")
        else:
            print("Item not found!")

if __name__ == "__main__":
    LazyScraper().main()
