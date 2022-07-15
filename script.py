# İletişim numaralarını kaydetmeden bir excel sayfasından WhatsApp web üzerinden toplu mesaj gönderme programı
# Author @iyunusadas

from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from time import sleep
import pandas

excel_data = pandas.read_excel('Recipients data.xlsx', sheet_name='Recipients')

count = 0

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
driver.get('https://web.whatsapp.com')
input("Whatsapp Web'e giriş yaptıktan sonra ENTER'a basın ve sohbetleriniz görünür hale gelir.")
for column in excel_data['Contact'].tolist():
    try:
        url = 'https://web.whatsapp.com/send?phone=' + str(excel_data['Contact'][count]) + '&text=' + \
              excel_data['Message'][0]
        sent = False
        # Herhangi bir hata oluşması durumunda 3 kez mesaj göndermeye çalışır.
        driver.get(url)
        # Whatsapp güncellemesine binaen send buton class tagi güncelleyiniz.
        try:
            click_btn = WebDriverWait(driver, 35).until(
                EC.element_to_be_clickable((By.CLASS_NAME, 'epia9gcq')))
        except Exception as e:
            print("Üzgünüm mesaj gönderilemedi : " + str(excel_data['Contact'][count]))
        else:
            sleep(2)
            click_btn.click()
            sent = True
            sleep(5)
            print('Mesaj gönderildi: ' + str(excel_data['Contact'][count]))
        count = count + 1
    except Exception as e:
        print('Mesaj gönderilemedi: ' + str(excel_data['Contact'][count]) + str(e))
driver.quit()
print("Komut dosyası başarıyla yürütüldü.")
