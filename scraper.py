from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import chromedriver_autoinstaller
import pandas as pd
import time
import os
from bs4 import BeautifulSoup
from io import StringIO

# Chromedriver'ı otomatik yükle
chromedriver_autoinstaller.install()

# WebDriver'ı başlat
driver = webdriver.Chrome()

# Web sitesine git
site_url = "https://evam.tuik.gov.tr/dataset/list/"
driver.get(site_url)

# Excel dosyalarını saklamak için klasör oluştur
output_folder = "EVAM_Datasets_2024"
os.makedirs(output_folder, exist_ok=True)

try:
    # tab-12 div'ini bul
    wait = WebDriverWait(driver, 10)
    tab_div = wait.until(EC.presence_of_element_located((By.ID, "tab-12")))

    # Linkleri bul
    links = tab_div.find_elements(By.TAG_NAME, "a")
    print(f"Toplam {len(links)} link bulundu.")

    # Her bir linke tıklama
    for i, link in enumerate(links):
        try:
            # Link URL'sini al
            link_url = link.get_attribute("href")
            print(f"{i+1}. Linke tıklanıyor: {link_url}")

            # Yeni sekmede linki aç
            driver.execute_script("window.open(arguments[0]);", link_url)
            driver.switch_to.window(driver.window_handles[1])

            # Sayfanın yüklenmesini bekle
            time.sleep(10)

            # Sayfanın HTML kaynağını al
            html_source = driver.page_source

            # BeautifulSoup ile HTML'i düzenle
            soup = BeautifulSoup(html_source, "html.parser")

            # Başlığı al (card-header sınıfındaki div'in içeriği)
            header_div = soup.find("div", class_="card-header")
            dataset_title = header_div.text.strip() if header_div else f"veri_seti_{i+1}"

            # Başlıkta geçerli bir metin olduğundan emin olalım
            dataset_title = dataset_title.replace("/", "_").replace("\\", "_")  # Dosya adlarında yasaklı karakterleri temizleyelim

            # İlgili div'i id ile bul
            dataset_div = soup.find("div", id="datasetData")
            
            # Eğer div bulunursa
            if dataset_div:
                # Tabloyu bu div içerisinden çek
                table = dataset_div.find("table")
                if table:
                    # 'colspan' gibi öğeleri temizle
                    for th in table.find_all("th", colspan=True):
                        th.attrs.pop("colspan")  # 'colspan' öğesini temizle
                    
                    # HTML'den tabloyu okuma işlemi
                    table_html = str(table)
                    # StringIO kullanarak html'i pandas'a gönderme
                    table_data = StringIO(table_html)
                    try:
                        df = pd.read_html(table_data)[0]

                        # Tabloyu kaydet
                        file_name = f"{dataset_title}.xlsx"
                        file_path = os.path.join(output_folder, file_name)
                        df.to_excel(file_path, index=False)
                        print(f"Tablo kaydedildi: {file_path}")
                    except Exception as e:
                        print(f"Tabloyu okuma hatası: {e}")
                else:
                    print(f"Tablo bulunamadı: {link_url}")
            else:
                print(f"Dataset verisi bulunamadı: {link_url}")

            # Sekmeyi kapat ve ana sayfaya dön
            driver.close()
            driver.switch_to.window(driver.window_handles[0])

        except Exception as e:
            print(f"Hata oluştu: {e}")
            driver.switch_to.window(driver.window_handles[0])

finally:
    # Tarayıcıyı kapat
    driver.quit()
    print("Tarayıcı kapatıldı.")
