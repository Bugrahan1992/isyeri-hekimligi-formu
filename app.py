import streamlit as st
import pandas as pd
import openpyxl
from openpyxl import load_workbook
import os
import io
import datetime

# --- Konfigürasyon ve Sabitler ---
TEMPLATE_FILE = 'ek 2 en yeni.xlsx'
SHEET_NAME = 'Sheet1'  # Varsayılan sayfa adı, şablonunuza göre değiştirebilirsiniz

def create_dummy_template_if_not_exists():
    """
    Eğer şablon dosyası yoksa, test amaçlı basit bir Excel oluşturur.
    Böylece uygulama hemen hata vermeden çalışabilir.
    """
    if not os.path.exists(TEMPLATE_FILE):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Sheet1"
        # Başlıklar
        ws['A1'] = "İşe Giriş / Periyodik Muayene Formu Şablonu"
        ws['A4'] = "Adı Soyadı:"
        ws['B4'] = ""  # Buraya veri gelecek
        ws['A5'] = "TC Kimlik No:"
        ws['B5'] = ""
        ws['A6'] = "Doğum Tarihi:"
        ws['B6'] = ""
        # ... Diğer alanlar için yer tutucular ...
        wb.save(TEMPLATE_FILE)
        st.warning(f"UYARI: '{TEMPLATE_FILE}' bulunamadı. Test için geçici bir dosya oluşturuldu.")

def validate_tc(tc_no):
    """Basit TC Kimlik No doğrulaması (11 hane ve sayısal)."""
    if not tc_no:
        return False
    if len(tc_no) != 11:
        return False
    if not tc_no.isdigit():
        return False
    return True

def main():
    st.set_page_config(page_title="İşe Giriş Muayene Formu", page_icon="🏥")

    # Şablon kontrolü
    create_dummy_template_if_not_exists()

    st.title("İşe Giriş / Periyodik Muayene Bilgi Formu")
    st.markdown("Lütfen aşağıdaki bilgileri eksiksiz doldurunuz.")

    with st.form("health_form"):
        # --- Kişisel Bilgiler ---
        st.subheader("Kişisel Bilgiler")
        
        col1, col2 = st.columns(2)
        with col1:
            name = st.text_input("Adı Soyadı")
            tc_no = st.text_input("TC Kimlik No", max_chars=11)
            birth_place = st.text_input("Doğum Yeri")
            birth_date = st.date_input("Doğum Tarihi", min_value=datetime.date(1950, 1, 1))

        with col2:
            gender = st.radio("Cinsiyet", ["Erkek", "Kadın"])
            marital_status = st.radio("Medeni Durumu", ["Evli", "Bekar"])
            education = st.selectbox("Eğitim Durumu", ["İlkokul", "Ortaokul", "Lise", "Üniversite"])
            phone = st.text_input("Telefon Numarası")

        address = st.text_area("Adres")

        # --- Tıbbi Beyan ---
        st.subheader("Tıbbi Beyan")
        st.info("Aşağıdaki durumlardan sizde mevcut olanları işaretleyiniz.")

        medical_questions = {
            "limb_loss": "Uzuv kaybı var mı?",
            "fear_of_heights": "Yükseklik korkusu var mı?",
            "substance_abuse": "Madde bağımlılığı var mı?",
            "epilepsy": "Epilepsi (Sara) hastalığı var mı?"
        }
        
        # Cevapları saklamak için bir sözlük
        answers = {}
        for key, question in medical_questions.items():
            # Radio button ile Evet/Hayır
            answers[key] = st.radio(question, ["Hayır", "Evet"], horizontal=True, key=key)

        submitted = st.form_submit_button("Formu Oluştur ve Hazırla")

        if submitted:
            # Validasyonlar
            if not name:
                st.error("Lütfen Ad Soyad giriniz.")
                return
            if not validate_tc(tc_no):
                st.error("Geçersiz TC Kimlik Numarası. Lütfen 11 haneli numaranızı kontrol ediniz.")
                return

            # --- Excel İşlemleri ---
            try:
                # 1. Şablonu yükle
                wb = load_workbook(TEMPLATE_FILE)
                # Aktif sayfayı veya isme göre sayfayı seçin
                if SHEET_NAME in wb.sheetnames:
                    ws = wb[SHEET_NAME]
                else:
                    ws = wb.active

                # 2. Verileri Hücrelere Yaz
                # NOT: Aşağıdaki hücre adresleri (C5, C7 vs.) tamamen örnektir.
                # Kendi Excel şablonunuzu açıp doğru hücre adreslerini buraya yazmalısınız.
                
                # Örnek Mapping:
                ws['B4'] = name               # Ad Soyad -> B4 (Örnek)
                ws['B5'] = tc_no              # TC No -> B5
                ws['B6'] = birth_date         # Doğum Tarihi -> B6
                ws['B7'] = birth_place
                ws['B8'] = gender
                ws['B9'] = education
                ws['B10'] = marital_status
                ws['B11'] = phone
                ws['B12'] = address

                # Tıbbi sorular (Örneğin alt kısımdaki bir tabloya Evet/Hayır yazıyoruz)
                # Soru 1: Uzuv Kaybı -> B20
                map_medical = {
                    "limb_loss": "B20",
                    "fear_of_heights": "B21",
                    "substance_abuse": "B22",
                    "epilepsy": "B23"
                }

                for key, cell_addr in map_medical.items():
                    ws[cell_addr] = answers[key]

                # 3. Dosyayı belleğe kaydet
                # Disk üzerine yazmak yerine RAM'de (BytesIO) tutuyoruz ki sunucuda dosya kirliliği olmasın
                # ve kullanıcı direkt indirebilsin.
                output = io.BytesIO()
                wb.save(output)
                output.seek(0)
                
                safe_name = "".join([c for c in name if c.isalpha() or c.isdigit() or c==' ']).strip().replace(" ", "_")
                file_name = f"{safe_name}_EK2.xlsx"

                st.success("Form başarıyla oluşturuldu! Aşağıdaki butona tıklayarak indirebilirsiniz.")
                
                st.download_button(
                    label="📥 Doldurulmuş Formu İndir",
                    data=output,
                    file_name=file_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            except Exception as e:
                st.error(f"Excel dosyası oluşturulurken bir hata oluştu: {e}")
                st.error("Lütfen 'ek 2 en yeni.xlsx' dosyasının proje klasöründe olduğundan ve bozuk olmadığından emin olun.")

if __name__ == "__main__":
    main()
