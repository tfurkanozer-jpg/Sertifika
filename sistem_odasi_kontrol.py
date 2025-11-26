# -*- coding: utf-8 -*-

import pandas as pd
from datetime import datetime
import smtplib
import json
from email.mime.text import MIMEText
from email.header import Header
from email.mime.multipart import MIMEMultipart

# EXCEL dosyasının yolu
excel_path = r"C:\Sertifika\sistem_odasi_kontrol.xlsx"

# SMTP Ayarları
SMTP_SERVER = "smtp.office365.com"
SMTP_PORT = 587
SMTP_USER = "alert@hedefdisticaret.com"
SMTP_PASSWORD = "1dtv5nQJ"

# Excel dosyasını oku
df = pd.read_excel(excel_path, engine="openpyxl")

bugun = datetime.today().date()
gun_adi = datetime.today().strftime('%A')

# Türkçe gün isimleri
gun_map = {
    'Monday': 'Pazartesi',
    'Tuesday': 'Salı',
    'Wednesday': 'Çarşamba',
    'Thursday': 'Perşembe',
    'Friday': 'Cuma',
    'Saturday': 'Cumartesi',
    'Sunday': 'Pazar'
}

bugun_tr = gun_map.get(gun_adi)

for _, row in df.iterrows():
    mail_adresi = row['Mail Adresi']

    try:
        icerik_json = json.loads(row['GünlükKonuVeIçerik'])
    except Exception as e:
        print(f"⚠️ JSON hatası (satır: {row['İsim']}): {e}")
        continue

    bugune_ait_mailler = icerik_json.get(bugun_tr, [])

    for mail in bugune_ait_mailler:
        try:
            konu, icerik = mail.split("|", 1)
        except ValueError:
            print(f"⚠️ Hatalı içerik formatı → {mail}")
            continue

        msg = MIMEMultipart()
        msg["From"] = SMTP_USER
        msg["To"] = mail_adresi
        msg["Subject"] = Header(konu.encode('utf-8'), 'utf-8').encode()
        msg.attach(MIMEText(icerik, 'plain', 'utf-8'))

        try:
            with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
                server.starttls()
                server.login(SMTP_USER, SMTP_PASSWORD)
                server.sendmail(SMTP_USER, mail_adresi, msg.as_string())
                print(f"✅ [{bugun_tr}] Mail gönderildi → {mail_adresi}")
        except Exception as e:
            print(f"❌ Mail gönderilemedi ({mail_adresi}): {e}")

print("✅ Script başarıyla çalıştı.")
