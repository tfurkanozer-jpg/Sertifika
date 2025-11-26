# -*- coding: utf-8 -*-

import pandas as pd
from datetime import datetime
import smtplib
from email.mime.text import MIMEText
from email.header import Header
from email.mime.multipart import MIMEMultipart

# EXCEL dosyasının yolu
excel_path = r"C:\Sertifika\SertifikaTakip.xlsx"  # Gerekirse dosya yolunu değiştir

# SMTP Ayarları
SMTP_SERVER = "smtp.office365.com"
SMTP_PORT = 587
SMTP_USER = "alert@hedefdisticaret.com"  # Gönderen adres
SMTP_PASSWORD = "1dtv5nQJ"               # Uygulama şifresi


# Excel dosyasını oku
df = pd.read_excel(excel_path, engine="openpyxl")

bugun = datetime.today().date()

for _, row in df.iterrows():
    bitis_tarihi = row['Bitiş Tarihi'].date() if not pd.isna(row['Bitiş Tarihi']) else None
    if not bitis_tarihi:
        continue

    kalan_gun = (bitis_tarihi - bugun).days
    mail_adresi = row['Mail Adresi']

    print(f"🕓 Sertifika: {row['İsim']} → Kalan gün: {kalan_gun}")

    if kalan_gun in [30, 15, 7] or (0 <= kalan_gun < 7):
        konu = f"[Uyarı] '{row['İsim']}' süresi dolmak üzere"
        icerik = f"""Merhaba,

'{row['İsim']}' süresi {kalan_gun} gün sonra ({bitis_tarihi}) sona erecek.

- Şirket: {row['Şirket']}
- Tür: {row['Tür']}
- Kurum: {row['Kurum']}
- Adet: {row['Adet']}
- Açıklama: {row['Açıklama']}

Lütfen gerekli aksiyonları alınız."""

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
                print(f"✅ Mail gönderildi → {mail_adresi}")
        except Exception as e:
            print(f"❌ Mail gönderilemedi ({mail_adresi}): {e}")

print("✅ Script başarıyla çalıştı.")
