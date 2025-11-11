import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime
import os

# ================= CONFIGURAZIONE SICURA =================
PERCORSO_JSON = "service_account.json"
if not os.path.exists(PERCORSO_JSON):
    print(f"ERRORE: '{PERCORSO_JSON}' NON TROVATO! Mettilo nella cartella.")
    exit()

# Usa variabile d'ambiente per la password!
MITT_EMAIL = "fazzi2007@gmail.com"
MITT_PASSWORD = os.getenv("GMAIL_APP_PASSWORD")
if not MITT_PASSWORD:
    print("ERRORE: Imposta GMAIL_APP_PASSWORD come variabile d'ambiente")
    print("Esempio: export GMAIL_APP_PASSWORD='ryqk gidb gklb yqrz'")
    exit()

OGGETTO_EMAIL = "Promemoria: È tempo di pagare!"

SCOPES = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
URL_SHEET = "https://docs.google.com/spreadsheets/d/17oAuZlV0OM7dcmABKQOypVY1AVIRBgvXkEPLcdLSFhM/edit?gid=0#gid=0"
NOME_FOGGIO = "Foglio1"

# ================= FUNZIONE EMAIL =================
def invia_email(dest_email, nome_utente, importo):
    msg = MIMEMultipart()
    msg['From'] = MITT_EMAIL
    msg['To'] = dest_email
    msg['Subject'] = OGGETTO_EMAIL

    corpo = f"""
    Ciao {nome_utente}!

    È tempo di effettuare il pagamento di **€{importo}**.

    Controlla la tua fattura e procedi entro oggi.

    Grazie!
    Il tuo team
    """

    msg.attach(MIMEText(corpo, 'plain'))

    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(MITT_EMAIL, MITT_PASSWORD)
        server.sendmail(MITT_EMAIL, dest_email, msg.as_string())
        server.quit()
        print(f"Email inviata a {nome_utente} <{dest_email}> - €{importo}")
        return True
    except Exception as e:
        print(f"Errore invio a {dest_email}: {e}")
        return False

# ================= MAIN =================
def main():
    # --- Leggi Google Sheet ---
    try:
        creds = ServiceAccountCredentials.from_json_keyfile_name(PERCORSO_JSON, SCOPES)
        client = gspread.authorize(creds)
        sheet = client.open_by_url(URL_SHEET).worksheet(NOME_FOGGIO)
        dati = sheet.get_all_records()
        df = pd.DataFrame(dati)
        print(f"Google Sheets letto: {len(df)} righe")
    except Exception as e:
        print(f"Errore Google Sheets: {e}")
        return

    oggi = datetime.now().date()
    inviati = 0

    for _, riga in df.iterrows():
        nome = riga.get('Nome', '').strip()
        email = riga.get('Email', '').strip()
        importo = riga.get('Importo', 'N/D')
        data_str = riga.get('Data_Scadenze', '')

        if not nome or not email:
            continue

        # --- Parsing data con dayfirst (Malta = DD/MM/YYYY) ---
        try:
            data_scad = pd.to_datetime(data_str, dayfirst=True).date()
        except:
            print(f"Data non valida per {nome}: '{data_str}'")
            continue

        # --- Controllo scadenza ---
        if data_scad > oggi:
            print(f"{nome}: non ancora scaduto ({data_scad})")
            continue

        # --- INVIO EMAIL (solo se scaduto) ---
        print(f"SCADUTO: {nome} - €{importo} - {data_scad}")
        if invia_email(email, nome, importo):
            inviati += 1

            # Opzionale: aggiungi colonna "Inviato_Il" nel Google Sheet
            # e aggiorna con oggi.strftime('%d/%m/%Y')

    print(f"\nFINE. Email inviate: {inviati}")

if __name__ == "__main__":
    main()