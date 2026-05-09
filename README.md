# PDF Booking Converter

Μετατρέπει PDF κρατήσεων σε Excel ίδιο με το `GenericBookingTemplate.xlsx` για import στο σύστημα.

## Τι κάνει

- Διαβάζει PDF κράτησης.
- Παίρνει στοιχεία όπως Reservation, Voucher, Driver, Stations, Dates, Product Code, Deposit, Amounts, Extras.
- Γράφει αποτέλεσμα στο `GenericBookingTemplate.xlsx`.
- Οι ημερομηνίες γράφονται σαν πραγματικά Excel datetime values.

## Mapping που υπάρχει ήδη

- `Product Code` → `Check-out Notes`
- `Deposit` → `Check-in Notes`
- `Voucher Value + Pay On Arrival` → `Total Amount`
- `Voucher Value` → `Voucher Amount`
- `Naxos Port / Naxos Airport` → `NAXOS`
- `Naxos Port` → `PORT`
- `Naxos Airport` → `Airport`

Το mapping αλλάζει από το αρχείο:

```text
mapping.json
```

## Local χρήση σε PC

```bash
git clone https://github.com/YOUR_USERNAME/pdf-booking-converter.git
cd pdf-booking-converter
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
streamlit run app.py
```

Σε Mac/Linux:

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
streamlit run app.py
```

## Online χρήση με Streamlit Cloud

1. Ανέβασε όλα τα αρχεία σε GitHub repo.
2. Πήγαινε στο https://streamlit.io/cloud
3. New app
4. Διάλεξε το GitHub repo.
5. Main file: `app.py`
6. Deploy

Μετά θα έχεις link τύπου:

```text
https://your-app-name.streamlit.app
```

## Πώς κάνω αλλαγές

Αλλάζεις το `mapping.json` ή το `converter.py` και μετά κάνεις commit στο GitHub.

Παράδειγμα:

```bash
git add .
git commit -m "Update mapping"
git push
```

Το online app ενημερώνεται αυτόματα.

## Σημαντικό

Το app δουλεύει σε PDF που έχουν επιλέξιμο κείμενο. Αν κάποιο PDF είναι σκαναρισμένο σαν εικόνα, θα χρειαστεί OCR.
