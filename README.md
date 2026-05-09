# PDF Booking Converter

Streamlit web app που μετατρέπει **πολλά PDF κρατήσεων** (AbbyCar / Discover Cars format) σε **ένα Excel** τύπου `GenericBookingTemplate.xlsx`, έτοιμο για mass import στο Wheels Car Rental σύστημα.

---

## 📋 Τι κάνει

- **Bulk upload**: Ανεβάζεις 1 ή πολλά PDF μαζί (5-10+ κρατήσεις)
- **Ένα Excel output**: όλες οι κρατήσεις σαν ξεχωριστές σειρές στο ίδιο αρχείο
- **Όνομα output**: `bookings_DD-MM-YYYY.xlsx` (αυτόματα με σημερινή ημερομηνία)
- **Σύνοψη κρατήσεων** σε πίνακα πριν το download
- **Επαλήθευση**: ✅ ή ⚠️ για κάθε κράτηση αν τα extras συμφωνούν με το Pay On Arrival
- Όλα τα extras υπολογίζονται με **+24% ΦΠΑ**
- Οι ημερομηνίες γράφονται ως πραγματικά Excel datetime values

---

## 🗂️ Αρχεία στο project

| Αρχείο | Τι κάνει |
|---|---|
| `app.py` | Το Streamlit web app (UI) |
| `converter.py` | Όλη η λογική του parser |
| `mapping.json` | Σταθερές τιμές + station codes + field mapping |
| `GenericBookingTemplate.xlsx` | Το Excel template (στήλες ADD, BBS, IS, CS, YD) |
| `requirements.txt` | Python dependencies |
| `README.md` | Αυτό το αρχείο |
| `.gitignore` | Αρχεία που αγνοεί το Git |

---

## 🔧 Mapping που υπάρχει ήδη

### Extras (όλα +24% ΦΠΑ)

| PDF | Excel column |
|---|---|
| `Add. Drivers` | `ADD` |
| `Baby Seats` | `BBS` |
| `Infant Seats` | `IS` |
| `Child Seats` + `Booster Seats` + `Todler Seats` | `CS` (συγχωνεύονται) |
| `P.AI.` | `YD` |

**Επαλήθευση:** `ADD + BBS + IS + CS + YD = Pay On Arrival`

### Άλλα fields

- `Product Code` → `Check-out Notes`
- `Deposit` → `Check-in Notes`
- `Voucher Value + Pay On Arrival` → `Total Amount`
- `Voucher Value` → `Voucher Amount`

### Station mapping

| Station στο PDF | Wheels Code | Location |
|---|---|---|
| Mykonos Airport / Mykonos Port | `DRAF` | Airport / PORT |
| Naxos Port / Naxos Airport | `NAXOS` | PORT / Airport |
| Paros Airport / Paros Port | `PAS` | Airport / PORT |

---

## 💻 Local χρήση σε PC

### Windows

```bash
git clone https://github.com/YOUR_USERNAME/pdf-booking-converter.git
cd pdf-booking-converter
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
streamlit run app.py
```

### Mac / Linux

```bash
git clone https://github.com/YOUR_USERNAME/pdf-booking-converter.git
cd pdf-booking-converter
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
streamlit run app.py
```

Μετά άνοιξε στον browser: `http://localhost:8501`

---

## ☁️ Online χρήση μέσω Streamlit Cloud

1. Ανέβασε όλα τα αρχεία σε GitHub repo
2. Πήγαινε στο https://streamlit.io/cloud
3. Πάτα **Create app**
4. Διάλεξε το GitHub repo
5. Main file: `app.py`
6. Πάτα **Deploy**

Θα έχεις link τύπου:
```
https://your-app-name.streamlit.app
```

---

## 🔄 Πώς κάνω αλλαγές

### Αλλαγή στο mapping (πιο εύκολο)

Άνοιξε το `mapping.json` και άλλαξε ό,τι θες:

- **Σταθερές τιμές** (Status, Agent Name, Brand): `fixed_values`
- **Νέα station**: πρόσθεσε στο `station_code_map` και `station_location_map`
- **Νέα στήλη Excel**: πρόσθεσε στο `field_mapping`

### Αλλαγή στη λογική του parser

Άλλαξε το `converter.py` (π.χ. αν αλλάξει το format του PDF).

### Push στο GitHub

```bash
git add .
git commit -m "Update mapping"
git push
```

Το online app ενημερώνεται **αυτόματα** σε ~1 λεπτό.

---

## 💾 Backup

Ο πιο εύκολος τρόπος για backup: **GitHub** κρατάει όλο το ιστορικό αλλαγών. Αρκεί να κάνεις `git push` τακτικά.

Εναλλακτικά, μπορείς να:
- Κάνεις download ZIP από το GitHub repo (Code → Download ZIP)
- Αντιγράψεις τον φάκελο σε Google Drive / Dropbox

---

## ⚠️ Σημαντικό

- Το app δουλεύει **μόνο σε PDF με επιλέξιμο κείμενο**. Σκαναρισμένα PDFs (εικόνες) χρειάζονται OCR.
- Στη σύνοψη, αν μια κράτηση έχει ⚠️ σημαίνει ότι το άθροισμα των extras δεν ταιριάζει με το Pay On Arrival — έλεγξε το πριν κάνεις import.

---

## 📞 Troubleshooting

**Πρόβλημα:** Δεν διαβάζονται οι ημερομηνίες
**Λύση:** Έλεγξε ότι το PDF έχει επιλέξιμο κείμενο (όχι εικόνα)

**Πρόβλημα:** `ADD = 0` αλλά υπάρχει Add. Driver στο PDF
**Λύση:** Έλεγξε στη σύνοψη τη στήλη "Extras Sum". Αν είναι 0, μάλλον το pypdf δεν καταλαβαίνει το format του συγκεκριμένου PDF.

**Πρόβλημα:** `Option: XXX is not existed in Additional Options / Insurances`
**Λύση:** Ο κωδικός στήλης δεν υπάρχει στο σύστημα. Πρόσθεσέ τον στο Wheels ή αφαίρεσέ τον από το `mapping.json` και το template.
