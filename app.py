import json
from io import BytesIO

import streamlit as st

from converter import (
    DEFAULT_MAPPING_FILE,
    DEFAULT_TEMPLATE_FILE,
    convert_pdf_to_excel,
)

st.set_page_config(page_title="PDF Booking Converter", page_icon="🚗", layout="centered")

st.title("🚗 PDF Booking Converter")
st.write("Ανεβάζεις PDF κράτησης και κατεβάζεις Excel έτοιμο για import στο σύστημα.")

uploaded_file = st.file_uploader("Ανέβασε PDF κράτησης", type=["pdf"])

with st.expander("Mapping / Ρυθμίσεις", expanded=False):
    st.code(DEFAULT_MAPPING_FILE.read_text(encoding="utf-8"), language="json")
    st.info("Για αλλαγές στο mapping, άνοιξε το αρχείο mapping.json και κάνε edit.")

if uploaded_file:
    pdf_bytes = uploaded_file.read()

    if st.button("Convert to Excel", type="primary"):
        try:
            result = convert_pdf_to_excel(BytesIO(pdf_bytes), DEFAULT_TEMPLATE_FILE, DEFAULT_MAPPING_FILE)
            st.success("Το Excel δημιουργήθηκε επιτυχώς.")

            preview = {k: v for k, v in result.parsed.items() if k != "raw_text"}

            # Highlight whether extras match Pay On Arrival
            if preview.get("extras_match_pay_on_arrival"):
                st.success(
                    f"✅ Extras Total ({preview.get('extras_total')}) = "
                    f"Pay On Arrival ({preview.get('pay_on_arrival')})"
                )
            else:
                st.warning(
                    f"⚠️ Extras Total ({preview.get('extras_total')}) ≠ "
                    f"Pay On Arrival ({preview.get('pay_on_arrival')}) — έλεγξε το PDF"
                )

            st.subheader("Parsed στοιχεία")
            st.json(preview, expanded=False)

            st.download_button(
                label="Download Excel",
                data=result.excel_bytes,
                file_name=result.output_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        except Exception as exc:
            st.error(f"Σφάλμα μετατροπής: {exc}")
            st.caption("Έλεγξε αν το PDF έχει κανονικό κείμενο και όχι σκαναρισμένη εικόνα.")
