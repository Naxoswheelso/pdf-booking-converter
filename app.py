from datetime import datetime
from io import BytesIO

import streamlit as st

from converter import (
    DEFAULT_MAPPING_FILE,
    DEFAULT_TEMPLATE_FILE,
    convert_many_pdfs_to_excel,
)

st.set_page_config(page_title="PDF Booking Converter", page_icon="🚗", layout="centered")

st.title("🚗 PDF Booking Converter")
st.write("Ανέβασε ένα ή περισσότερα PDF κρατήσεων και κατέβασε ένα Excel με όλες τις κρατήσεις, έτοιμο για import στο σύστημα.")

uploaded_files = st.file_uploader(
    "Ανέβασε PDF κρατήσεων (μπορείς να επιλέξεις πολλά μαζί)",
    type=["pdf"],
    accept_multiple_files=True,
)

if uploaded_files:
    st.caption(f"📎 {len(uploaded_files)} αρχεία επιλέχθηκαν")

    if st.button("Convert to Excel", type="primary"):
        try:
            # Read all PDF bytes into BytesIO objects
            pdf_inputs = [BytesIO(f.read()) for f in uploaded_files]

            with st.spinner(f"Επεξεργασία {len(pdf_inputs)} κρατήσεων..."):
                excel_bytes, parsed_list = convert_many_pdfs_to_excel(
                    pdf_inputs, DEFAULT_TEMPLATE_FILE, DEFAULT_MAPPING_FILE
                )

            st.success(f"✅ Δημιουργήθηκε Excel με {len(parsed_list)} κρατήσεις.")

            # Build summary table for preview
            summary_rows = []
            mismatch_count = 0
            for i, p in enumerate(parsed_list, start=1):
                match_ok = p.get("extras_match_pay_on_arrival", False)
                if not match_ok:
                    mismatch_count += 1
                summary_rows.append({
                    "#": i,
                    "Reservation": p.get("reservation_number", ""),
                    "Driver": p.get("driver_name", ""),
                    "Pickup": p.get("pickup_datetime").strftime("%d/%m/%Y %H:%M") if p.get("pickup_datetime") else "",
                    "Station": p.get("pickup_station_code", ""),
                    "Total": p.get("total_amount", ""),
                    "Pay On Arrival": p.get("pay_on_arrival", ""),
                    "Extras Sum": p.get("extras_total", ""),
                    "Match": "✅" if match_ok else "⚠️",
                })

            # Show summary table
            st.subheader("Σύνοψη κρατήσεων")
            st.dataframe(summary_rows, use_container_width=True, hide_index=True)

            # Highlight mismatches if any
            if mismatch_count > 0:
                st.warning(
                    f"⚠️ Σε {mismatch_count} από {len(parsed_list)} κρατήσεις, "
                    f"το άθροισμα των extras δεν συμφωνεί με το Pay On Arrival. "
                    f"Έλεγξε τις γραμμές με ⚠️."
                )

            # Filename: bookings_DD-MM-YYYY.xlsx
            today_str = datetime.now().strftime("%d-%m-%Y")
            filename = f"bookings_{today_str}.xlsx"

            st.download_button(
                label=f"📥 Download {filename}",
                data=excel_bytes,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
            )

        except Exception as exc:
            st.error(f"Σφάλμα μετατροπής: {exc}")
            st.caption("Έλεγξε αν τα PDF έχουν κανονικό κείμενο και όχι σκαναρισμένη εικόνα.")
