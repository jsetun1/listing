from __future__ import annotations

import streamlit as st
from ua_listing_core import build_files

st.set_page_config(page_title="UA Seasonal Listing Builder", layout="wide")
st.title("UA Seasonal Listing Builder")
st.caption(
    "Finální zákaznický listing vychází z aktivního Line Listu a Licensed Listu, "
    "rozšířeného o použitelné ADD z Change Logu. OOB pouze označuje již objednané položky "
    "a zachovává potvrzené EANové výjimky."
)

with st.sidebar:
    st.header("Nastavení")
    season = st.text_input("Selling Season", value="FW26").strip().upper()
    st.caption(
        "Produkty, které jsou pouze v Masterdatech, se automaticky nevkládají. "
        "Musí být v aktuálním Line Listu, Licensed Listu nebo být novým ADD s EANy v Masterdatech."
    )

left, right = st.columns(2)
with left:
    oob_file = st.file_uploader("1. OOB (objednané položky)", type=["xlsx"], key="oob")
    material_file = st.file_uploader("2. Material Data Report (EANová master data)", type=["xlsx"], key="material")
    line_list_file = st.file_uploader("3. EMEA Line List", type=["xlsx"], key="line")
with right:
    change_log_file = st.file_uploader("4. EMEA Line List – Change Log", type=["xlsx"], key="changes")
    template_file = st.file_uploader("5. Referenční listing / muster", type=["xlsx"], key="template")
    st.info(
        "Muster určuje cílové sloupce a formát. Používá se také jako reference pro "
        "převod Size UA → Size EUR a Size Scale."
    )

files_ready = all([oob_file, material_file, line_list_file, change_log_file, template_file])
if st.button("Vytvořit aktivní sezónní listing", type="primary", disabled=not files_ready):
    with st.spinner("Sestavuji aktivní portfolio, EANy a kontrolní report…"):
        try:
            listing, audit, summary, issues = build_files(
                oob_file.getvalue(),
                material_file.getvalue(),
                line_list_file.getvalue(),
                change_log_file.getvalue(),
                template_file.getvalue(),
                season,
            )
        except Exception as error:
            st.error(str(error))
        else:
            st.success(f"Hotovo: {summary['Final listing rows']:,} řádků v kompletním aktivním listingu.")
            col1, col2, col3, col4 = st.columns(4)
            col1.metric("Finální řádky", f"{summary['Final listing rows']:,}")
            col2.metric("Aktivní EANy z Line Listu", f"{summary['Active portfolio EANs from Line List / Master Data']:,}")
            col3.metric("OOB výjimky", f"{summary['OOB confirmed exception EANs retained']:,}")
            col4.metric("Neuzavřené Size EUR", f"{summary['Unresolved Size EUR rows']:,}")
            st.dataframe(
                [{"Metrika": key, "Hodnota": value} for key, value in summary.items()],
                use_container_width=True,
                hide_index=True,
            )
            st.download_button(
                "Stáhnout kompletní aktivní listing",
                data=listing,
                file_name=f"{season}_active_listing_data.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
            )
            st.download_button(
                "Stáhnout audit a rozsah portfolia",
                data=audit,
                file_name=f"{season}_active_listing_audit.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
            review = [issue for issue in issues if issue["Severity"] == "Review"]
            if review:
                st.warning(
                    f"{len(review):,} řádků vyžaduje manuální kontrolu. "
                    "Zejména jde o Size EUR, odvozené Size Scale a konflikty OOB × DROP."
                )
