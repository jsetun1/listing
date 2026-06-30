from __future__ import annotations

import streamlit as st
from ua_listing_core import build_files

st.set_page_config(page_title="UA Seasonal Listing Builder", layout="wide")
st.title("UA Seasonal Listing Builder")
st.caption(
    "Kompletní sezónní listing: aktivní standardní UA portfolio a samostatný licenční "
    "portfolio stream od Centric Brands. OOB slouží pouze pro UA objednávky a potvrzené výjimky."
)

with st.sidebar:
    st.header("Nastavení")
    season = st.text_input("Selling Season", value="FW26").strip().upper()
    st.caption(
        "Centric: používají se pouze In-line produkty. Sheet "
        "MFO_Underwear_Undershirts se vždy ignoruje."
    )

st.subheader("Standardní UA podklady")
left, right = st.columns(2)
with left:
    oob_file = st.file_uploader("1. OOB – objednané standardní UA položky", type=["xlsx"], key="oob")
    material_file = st.file_uploader("2. UA Material Data Report", type=["xlsx"], key="material")
    line_list_file = st.file_uploader("3. EMEA Line List", type=["xlsx"], key="line")
with right:
    change_log_file = st.file_uploader("4. EMEA Line List – Change Log", type=["xlsx"], key="changes")
    template_file = st.file_uploader("5. Referenční listing / muster", type=["xlsx"], key="template")
    st.info(
        "Muster určuje cílové sloupce a formát. Používá se také jako reference pro "
        "Size UA → Size EUR a Size Scale standardního UA sortimentu. GHL se přebírá "
        "z neprázdného pole Hero Look Name v Line Listu (ANO / NE)."
    )

st.subheader("Centric Brands – licencované produkty")
centric_left, centric_right, centric_notes = st.columns([1, 1, 1])
with centric_left:
    centric_underwear_file = st.file_uploader(
        "6. Centric Underwear / Undershirts", type=["xlsx"], key="centric_underwear"
    )
with centric_right:
    centric_outerwear_file = st.file_uploader(
        "7. Centric Kids Outerwear – In Line", type=["xlsx"], key="centric_outerwear"
    )
    centric_sportswear_file = st.file_uploader(
        "8. Centric Kids Sportswear – In Line", type=["xlsx"], key="centric_sportswear"
    )
with centric_notes:
    st.info(
        "Centric master data má přednost před UA Material Data. "
        "U kódů typu 25UJFJM07F-001-JPC se do listingu zapisuje "
        "Style 25UJFJM07F a Article 25UJFJM07F-001; přípona -JPC zůstává jen v auditu."
    )

files_ready = all([
    oob_file, material_file, line_list_file, change_log_file, template_file,
    centric_underwear_file, centric_outerwear_file, centric_sportswear_file,
])
if st.button("Vytvořit kompletní sezónní listing", type="primary", disabled=not files_ready):
    with st.spinner("Sestavuji standardní UA portfolio, licenční portfolio Centric a audit…"):
        try:
            listing, audit, summary, issues = build_files(
                oob_file.getvalue(),
                material_file.getvalue(),
                line_list_file.getvalue(),
                change_log_file.getvalue(),
                template_file.getvalue(),
                centric_underwear_file.getvalue(),
                centric_outerwear_file.getvalue(),
                centric_sportswear_file.getvalue(),
                season,
            )
        except Exception as error:
            st.error(str(error))
        else:
            centric_total = (
                summary.get("Centric Underwear In-line EANs", 0)
                + summary.get("Centric Boys Underwear EANs", 0)
                + summary.get("Centric Outerwear In-line EANs", 0)
                + summary.get("Centric Sportswear In-line EANs", 0)
            )
            st.success(f"Hotovo: {summary['Final listing rows']:,} řádků v kompletním sezónním listingu.")
            col1, col2, col3, col4, col5 = st.columns(5)
            col1.metric("Finální řádky", f"{summary['Final listing rows']:,}")
            col2.metric("Standardní UA portfolio", f"{summary['Active portfolio EANs from Line List / Master Data']:,}")
            col3.metric("Centric In-line EANy", f"{centric_total:,}")
            col4.metric("GHL = ANO", f"{summary['GHL = ANO rows']:,}")
            col5.metric("Neuzavřené Size EUR", f"{summary['Unresolved Size EUR rows']:,}")
            st.dataframe(
                [{"Metrika": key, "Hodnota": value} for key, value in summary.items()],
                use_container_width=True,
                hide_index=True,
            )
            st.download_button(
                "Stáhnout kompletní sezónní listing",
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
                    "Zejména jde o EANy chybějící u Centric, Size EUR, odvozené Size Scale a konflikty OOB × DROP."
                )
