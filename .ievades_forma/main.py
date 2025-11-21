import streamlit as st
from datetime import datetime
import time
from supabase import create_client, Client

# Šeit ievietojiet savus Supabase piekļuves parametrus
SUPABASE_URL = "https://uhwbflqdripatfpbbetf.supabase.co"  # Aizstāt ar jūsu Supabase URL
SUPABASE_ANON_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InVod2JmbHFkcmlwYXRmcGJiZXRmIiwicm9sZSI6ImFub24iLCJpYXQiOjE3MzA3MTgxNjMsImV4cCI6MjA0NjI5NDE2M30.PxEpya4E51gCrKa2yNwVRbmK10O-LqQ3uwqNTeXxkck"  # Aizstāt ar jūsu Anon key


# Savienojums ar Supabase
supabase: Client = create_client(SUPABASE_URL, SUPABASE_ANON_KEY)

# Iestatījumi
st.set_page_config(
    page_title="METRUM Vēstuļu draugs",
    layout="centered",
    initial_sidebar_state="collapsed",
)

# Galvenais virsraksts
st.markdown("<h1 style='text-align: center; color: #9C4545;'>METRUM Vēstuļu draugs</h1>", unsafe_allow_html=True)

# Nolasa uzņēmumu sarakstu no Supabase (VD_uzņēmums)
response_company = supabase.table("VD_uzņēmums").select("uzņēmums").execute()
error_company = getattr(response_company, "error", None)
if error_company is not None:
    st.error("Neizdevās ielādēt uzņēmumu sarakstu no datu bāzes.")
    company_options = []
else:
    company_options = [row["uzņēmums"] for row in response_company.data if "uzņēmums" in row]

# Nolasa vietu sarakstu no Supabase (VD_vieta)
response_place = supabase.table("VD_vieta").select('"Sagatavošanas vieta"').execute()
error_place = getattr(response_place, "error", None)
if error_place is not None:
    st.error("Neizdevās ielādēt vietu sarakstu no datu bāzes.")
    place_options = []
else:
    place_options = [row["Sagatavošanas vieta"] for row in response_place.data if "Sagatavošanas vieta" in row]

# Nolasa pagastu/novadu sarakstu no Supabase (VD_pagasts_un_novads)
response_municipality = supabase.table("VD_pagasts_un_novads").select("pagasts_un_novads").execute()
error_municipality = getattr(response_municipality, "error", None)
if error_municipality is not None:
    st.error("Neizdevās ielādēt novadu sarakstu no datu bāzes.")
    municipality_options = []
else:
    municipality_options = [row["pagasts_un_novads"] for row in response_municipality.data if "pagasts_un_novads" in row]

# Nolasa mērnieku sarakstu no Supabase (VD_mērnieks)
response_surveyor = supabase.table("VD_mērnieks").select('"Vārds Uzvārds (sertifikāts Nr.) mērnieka tel. nr."').execute()
error_surveyor = getattr(response_surveyor, "error", None)
if error_surveyor is not None:
    st.error("Neizdevās ielādēt mērnieku sarakstu no datu bāzes.")
    surveyor_options = []
else:
    surveyor_options = [row["Vārds Uzvārds (sertifikāts Nr.) mērnieka tel. nr."] for row in response_surveyor.data if "Vārds Uzvārds (sertifikāts Nr.) mērnieka tel. nr." in row]

# Nolasa sagatavotāja sarakstu no Supabase (VD_sagatavotāja)
response_prepared_by = supabase.table("VD_sagatavotāja").select('"Vārds Uzvārds telefona nr."').execute()
error_prepared_by = getattr(response_prepared_by, "error", None)
if error_prepared_by is not None:
    st.error("Neizdevās ielādēt sagatavotāju sarakstu no datu bāzes.")
    prepared_by_options = []
else:
    prepared_by_options = [row["Vārds Uzvārds telefona nr."] for row in response_prepared_by.data if "Vārds Uzvārds telefona nr." in row]

with st.form("main_form"):
    pdf_file = st.file_uploader("Izvēlieties PDF", type=["pdf"], help="Nav izvēlēta kadastra informācija")
    
    # Uzņēmums tikai no saraksta
    company = st.selectbox("Izvēlieties uzņēmumu:", options=company_options, help="Izvēlieties uzņēmumu no saraksta")

    # Vieta tikai no saraksta
    place = st.selectbox("Izvēlieties vietu:", options=place_options, help="Izvēlieties vietu no saraksta")

    # Pagasts un novads tikai no saraksta
    municipality = st.selectbox("Izvēlieties pagastu un novadu:", options=municipality_options, help="Izvēlieties novadu no saraksta")

    meeting_place = st.text_input("Ievadiet tikšanās vietu un laiku:")
    meeting_date = st.date_input("Ievadiet tikšanās datumu:", datetime.today())

    # Mērnieks tikai no saraksta
    surveyor = st.selectbox("Izvēlieties mērnieku:", options=surveyor_options, help="Izvēlieties mērnieku no saraksta")

    # Sagatavotājs tikai no saraksta
    prepared_by = st.selectbox("Izvēlieties sagatavotāju:", options=prepared_by_options, help="Izvēlieties sagatavotāju no saraksta")

    cols = st.columns(2)
    with cols[0]:
        submitted = st.form_submit_button("Iesniegt")
    with cols[1]:
        canceled = st.form_submit_button("Atcelt")

    if submitted:
        if pdf_file is not None:
            progress_bar = st.progress(0)
            for percent_complete in range(100):
                time.sleep(0.02)
                progress_bar.progress(percent_complete + 1)
            st.success("Faila augšupielāde veiksmīga!")
        else:
            st.warning("Lūdzu, izvēlieties PDF failu.")
