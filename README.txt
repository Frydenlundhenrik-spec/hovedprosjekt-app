BYGGAPP – Excel til app

Dette er en Streamlit-prototype som leser Excel-filen og viser:
- mengder
- filtrering på type, materiale og profil
- estimert vekt, volum og kostnad
- 2D-visning av segmenter
- eksport av filtrerte data til CSV

Slik kjører du appen:
1. Installer Python 3.11+ hvis du ikke har det.
2. Åpne terminal i mappen.
3. Kjør:
   pip install -r requirements.txt
4. Start appen:
   streamlit run app.py

Appen bruker example_model.xlsx som standard dersom den ligger i samme mappe.
Du kan også laste opp en annen .xlsx-fil direkte i appen.

Forslag til videre utvikling:
- IFC-import
- NS3420-koder
- CO2-beregning
- PDF-/tilbudseksport
- innlogging og prosjektlagring
