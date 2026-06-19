"""
byggTotal AI-assistent – ChatGPT-kobling via OpenAI API.
Legg OPENAI_API_KEY i .streamlit/secrets.toml:
    OPENAI_API_KEY = "sk-..."
"""
from __future__ import annotations
import streamlit as st

AI_MODELL  = "gpt-4o-mini"
MAX_HIST   = 20          # maks meldinger i kontekst (spar tokens)
SYSTEM_MSG = """Du er byggTotal-assistenten, en ekspert på norsk byggebransje og konstruksjonsprosjekter.
Du hjelper brukere med:
- IFC-filer, BIM og mengdeuttak
- Konstruksjonsberegninger (stål, betong, tre/limtre/CLT)
- NS-3420, BREEAM, norsk Prisbok og andre norske standarder
- Grunnarbeider, utgraving og geoteknikk
- CO₂-beregninger og EPD-faktorer
- Tolkning av kalkulasjoner, rapporter og tegninger
- Kostnadsestimering og tilbudsarbeid

Svar alltid på norsk. Vær presis og praktisk. Hvis du trenger mer informasjon, spør om det."""


def _hent_klient():
    """Henter OpenAI-klient. Returnerer None og viser melding hvis nøkkel mangler."""
    try:
        from openai import OpenAI
        nøkkel = st.secrets.get("OPENAI_API_KEY", "")
        if not nøkkel or nøkkel.startswith("sk-din"):
            st.warning(
                "⚠️ OpenAI API-nøkkel mangler. "
                "Legg inn nøkkelen i `.streamlit/secrets.toml`:\n\n"
                "```toml\nOPENAI_API_KEY = \"sk-...\"\n```",
                icon="🔑",
            )
            return None
        return OpenAI(api_key=nøkkel)
    except ImportError:
        st.error("Installer openai-pakken: `pip install openai`")
        return None
    except Exception as e:
        st.error(f"Klarte ikke koble til OpenAI: {e}")
        return None


def vis_chatbot(kontekst: str = ""):
    """
    Tegner chatbot-grensesnittet.

    Args:
        kontekst: Valgfri tekststreng med prosjektdata som gis til AI-en
                  slik at svarene blir relevante for det aktive prosjektet.
    """
    st.divider()
    st.subheader("🤖 AI-assistent")
    st.caption("Powered by ChatGPT – spør om IFC, mengder, materialer, kostnader eller konstruksjon.")

    # ── Session state ──────────────────────────────────────────────────────────
    if "chatbot_meldinger" not in st.session_state:
        st.session_state.chatbot_meldinger = []

    col_chat, col_btn = st.columns([6, 1])
    with col_btn:
        if st.button("🗑️ Nullstill", key="chatbot_clear", help="Slett samtalehistorikk"):
            st.session_state.chatbot_meldinger = []
            st.rerun()

    # ── Vis samtalehistorikk ──────────────────────────────────────────────────
    for msg in st.session_state.chatbot_meldinger:
        with st.chat_message(msg["role"]):
            st.markdown(msg["content"])

    # ── Ny melding ────────────────────────────────────────────────────────────
    bruker_input = st.chat_input("Hva kan jeg hjelpe deg med?", key="chatbot_input")
    if not bruker_input:
        return

    # Vis bruker-melding
    with st.chat_message("user"):
        st.markdown(bruker_input)
    st.session_state.chatbot_meldinger.append({"role": "user", "content": bruker_input})

    # ── Send til OpenAI ───────────────────────────────────────────────────────
    klient = _hent_klient()
    if klient is None:
        return

    system_innhold = SYSTEM_MSG
    if kontekst.strip():
        system_innhold += f"\n\nAktuelt prosjekt:\n{kontekst.strip()}"

    meldinger = [{"role": "system", "content": system_innhold}]
    # Begrens historikk
    historikk = st.session_state.chatbot_meldinger[-(MAX_HIST * 2):]
    meldinger += historikk

    try:
        with st.chat_message("assistant"):
            with st.spinner("Tenker…"):
                respons = klient.chat.completions.create(
                    model=AI_MODELL,
                    messages=meldinger,
                    temperature=0.4,
                    max_tokens=1200,
                )
            svar = respons.choices[0].message.content
            st.markdown(svar)
        st.session_state.chatbot_meldinger.append({"role": "assistant", "content": svar})
    except Exception as e:
        st.error(f"Feil fra OpenAI: {e}")
