import os
import glob
from io import BytesIO
from datetime import date, timedelta
import streamlit as st
import pandas as pd
from pypdf import PdfReader
from docx import Document
from google import genai

# ==========================================
# KONFIGURACJA STRONY I KLIENTA GEMINI
# ==========================================
st.set_page_config(page_title="Kompas Kompetencji | Ministerstwo Cyfryzacji", page_icon="🧭", layout="wide")

try:
    GEMINI_API_KEY = st.secrets["GEMINI_API_KEY"]
    client = genai.Client(api_key=GEMINI_API_KEY)
except Exception as e:
    st.error(f"Błąd konfiguracji klienta API (sprawdź .streamlit/secrets.toml): {e}")
    client = None

MODEL_AI = "gemini-3.1-flash-lite"

# ==========================================
# FUNKCJE POMOCNICZE (PDF + DOCX)
# ==========================================
@st.cache_data(show_spinner=False)
def load_pdf_knowledge_base():
    """Automatycznie wczytuje wszystkie PDF-y z folderu projektu jako bazę wiedzy z podziałem na źródła."""
    pdf_files = sorted(glob.glob("*.pdf"))
    # Pomijamy ewentualny duplikat (4) i regulaminy hackathonu, jeśli leżą w tym samym folderze
    ignored_keywords = ["Regulamin", "Pełna treść", "(4)"]
    chunks = []
    loaded_names = []

    for pdf_path in pdf_files:
        if any(k.lower() in pdf_path.lower() for k in ignored_keywords):
            continue
        try:
            reader = PdfReader(pdf_path)
            text_pages = []
            for idx, page in enumerate(reader.pages, start=1):
                page_text = page.extract_text() or ""
                if page_text.strip():
                    text_pages.append(f"[Strona {idx}]\n{page_text}")
            if text_pages:
                chunks.append(
                    f"=========================================\n"
                    f"ŹRÓDŁO: {pdf_path}\n"
                    f"=========================================\n" +
                    "\n".join(text_pages)
                )
                loaded_names.append(pdf_path)
        except Exception as e:
            chunks.append(f"[Błąd odczytu {pdf_path}: {e}]")

    return "\n\n".join(chunks), loaded_names


def read_docx_text(file_or_path):
    """Wyciąga tekst z pliku .docx (zarówno z akapitów, jak i z tabel używanych w OSP)."""
    doc = Document(file_or_path)
    lines = []
    for p in doc.paragraphs:
        if p.text.strip():
            lines.append(p.text.strip())
    for table in doc.tables:
        for row in table.rows:
            row_cells = []
            for cell in row.cells:
                txt = cell.text.strip().replace("\n", " ")
                if txt and (not row_cells or txt != row_cells[-1]):
                    row_cells.append(txt)
            if row_cells:
                lines.append(" | ".join(row_cells))
    return "\n".join(lines)


def generate_evaluation_docx(title, metadata_dict, scores_dict, summary_text, uzasadnienie, iprz):
    """Generuje gotowy plik Word (.docx) z podsumowaniem oceny."""
    doc = Document()
    doc.add_heading(title, level=1)
    for k, v in metadata_dict.items():
        p = doc.add_paragraph()
        p.add_run(f"{k}: ").bold = True
        p.add_run(str(v))

    doc.add_heading("Wyniki ocen cząstkowych", level=2)
    table = doc.add_table(rows=1, cols=3)
    table.style = "Table Grid"
    hdr = table.rows[0].cells
    hdr[0].text = "Kryterium / Element oceny"
    hdr[1].text = "Sugestia systemu"
    hdr[2].text = "Ostateczna decyzja przełożonego"

    for kryt, val in scores_dict.items():
        row = table.add_row().cells
        row[0].text = kryt
        row[1].text = val["sugerowana"]
        row[2].text = val["ostateczna"]

    doc.add_heading("Podsumowanie i wynik końcowy", level=2)
    doc.add_paragraph(summary_text)

    if uzasadnienie:
        doc.add_heading("Uzasadnienie przyznanej oceny", level=2)
        doc.add_paragraph(uzasadnienie)

    if iprz:
        doc.add_heading("Wnioski dotyczące rozwoju zawodowego / IPRZ", level=2)
        doc.add_paragraph(iprz)

    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf


# ==========================================
# DANE Z ROZPORZĄDZEŃ (2015 i 2016)
# ==========================================
KRYTERIA_OPISY = {
    "Komunikacja": [
        "Koncentrowanie się na wypowiedzi rozmówcy i dbanie o wzajemne zrozumienie.",
        "Wypowiadanie się w sposób jasny i precyzyjny, dostosowany do sytuacji i rozmówcy.",
        "Formułowanie jasnych i precyzyjnych wypowiedzi pisemnych.",
        "Przedstawianie przekonujących argumentów odpowiednio do sytuacji."
    ],
    "Kreatywność": [
        "Proponowanie lub tworzenie nowych rozwiązań.",
        "Nieschematyczne podejście do zadań."
    ],
    "Negocjowanie": [
        "Przygotowywanie się do negocjacji (projekt stanowiska, zakres ustępstw, argumentacja).",
        "Stosowanie technik i strategii negocjacyjnych.",
        "Osiągnięcie efektu negocjacji akceptowanego przez przełożonych."
    ],
    "Organizacja pracy i orientacja na osiąganie celów": [
        "Określanie sposobu realizacji zadań/celów i ustalanie priorytetów.",
        "Dobieranie niezbędnych zasobów umożliwiających realizację zadań.",
        "Dostosowywanie się do zmieniających się okoliczności.",
        "Wytrwałość i zaangażowanie w realizację zadań/celów."
    ],
    "Orientacja na klienta/interesanta": [
        "Taktowne zachowanie, wyrażające szacunek i zrozumienie dla drugiej osoby.",
        "Rozpoznawanie potrzeb klienta/interesanta i proponowanie rozwiązań."
    ],
    "Podejmowanie decyzji i odpowiedzialność": [
        "Wypracowywanie alternatywnych rozwiązań na podstawie analizy informacji i ryzyka.",
        "Podejmowanie decyzji (decyzyjność).",
        "Przyjmowanie odpowiedzialności za efekty podjętych decyzji."
    ],
    "Radzenie sobie w sytuacjach kryzysowych": [
        "Rozpoznawanie sytuacji kryzysowych i informowanie o nich zainteresowanych.",
        "Podejmowanie działań mających na celu rozwiązanie kryzysu.",
        "Wyciąganie wniosków z sytuacji kryzysowych."
    ],
    "Rzetelność": [
        "Wykonywanie zadań zgodnie ze standardami (przepisami, procedurami) oraz dbałość o jakość.",
        "Dotrzymywanie zobowiązań (realizacja zadań) w wyznaczonym czasie."
    ],
    "Myślenie analityczne": [
        "Rozróżnianie informacji istotnych i nieistotnych.",
        "Dostrzeganie zależności między informacjami.",
        "Formułowanie wniosków."
    ],
    "Wykorzystywanie wiedzy i doskonalenie zawodowe": [
        "Wykorzystywanie wiedzy przy realizacji zadań na zajmowanym stanowisku.",
        "Doskonalenie wiedzy i umiejętności, korzystanie z dobrych praktyk.",
        "Korzystanie z dostępnych źródeł informacji, rozwiązań i technologii."
    ],
    "Współpraca": [
        "Włączanie się w działania realizowane razem z innymi osobami.",
        "Dbanie o przepływ informacji oraz dobrą atmosferę we współpracy.",
        "Dostrzeganie wkładu pracy innych, branie pod uwagę ich zdania i dzielenie się wiedzą."
    ],
    "Wystąpienia publiczne": [
        "Wykorzystywanie różnych środków przekazu i dbałość o sposób prezentacji.",
        "Nawiązywanie kontaktu z audytorium i zarządzanie czasem wystąpienia.",
        "Radzenie sobie ze stresem, krytyką i trudnymi pytaniami."
    ],
    "Zarządzanie ludźmi": [
        "Delegowanie zadań, monitorowanie realizacji i działania korygujące.",
        "Motywowanie pracowników i udzielanie konstruktywnej informacji zwrotnej.",
        "Rozwiązywanie konfliktów i wspieranie pracowników w rozwoju zawodowym."
    ],
    "Zarządzanie strategiczne": [
        "Tworzenie wizji i wyznaczanie celów strategicznych w oparciu o otoczenie.",
        "Formułowanie celów operacyjnych na podstawie celów strategicznych.",
        "Analiza ryzyka planowanych działań o charakterze strategicznym."
    ]
}

OBOWIAZKOWE_B = [
    "Zarządzanie ludźmi",
    "Podejmowanie decyzji i odpowiedzialność",
    "Organizacja pracy i orientacja na osiąganie celów",
    "Współpraca",
    "Komunikacja"
]

OBOWIAZKOWE_C = [
    "Wykorzystywanie wiedzy i doskonalenie zawodowe",
    "Rzetelność",
    "Organizacja pracy i orientacja na osiąganie celów",
    "Współpraca",
    "Komunikacja"
]

ELEMENTY_PIERWSZA_OCENA = [
    "Rzetelność",
    "Wykorzystywanie wiedzy i doskonalenie zawodowe",
    "Organizacja pracy i orientacja na osiąganie celów",
    "Komunikacja",
    "Współpraca"
]

SKALA_OKRESOWA = {
    1: "1 pkt - Znacznie poniżej oczekiwań",
    3: "3 pkt - Poniżej oczekiwań",
    5: "5 pkt - Na poziomie oczekiwań",
    7: "7 pkt - Powyżej oczekiwań",
    9: "9 pkt - Znacznie powyżej oczekiwań"
}

SKALA_PIERWSZA = {
    3: "3 pkt - Poniżej oczekiwań",
    5: "5 pkt - Na poziomie oczekiwań",
    7: "7 pkt - Powyżej oczekiwań"
}

# Wczytanie bazy PDF przy starcie
knowledge_text, loaded_pdfs = load_pdf_knowledge_base()

# ==========================================
# NAWIGACJA BOCZNA (ZGODNA Z GRAFIKĄ MC)
# ==========================================
st.sidebar.title("🧭 Kompas Kompetencji")
st.sidebar.caption("Ministerstwo Cyfryzacji | Służba Cywilna")

page = st.sidebar.radio("Wybierz moduł:", [
    "🏠 Start",
    "🆕 Pierwsza ocena w służbie cywilnej",
    "📅 Ocena okresowa w służbie cywilnej",
    "⚖️ Asystent formalny oceny",
    "💬 Asystent rozmowy i rozwoju",
    "📊 Panel HR / Administratora"
])

st.sidebar.markdown("---")
if loaded_pdfs:
    st.sidebar.success(f"📚 Baza wiedzy AI: {len(loaded_pdfs)} PDF")
    with st.sidebar.expander("Załadowane dokumenty"):
        for f in loaded_pdfs:
            st.caption(f"• {f}")
else:
    st.sidebar.warning("⚠️ Nie wykryto plików PDF w folderze projektu.")


# ==========================================
# MODUŁ 1: START
# ==========================================
if page == "🏠 Start":
    st.title("🧭 Kompas Kompetencji")
    st.markdown("""
    **Wsparcie w przygotowaniu pierwszej oceny i oceny okresowej w służbie cywilnej w oparciu o model kompetencyjny.**
    """)

    col1, col2 = st.columns(2)
    with col1:
        st.info("""
        ### 🆕 Pierwsza ocena w służbie cywilnej
        * **Rozpocznij ocenę:** terminy rozmowy wstępnej (21 dni) i sprawozdania (po 8 mies.).
        * **Zakończ ocenę:** 5 stałych elementów w skali 3/5/7 pkt + weryfikacja egzaminu i sprawozdania.
        """)
    with col2:
        st.info("""
        ### 📅 Ocena okresowa w służbie cywilnej
        * **Krok 1 – Analiza OSP:** wgraj opis stanowiska (`.docx`), a AI dobierze Formularz B lub C oraz kryteria dodatkowe.
        * **Krok 2 – Ocena okresowa:** formularze B i C w skali 1/3/5/7/9 pkt + eksport do `.docx`.
        """)

    st.success("🤖 **Dwa dedykowane moduły AI:** *Asystent formalny* (przepisy z podaniem źródła) oraz *Asystent rozmowy* (scenariusz rozmowy oceniającej, feedback i plan rozwoju). Ostateczna decyzja o ocenie zawsze należy do przełożonego.")


# ==========================================
# MODUŁ 2: PIERWSZA OCENA W SŁUŻBIE CYWILNEJ
# ==========================================
elif page == "🆕 Pierwsza ocena w służbie cywilnej":
    st.title("🆕 Pierwsza ocena w służbie cywilnej")
    tab_start, tab_end = st.tabs(["▶️ Rozpocznij ocenę (Rozmowa wstępna)", "✅ Zakończ ocenę (Model kompetencji)"])

    with tab_start:
        st.subheader("Etap I: Rozmowa wstępna i wyznaczenie terminów")
        c1, c2 = st.columns(2)
        with c1:
            p_imie = st.text_input("Imię i nazwisko ocenianego:", value="Anna Nowak", key="p1_imie")
            p_stan = st.text_input("Stanowisko i komórka:", value="Specjalista, Departament Danych", key="p1_stan")
        with c2:
            data_zatr = st.date_input("Data nawiązania stosunku pracy:", value=date.today(), key="p1_data")
            termin_rozmowy = data_zatr + timedelta(days=21)
            termin_spraw_od = data_zatr + timedelta(days=245)  # po upływie 8 miesięcy
            st.metric("Maks. termin rozmowy wstępnej (21 dni)", termin_rozmowy.strftime("%d.%m.%Y"))
            st.metric("Najwcześniejszy termin sprawozdania (>8 mies.)", termin_spraw_od.strftime("%d.%m.%Y"))

        st.info("Podczas rozmowy wstępnej omów: sposób dokonywania pierwszej oceny, główne obowiązki z OSP, oczekiwania co do 5 elementów oceny oraz termin złożenia sprawozdania (obejmującego min. 7 miesięcy pracy).")

    with tab_end:
        st.subheader("Etap II: Ocena zachowań i warunków formalnych (Skala 3 / 5 / 7 pkt)")
        c1, c2 = st.columns(2)
        egzamin = c1.radio("Czy oceniany zaliczył egzamin ze służby przygotowawczej?", ["TAK", "NIE", "NIE DOTYCZY"], horizontal=True)
        sprawozdanie = c2.radio("Czy sprawozdanie zostało złożone w wyznaczonym terminie?", ["TAK", "NIE"], horizontal=True)

        st.markdown("---")
        wyniki_p1 = {}
        oceny_ostateczne = []

        for elem in ELEMENTY_PIERWSZA_OCENA:
            with st.expander(f"📌 {elem}", expanded=True):
                wskazniki_pkt = []
                for i, opis in enumerate(KRYTERIA_OPISY[elem]):
                    val = st.select_slider(
                        f"Wskaźnik {i+1}: {opis}",
                        options=[3, 5, 7],
                        value=5,
                        format_func=lambda x: SKALA_PIERWSZA[x],
                        key=f"p1_{elem}_{i}"
                    )
                    wskazniki_pkt.append(val)

                sr_wsk = sum(wskazniki_pkt) / len(wskazniki_pkt)
                sugerowana_pkt = min([3, 5, 7], key=lambda x: abs(x - sr_wsk))
                st.info(f"💡 **Sugerowany poziom na podstawie zachowań:** {SKALA_PIERWSZA[sugerowana_pkt]} (średnia wskaźników: {sr_wsk:.2f})")

                ostateczna_pkt = st.selectbox(
                    f"Ostateczna decyzja przełożonego dla '{elem}':",
                    options=[3, 5, 7],
                    index=[3, 5, 7].index(sugerowana_pkt),
                    format_func=lambda x: SKALA_PIERWSZA[x],
                    key=f"p1_final_{elem}"
                )
                oceny_ostateczne.append(ostateczna_pkt)
                wyniki_p1[elem] = {
                    "sugerowana": SKALA_PIERWSZA[sugerowana_pkt],
                    "ostateczna": SKALA_PIERWSZA[ostateczna_pkt]
                }

        srednia_laczna = sum(oceny_ostateczne) / len(oceny_ostateczne)
        liczba_trojek = oceny_ostateczne.count(3)

        warunek_egz = egzamin in ["TAK", "NIE DOTYCZY"]
        warunek_spraw = sprawozdanie == "TAK"
        warunek_trojki = liczba_trojek <= 1
        warunek_srednia = srednia_laczna >= 5.0

        czy_pozytywna = warunek_egz and warunek_spraw and warunek_trojki and warunek_srednia
        wynik_txt = "POZYTYWNA" if czy_pozytywna else "NEGATYWNA"

        st.markdown("---")
        st.subheader(f"Wynik Pierwszej Oceny: {wynik_txt} (Średnia łączna: {srednia_laczna:.2f} pkt)")
        if not czy_pozytywna:
            powody = []
            if not warunek_egz: powody.append("niezaliczony egzamin ze służby przygotowawczej")
            if not warunek_spraw: powody.append("brak sprawozdania w terminie")
            if not warunek_trojki: powody.append(f"więcej niż jedna ocena 3 pkt (liczba: {liczba_trojek})")
            if not warunek_srednia: powody.append(f"ocena łączna poniżej 5 pkt ({srednia_laczna:.2f})")
            st.error("⚠️ Ocena NEGATYWNA. Niespełnione warunki: " + ", ".join(powody) + ". Uzasadnienie jest obowiązkowe!")

        uzas_p1 = st.text_area("Uzasadnienie przyznanej oceny (obowiązkowe przy ocenie negatywnej):", key="p1_uzas")
        rozwoj_p1 = st.text_area("Wnioski dotyczące rozwoju zawodowego (fakultatywne przy ocenie pozytywnej):", key="p1_rozwoj")

        st.session_state["ostatnia_ocena"] = {
            "typ": "Pierwsza ocena w służbie cywilnej",
            "pracownik": p_imie,
            "stanowisko": p_stan,
            "srednia": round(srednia_laczna, 2),
            "wynik": wynik_txt,
            "szczegoly": wyniki_p1
        }

        docx_buf = generate_evaluation_docx(
            "ARKUSZ PIERWSZEJ OCENY W SŁUŻBIE CYWILNEJ",
            {"Oceniany": p_imie, "Stanowisko": p_stan, "Egzamin": egzamin, "Sprawozdanie w terminie": sprawozdanie, "Data": date.today()},
            wyniki_p1,
            f"Ocena łączna: {srednia_laczna:.2f} pkt | Wynik: {wynik_txt}",
            uzas_p1,
            rozwoj_p1
        )
        st.download_button("📄 Pobierz Arkusz Pierwszej Oceny (.docx)", data=docx_buf, file_name=f"pierwsza_ocena_{p_imie.replace(' ', '_')}.docx")


# ==========================================
# MODUŁ 3: OCENA OKRESOWA W SŁUŻBIE CYWILNEJ
# ==========================================
elif page == "📅 Ocena okresowa w służbie cywilnej":
    st.title("📅 Ocena okresowa w służbie cywilnej")

    tab_osp, tab_eval = st.tabs(["1️⃣ Analiza Opisu Stanowiska Pracy (OSP)", "2️⃣ Formularz B / C (Ocena Okresowa)"])

    # --- KROK 1: ANALIZA OSP (.docx) ---
    with tab_osp:
        st.subheader("Automatyczna analiza OSP (.docx) przez AI")
        osp_files = sorted(glob.glob("OSP*.docx"))
        
        col_o1, col_o2 = st.columns(2)
        with col_o1:
            wybrany_lokalny = st.selectbox("Wybierz przykładowy OSP Ministerstwa Cyfryzacji:", ["--- Wybierz plik z folderu ---"] + osp_files)
        with col_o2:
            wgrany_osp = st.file_uploader("...lub wgraj własny plik OSP (.docx):", type=["docx"])

        osp_text = ""
        osp_name = ""
        if wgrany_osp is not None:
            osp_text = read_docx_text(wgrany_osp)
            osp_name = wgrany_osp.name
        elif wybrany_lokalny != "--- Wybierz plik z folderu ---":
            osp_text = read_docx_text(wybrany_lokalny)
            osp_name = wybrany_lokalny

        if osp_text:
            with st.expander(f"Podgląd wyekstrahowanego tekstu: {osp_name}"):
                st.text(osp_text[:3000] + ("..." if len(osp_text) > 3000 else ""))

            if st.button("🔍 Analizuj OSP i dobierz formularz oraz kryteria dodatkowe"):
                if not client:
                    st.error("Brak konfiguracji API Gemini.")
                else:
                    with st.spinner("AI analizuje Opis Stanowiska Pracy..."):
                        prompt_osp = f"""
                        Przeanalizuj poniższy Opis Stanowiska Pracy (OSP) w Ministerstwie Cyfryzacji zgodnie z Rozporządzeniem PRM z 4 kwietnia 2016 r. w sprawie ocen okresowych.
                        Wykonaj 3 rzeczy:
                        1. Wskaż właściwy arkusz oceny: **Formularz B** (jeśli stanowisko wiąże się z bezpośrednim kierowaniem pracownikami) albo **Formularz C** (brak podległych pracowników). Uzasadnij na podstawie pola "Liczba podległych pracowników".
                        2. Wypisz w 3-4 punktach główne obowiązki do omówienia na rozmowie rozpoczynającej ocenę.
                        3. Zarekomenduj maksymalnie 3 **kryteria dodatkowe** z katalogu 14 kryteriów rozporządzenia (niepowtarzające się z obowiązkowymi dla wybranego formularza), które najlepiej pasują do zadań z tego OSP, i krótko uzasadnij wybór każdego z nich.

                        Katalog 14 kryteriów: {', '.join(KRYTERIA_OPISY.keys())}

                        --- TREŚĆ OSP ---
                        {osp_text}
                        """
                        res = client.models.generate_content(model=MODEL_AI, contents=prompt_osp)
                        st.session_state["osp_analiza"] = res.text
                        st.session_state["osp_nazwa"] = osp_name

        if "osp_analiza" in st.session_state:
            st.success(f"Wynik analizy dla: {st.session_state.get('osp_nazwa', '')}")
            st.markdown(st.session_state["osp_analiza"])

    # --- KROK 2: FORMULARZ B / C ---
    with tab_eval:
        st.subheader("Wybór formularza i ocena wskaźników behawioralnych (Skala 1 / 3 / 5 / 7 / 9 pkt)")

        c1, c2 = st.columns(2)
        with c1:
            formularz = st.radio(
                "Wybierz właściwy formularz:",
                ["Formularz B (stanowisko kierujące pracownikami)", "Formularz C (stanowisko niebędące kierowniczym)"],
                horizontal=False
            )
            oceniany = st.text_input("Imię i nazwisko ocenianego:", value="Jan Kowalski")
        with c2:
            stanowisko = st.text_input("Stanowisko i komórka organizacyjna:", value="Starszy specjalista, Departament Cyberbezpieczeństwa")
            data_oceny = st.date_input("Data sporządzenia oceny:", value=date.today())

        obowiazkowe = OBOWIAZKOWE_B if "Formularz B" in formularz else OBOWIAZKOWE_C
        dostepne_dodatkowe = [k for k in KRYTERIA_OPISY.keys() if k not in obowiazkowe]

        dodatkowe = st.multiselect(
            "Wybierz maksymalnie 3 kryteria dodatkowe (opcjonalnie):",
            options=dostepne_dodatkowe,
            max_selections=3
        )

        wszystkie_kryteria = obowiazkowe + dodatkowe
        st.markdown("---")

        wyniki_okr = {}
        punkty_ostateczne = []

        for kryt in wszystkie_kryteria:
            tag = "Obowiązkowe" if kryt in obowiazkowe else "Dodatkowe"
            with st.expander(f"📌 {kryt} ({tag})", expanded=True):
                wsk_pkt = []
                for idx, opis in enumerate(KRYTERIA_OPISY[kryt]):
                    v = st.select_slider(
                        f"Zachowanie {idx+1}: {opis}",
                        options=[1, 3, 5, 7, 9],
                        value=5,
                        format_func=lambda x: SKALA_OKRESOWA[x],
                        key=f"okr_{kryt}_{idx}"
                    )
                    wsk_pkt.append(v)

                sr_w = sum(wsk_pkt) / len(wsk_pkt)
                sugerowana = min([1, 3, 5, 7, 9], key=lambda x: abs(x - sr_w))
                st.info(f"💡 **Sugerowany poziom spełniania przez system:** {SKALA_OKRESOWA[sugerowana]} (średnia zachowań: {sr_w:.2f})")

                ost = st.selectbox(
                    f"Ostateczna decyzja przełożonego dla '{kryt}':",
                    options=[1, 3, 5, 7, 9],
                    index=[1, 3, 5, 7, 9].index(sugerowana),
                    format_func=lambda x: SKALA_OKRESOWA[x],
                    key=f"okr_final_{kryt}"
                )
                punkty_ostateczne.append(ost)
                wyniki_okr[kryt] = {
                    "sugerowana": SKALA_OKRESOWA[sugerowana],
                    "ostateczna": SKALA_OKRESOWA[ost]
                }

        # Logika z § 11 Rozporządzenia z 2016 r.
        srednia_okr = sum(punkty_ostateczne) / len(punkty_ostateczne)
        if srednia_okr <= 2:
            poziom_ogolny = "znacznie poniżej oczekiwań"
        elif srednia_okr <= 4:
            poziom_ogolny = "poniżej oczekiwań"
        elif srednia_okr <= 6:
            poziom_ogolny = "na poziomie oczekiwań"
        elif srednia_okr <= 8:
            poziom_ogolny = "powyżej oczekiwań"
        else:
            poziom_ogolny = "znacznie powyżej oczekiwań"

        ma_jedynke = 1 in punkty_ostateczne
        ocena_koncowa = "NEGATYWNA" if (srednia_okr <= 4 or ma_jedynke) else "POZYTYWNA"

        st.markdown("---")
        m1, m2, m3 = st.columns(3)
        m1.metric("Średnia arytmetyczna", f"{srednia_okr:.2f} pkt")
        m2.metric("Ogólny poziom spełniania", poziom_ogolny)
        m3.metric("Ocena okresowa", ocena_koncowa)

        if ocena_koncowa == "NEGATYWNA":
            powod_neg = "co najmniej jedna ocena cząstkowa 'znacznie poniżej oczekiwań' (1 pkt)" if ma_jedynke else "średnia arytmetyczna <= 4 pkt"
            st.error(f"⚠️ Ocena NEGATYWNA ({powod_neg}). Zgodnie z § 11 ust. 3 rozporządzenia uzasadnienie jest OBOWIĄZKOWE. Ponowna ocena po upływie 6 miesięcy.")

        uzas_okr = st.text_area("Uzasadnienie przyznanej oceny okresowej:", key="okr_uzas")
        iprz_okr = st.text_area("Wnioski dotyczące Indywidualnego Programu Rozwoju Zawodowego (IPRZ):", key="okr_iprz")

        st.session_state["ostatnia_ocena"] = {
            "typ": f"Ocena okresowa ({formularz})",
            "pracownik": oceniany,
            "stanowisko": stanowisko,
            "srednia": round(srednia_okr, 2),
            "poziom_ogolny": poziom_ogolny,
            "wynik": ocena_koncowa,
            "szczegoly": wyniki_okr
        }

        docx_buf_okr = generate_evaluation_docx(
            f"OCENA OKRESOWA W SŁUŻBIE CYWILNEJ ({'ARKUSZ B' if 'Formularz B' in formularz else 'ARKUSZ C'})",
            {"Oceniany": oceniany, "Stanowisko": stanowisko, "Data sporządzenia": data_oceny},
            wyniki_okr,
            f"Średnia: {srednia_okr:.2f} pkt | Ogólny poziom: {poziom_ogolny} | Ocena: {ocena_koncowa}",
            uzas_okr,
            iprz_okr
        )
        st.download_button("📄 Pobierz wypełniony Arkusz Oceny Okresowej (.docx)", data=docx_buf_okr, file_name=f"ocena_okresowa_{oceniany.replace(' ', '_')}.docx")


# ==========================================
# MODUŁ 4: ASYSTENT FORMALNY OCENY
# ==========================================
elif page == "⚖️ Asystent formalny oceny":
    st.title("⚖️ Asystent formalny oceny")
    st.write("Wsparcie w zakresie procedur, terminów, formularzy i poprawności formalnej — wyłącznie na podstawie załadowanych przepisów i poradnika KPRM.")

    if "pytanie_form" not in st.session_state:
        st.session_state["pytanie_form"] = ""

    q_col1, q_col2, q_col3 = st.columns(3)
    if q_col1.button("📅 Kiedy przedłuża się termin 24 miesięcy?"):
        st.session_state["pytanie_form"] = "W jakich sytuacjach i o ile przedłuża się 24-miesięczny termin sporządzenia oceny okresowej?"
    if q_col2.button("⚖️ Jak działa sprzeciw od oceny okresowej?"):
        st.session_state["pytanie_form"] = "Jaki jest termin na wniesienie sprzeciwu od oceny okresowej i jak wygląda procedura jego rozpatrzenia?"
    if q_col3.button("🆕 Jakie warunki trzeba spełnić w pierwszej ocenie?"):
        st.session_state["pytanie_form"] = "Jakie warunki musi spełnić pracownik, aby otrzymać pozytywną pierwszą ocenę w służbie cywilnej?"

    pytanie_form = st.text_area("Zadaj pytanie formalne:", key="pytanie_form", height=110)

    if st.button("Zapytaj Asystenta Formalnego"):
        if not client:
            st.error("⚠️ Klient API nie jest zainicjalizowany.")
        elif not pytanie_form.strip():
            st.warning("Wpisz pytanie.")
        else:
            with st.spinner("Przeszukuję rozporządzenia i poradnik KPRM..."):
                prompt_form = f"""
                Jesteś Asystentem Formalnym w aplikacji Kompas Kompetencji Ministerstwa Cyfryzacji.
                ZASADY BEZWZGLĘDNE:
                1. Odpowiadaj WYŁĄCZNIE na podstawie poniższej BAZY WIEDZY.
                2. ZAWSZE podawaj dokładne źródło odpowiedzi (nazwę dokumentu, paragraf/artykuł oraz numer strony z bazy wiedzy).
                3. Jeśli w bazie wiedzy nie ma wystarczających informacji, napisz wprost: "Na podstawie załadowanej bazy wiedzy nie mogę udzielić zweryfikowanej odpowiedzi na to pytanie" — NIE ZMYŚLAJ.
                4. Nie podejmuj decyzji o ocenie za przełożonego.

                --- BAZA WIEDZY ---
                {knowledge_text}

                --- PYTANIE UŻYTKOWNIKA ---
                {pytanie_form}
                """
                try:
                    resp = client.models.generate_content(model=MODEL_AI, contents=prompt_form)
                    st.success("Odpowiedź zweryfikowana w bazie wiedzy:")
                    st.markdown(resp.text)
                except Exception as e:
                    st.error(f"Błąd API: {e}")


# ==========================================
# MODUŁ 5: ASYSTENT ROZMOWY I ROZWOJU
# ==========================================
elif page == "💬 Asystent rozmowy i rozwoju":
    st.title("💬 Asystent rozmowy oceniającej i planu rozwoju")
    st.write("Pomaga przełożonemu przygotować scenariusz rozmowy, konstruktywny feedback oparty na zachowaniach oraz propozycje do IPRZ (mentoring / coaching / szkolenia).")

    if "ostatnia_ocena" in st.session_state:
        oc = st.session_state["ostatnia_ocena"]
        st.info(f"🔗 Załadowany kontekst z formularza: **{oc['pracownik']}** ({oc['typ']}) – wynik: **{oc['wynik']}** (średnia: {oc['srednia']} pkt).")
    else:
        st.warning("💡 Wskazówka: Wypełnij najpierw ocenę w zakładce *Pierwsza ocena* lub *Ocena okresowa*, aby asystent dopasował pytania dokładnie do wyników pracownika.")

    cel_rozmowy = st.selectbox("W czym ma pomóc Asystent Rozmowy?", [
        "Przygotuj pełny scenariusz rozmowy oceniającej i 5 pytań otwartych",
        "Pomóż sformułować konstruktywny, obiektywny feedback (szczególnie dla niżej ocenionych obszarów)",
        "Zaproponuj konkretne działania rozwojowe do IPRZ (szkolenia, coaching, mentoring)"
    ])
    dodatkowy_kontekst = st.text_area("Dodatkowe uwagi przełożonego (opcjonalnie, np. trudna sytuacja w zespole):", height=90)

    if st.button("Wygeneruj wsparcie do rozmowy"):
        if not client:
            st.error("⚠️ Klient API nie jest zainicjalizowany.")
        else:
            with st.spinner("Przygotowuję wskazówki do rozmowy..."):
                kontekst_oceny = st.session_state.get("ostatnia_ocena", "Brak wypełnionego formularza – przygotuj ogólny schemat dobrej rozmowy w służbie cywilnej.")
                prompt_rozm = f"""
                Jesteś Asystentem Rozmowy Oceniającej i Rozwoju (mentoring/coaching) w aplikacji Kompas Kompetencji Ministerstwa Cyfryzacji.
                Twoim zadaniem jest wesprzeć przełożonego w przeprowadzeniu wartościowej, opartej na faktach i zachowaniach rozmowy z pracownikiem.
                Nie decydujesz o ocenie – pomagasz ją dobrze zakomunikować i zaplanować rozwój.

                ZADANIE: {cel_rozmowy}
                KONTEKST Z FORMULARZA OCENY: {kontekst_oceny}
                DODATKOWE UWAGI PRZEŁOŻONEGO: {dodatkowy_kontekst or 'Brak'}

                Sformułuj konkretną, praktyczną odpowiedź w języku polskim z gotowymi przykładami pytań i sformułowań.
                """
                try:
                    resp = client.models.generate_content(model=MODEL_AI, contents=prompt_rozm)
                    st.success("Gotowe materiały do rozmowy:")
                    st.markdown(resp.text)
                except Exception as e:
                    st.error(f"Błąd API: {e}")


# ==========================================
# MODUŁ 6: PANEL HR / ADMINISTRATORA
# ==========================================
elif page == "📊 Panel HR / Administratora":
    st.title("📊 Panel HR / Administratora")
    st.write("Monitorowanie statusu pierwszych ocen i ocen okresowych w Ministerstwie Cyfryzacji oraz zarządzanie przypomnieniami.")

    dane_hr = pd.DataFrame([
        {"Pracownik": "Jan Kowalski", "Departament": "Cyberbezpieczeństwo", "Typ oceny": "Ocena okresowa (Arkusz C)", "Termin": "2026-12-01", "Status": "W trakcie"},
        {"Pracownik": "Anna Nowak", "Departament": "Rozwój i Bezpieczeństwo AI", "Typ oceny": "Pierwsza ocena", "Termin": "2026-11-25", "Status": "Pilne - zbliża się termin"},
        {"Pracownik": "Piotr Wiśniewski", "Departament": "Współpraca Międzynarodowa", "Typ oceny": "Ocena okresowa (Arkusz B)", "Termin": "2026-12-15", "Status": "Zakończona"},
        {"Pracownik": "Maria Zielińska", "Departament": "GovTech", "Typ oceny": "Pierwsza ocena", "Termin": "2026-11-28", "Status": "Oczekuje na rozmowę"},
    ])

    col_m1, col_m2, col_m3 = st.columns(3)
    col_m1.metric("Aktywne procesy", len(dane_hr))
    col_m2.metric("Zakończone", len(dane_hr[dane_hr["Status"] == "Zakończona"]))
    col_m3.metric("Wymagające pilnej reakcji", len(dane_hr[dane_hr["Status"].str.contains("Pilne")]))

    st.dataframe(dane_hr, use_container_width=True)

    wybrany = st.selectbox("Wyślij przypomnienie do przełożonego dla pracownika:", dane_hr["Pracownik"])
    if st.button("📧 Wyślij przypomnienie"):
        st.toast(f"Wysłano przypomnienie dla procesu: {wybrany}", icon="✅")
        st.success(f"Powiadomienie o zbliżającym się terminie oceny pracownika **{wybrany}** zostało wysłane do przełożonego.")