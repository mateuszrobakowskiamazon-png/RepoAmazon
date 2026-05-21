"""
Amazon DE - Generator Tytułów (Streamlit)
Łapie keywordy z Cerebro (Helium10), tytuły konkurencji, parametry produktu
i generuje 3 wersje tytułu + plik Excel z analizą.
"""

import re
from collections import Counter
from io import BytesIO

import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side


# ===================== STAŁE =====================

EXEMPT_WORDS = {
    "und", "oder", "der", "die", "das", "für", "mit", "auf", "in", "zu",
    "von", "aus", "bei", "am", "im", "als", "set", "cm", "er", "&",
    "ein", "eine", "einen", "einer", "eines",
}

FORBIDDEN_CHARS = "!$?_{}^¬¦|"
TITLE_LIMIT = 200

# Cechy → ich formy przymiotnikowe dla wersji zdaniowej
FEATURE_ADJ = {
    "waschbar": "waschbare",
    "rutschfest": "rutschfeste",
    "wasserdicht": "wasserdichte",
    "wetterfest": "wetterfeste",
    "abnehmbar": "abnehmbare",
    "gesteppt": "gesteppte",
    "atmungsaktiv": "atmungsaktive",
    "uv-beständig": "UV-beständige",
}


# ===================== PARSOWANIE =====================

def parse_cerebro(uploaded_file) -> pd.DataFrame:
    """Czyta plik Cerebro (Helium10) i zwraca top keywordy posortowane po SV."""
    df = pd.read_excel(uploaded_file)
    expected = ["Keyword Phrase", "Search Volume", "Keyword Sales"]
    if not all(c in df.columns for c in expected):
        raise ValueError(
            f"Plik nie wygląda na export z Cerebro - brak kolumn: {expected}"
        )
    return (
        df[expected + [c for c in df.columns if c.startswith("B0")]]
        .sort_values("Search Volume", ascending=False)
        .reset_index(drop=True)
    )


def parse_competitor_titles(text: str) -> list[str]:
    """Lista tytułów konkurencji - jeden per linia."""
    return [line.strip() for line in text.splitlines() if line.strip()]


def extract_keywords_from_titles(titles: list[str]) -> list[tuple[str, int]]:
    """Wyciąga najczęstsze słowa z tytułów konkurencji."""
    all_words = []
    for t in titles:
        words = re.findall(r"[A-Za-zÄÖÜäöüß]+", t.lower())
        all_words.extend(w for w in words if w not in EXEMPT_WORDS and len(w) > 3)
    return Counter(all_words).most_common(20)


# ===================== WALIDACJA =====================

def validate_title(title: str, no_commas: bool = True) -> dict:
    """Sprawdza zgodność z regułami Amazon DE od 21.01.2025."""
    issues = []

    if len(title) > TITLE_LIMIT:
        issues.append(f"Za długi: {len(title)}/{TITLE_LIMIT} znaków")

    bad = [c for c in FORBIDDEN_CHARS if c in title]
    if bad:
        issues.append(f"Zakazane znaki: {' '.join(bad)}")

    if no_commas and "," in title:
        issues.append("Zawiera przecinki (wyłączone w ustawieniach)")

    words = re.findall(r"[A-Za-zÄÖÜäöüß]+", title.lower())
    counts = Counter(w for w in words if w not in EXEMPT_WORDS and len(w) > 1)
    repeated = {w: c for w, c in counts.items() if c > 2}
    if repeated:
        issues.append(f"Powtórzenia 3+ razy: {repeated}")

    return {
        "length": len(title),
        "valid": not issues,
        "issues": issues,
    }


# ===================== GENERATORY TYTUŁÓW =====================

def join_parts(parts: list[str], sep: str = " ") -> str:
    return sep.join(p for p in parts if p)


def build_title(
    style: str,
    brand: str,
    main_kw: str,
    size: str,
    set_size: str,
    material: str,
    features: list[str],
    synonyms: list[str],
    contexts: list[str],
    color: str,
    no_commas: bool,
) -> str:
    """style: 'A' (zlepek), 'B' (hybryda), 'C' (zdaniowa)"""

    dash = " – "  # em dash
    amp = " & "
    sep = " " if no_commas else ", "

    # zabezpiecz: brak duplikatów w synonimach i kontekstach
    syns = list(dict.fromkeys(s.strip() for s in synonyms if s.strip()))
    ctxs = list(dict.fromkeys(c.strip() for c in contexts if c.strip()))
    feats = [f.strip() for f in features if f.strip()]

    if style == "A":
        # Klasyczny zlepek (jak top konkurencji)
        parts = [
            brand,
            main_kw,
            size,
            set_size,
            material,
            dash,
            *syns,
            *feats,
            "für",
            *ctxs,
        ]
        title = " ".join(p for p in parts if p)
        if color:
            title += dash + color
        return title

    if style == "B":
        # Hybryda - bardziej naturalne, z & jako separatorem
        feat_str = " ".join(feats)
        syn_block = amp.join(syns[:2]) + " " + " ".join(syns[2:]) if len(syns) > 2 else amp.join(syns)
        ctx_block = " ".join(ctxs[:-1]) + amp + ctxs[-1] if len(ctxs) > 1 else (ctxs[0] if ctxs else "")

        parts = [
            brand,
            main_kw,
            size,
            set_size,
            "aus",
            material,
            feat_str,
            dash,
            syn_block,
            "für",
            ctx_block,
        ]
        title = " ".join(p for p in parts if p).replace("  ", " ")
        if color:
            title += dash + color
        return title

    if style == "C":
        # Zdaniowa - najczystsza, zgodna z duchem nowych reguł Amazon
        adj_form = FEATURE_ADJ.get(feats[0].lower(), feats[0] + "e") if feats else ""
        main_syn = syns[0] if syns else main_kw
        other_syns = syns[1:]
        ctx_str = " ".join(ctxs[:-1]) + " und " + ctxs[-1] if len(ctxs) > 1 else (ctxs[0] if ctxs else "")

        parts = [
            brand,
            main_kw,
            size,
            "aus",
            material,
            set_size,
            dash,
            f"{adj_form} {main_syn}".strip(),
            "für",
            ctx_str,
        ]
        if other_syns:
            parts.extend([dash, amp.join(other_syns[:2])])

        title = " ".join(p for p in parts if p).replace("  ", " ")
        if color:
            title += dash + color
        return title

    raise ValueError(f"Nieznany styl: {style}")


def trim_to_fit(title_fn, max_len: int = TITLE_LIMIT, **kwargs) -> str:
    """Jeśli tytuł za długi, stopniowo wycina konteksty/synonimy."""
    title = title_fn(**kwargs)
    if len(title) <= max_len:
        return title

    # próbuj wycinać konteksty od końca
    contexts = list(kwargs.get("contexts", []))
    while contexts and len(title) > max_len:
        contexts.pop()
        kwargs["contexts"] = contexts
        title = title_fn(**kwargs)

    # potem synonimy
    syns = list(kwargs.get("synonyms", []))
    while len(syns) > 2 and len(title) > max_len:
        syns.pop()
        kwargs["synonyms"] = syns
        title = title_fn(**kwargs)

    return title


# ===================== EXCEL OUTPUT =====================

def build_excel(
    titles: dict[str, dict],
    cerebro_df: pd.DataFrame,
    competitor_titles: list[str],
    backend_terms: str,
    product_params: dict,
) -> bytes:
    """Buduje plik Excel z 5 arkuszami i zwraca jako bajty."""
    wb = Workbook()

    # === Arkusz 1: Opcje tytułów ===
    ws = wb.active
    ws.title = "Opcje tytułów"
    headers = ["Opcja", "Tytuł", "Długość", "Walidacja", "Problemy"]
    for col, h in enumerate(headers, 1):
        c = ws.cell(row=1, column=col, value=h)
        c.font = Font(bold=True, color="FFFFFF", name="Arial")
        c.fill = PatternFill("solid", start_color="2C3E50")
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    style_descriptions = {
        "A": "A – Klasyczny zlepek",
        "B": "B – Hybryda (REKOMENDOWANE)",
        "C": "C – Zdaniowa naturalna",
    }
    fills = {"A": "FFF3CD", "B": "D4EDDA", "C": "D1ECF1"}

    thin = Side(border_style="thin", color="CCCCCC")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for r_idx, style in enumerate(["A", "B", "C"], 2):
        data = titles[style]
        ws.cell(row=r_idx, column=1, value=style_descriptions[style])
        ws.cell(row=r_idx, column=2, value=data["title"])
        ws.cell(row=r_idx, column=3, value=f"{data['length']} zn.")
        ws.cell(row=r_idx, column=4, value="✅ OK" if data["valid"] else "❌ Problem")
        ws.cell(row=r_idx, column=5, value="\n".join(data["issues"]) if data["issues"] else "")
        for col in range(1, 6):
            cell = ws.cell(row=r_idx, column=col)
            cell.font = Font(name="Arial", size=10, bold=(col == 1))
            cell.fill = PatternFill("solid", start_color=fills[style])
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            cell.border = border

    ws.column_dimensions["A"].width = 26
    ws.column_dimensions["B"].width = 70
    ws.column_dimensions["C"].width = 12
    ws.column_dimensions["D"].width = 14
    ws.column_dimensions["E"].width = 35
    for r in range(2, 5):
        ws.row_dimensions[r].height = 90

    # === Arkusz 2: Top słowa kluczowe ===
    ws2 = wb.create_sheet("Top słowa kluczowe")
    kw_headers = ["#", "Słowo kluczowe", "Search Volume", "Keyword Sales"]
    for col, h in enumerate(kw_headers, 1):
        c = ws2.cell(row=1, column=col, value=h)
        c.font = Font(bold=True, color="FFFFFF", name="Arial")
        c.fill = PatternFill("solid", start_color="2C3E50")
        c.alignment = Alignment(horizontal="center", vertical="center")

    top_kws = cerebro_df.head(20)
    for r_idx, (_, row) in enumerate(top_kws.iterrows(), 2):
        ws2.cell(row=r_idx, column=1, value=r_idx - 1)
        ws2.cell(row=r_idx, column=2, value=row["Keyword Phrase"])
        ws2.cell(row=r_idx, column=3, value=int(row["Search Volume"]) if pd.notna(row["Search Volume"]) else "-")
        ws2.cell(row=r_idx, column=4, value=int(row["Keyword Sales"]) if pd.notna(row["Keyword Sales"]) else "-")
        for col in range(1, 5):
            cell = ws2.cell(row=r_idx, column=col)
            cell.font = Font(name="Arial", size=10)
            cell.alignment = Alignment(vertical="center")
            cell.border = border
            if r_idx <= 4:
                cell.fill = PatternFill("solid", start_color="FFEAA7")

    ws2.column_dimensions["A"].width = 5
    ws2.column_dimensions["B"].width = 40
    ws2.column_dimensions["C"].width = 15
    ws2.column_dimensions["D"].width = 15

    # === Arkusz 3: Backend Search Terms ===
    ws3 = wb.create_sheet("Backend Search Terms")
    ws3["A1"] = "Backend Search Terms (Seller Central → Schlüsselwörter)"
    ws3["A1"].font = Font(bold=True, size=12, name="Arial")
    ws3["A2"] = "Wklej do pola „Allgemeine Schlüsselwörter” w Seller Central. Limit Amazon DE: 249 znaków."
    ws3["A2"].font = Font(italic=True, size=9, color="666666", name="Arial")
    ws3["A2"].alignment = Alignment(wrap_text=True)
    ws3["A4"] = "Sugerowane słowa kluczowe do backendu:"
    ws3["A4"].font = Font(bold=True, name="Arial")
    ws3["A5"] = backend_terms
    ws3["A5"].font = Font(name="Arial", size=10)
    ws3["A5"].alignment = Alignment(wrap_text=True, vertical="top")
    ws3["A5"].fill = PatternFill("solid", start_color="F0F0F0")
    ws3["A6"] = f"Długość: {len(backend_terms)} znaków / 249"
    ws3["A6"].font = Font(italic=True, size=9, color="666666", name="Arial")
    ws3.column_dimensions["A"].width = 90
    ws3.row_dimensions[5].height = 60

    # === Arkusz 4: Konkurencja ===
    ws4 = wb.create_sheet("Konkurencja")
    ws4["A1"] = "Tytuły konkurencji (wklejone)"
    ws4["A1"].font = Font(bold=True, size=12, name="Arial")
    for i, t in enumerate(competitor_titles, 3):
        ws4.cell(row=i, column=1, value=f"#{i-2}")
        ws4.cell(row=i, column=2, value=t)
        ws4.cell(row=i, column=3, value=f"{len(t)} zn.")
        for col in range(1, 4):
            ws4.cell(row=i, column=col).font = Font(name="Arial", size=10)
            ws4.cell(row=i, column=col).alignment = Alignment(wrap_text=True, vertical="top")
    ws4.column_dimensions["A"].width = 5
    ws4.column_dimensions["B"].width = 90
    ws4.column_dimensions["C"].width = 12

    # === Arkusz 5: Parametry użyte ===
    ws5 = wb.create_sheet("Parametry produktu")
    ws5["A1"] = "Parametry użyte do generowania tytułów"
    ws5["A1"].font = Font(bold=True, size=12, name="Arial")
    for i, (key, val) in enumerate(product_params.items(), 3):
        ws5.cell(row=i, column=1, value=key).font = Font(name="Arial", size=10, bold=True)
        ws5.cell(row=i, column=2, value=str(val)).font = Font(name="Arial", size=10)
    ws5.column_dimensions["A"].width = 25
    ws5.column_dimensions["B"].width = 70

    # Zapis do bajtów
    buffer = BytesIO()
    wb.save(buffer)
    return buffer.getvalue()


# ===================== STREAMLIT UI =====================

st.set_page_config(
    page_title="Amazon DE - Generator Tytułów",
    page_icon="📝",
    layout="wide",
)

st.title("📝 Amazon DE – Generator Tytułów")
st.caption(
    "Wrzuć plik z Cerebro (Helium10), wklej tytuły konkurencji, "
    "wpisz parametry produktu → wyjście to 3 wersje tytułu + Excel z analizą."
)

# === Sidebar: parametry produktu ===
with st.sidebar:
    st.header("🏷️ Parametry produktu")

    brand = st.text_input("Marka", placeholder="np. BIELIK", help="Pierwsze słowo tytułu")
    main_kw = st.text_input(
        "Główna fraza produktu",
        placeholder="np. Sitzkissen Rund",
        help="Najsilniejszy keyword - powinien być pierwszy po marce",
    )
    size = st.text_input("Wymiar", placeholder="np. 40 cm albo 120x80")
    set_size = st.selectbox(
        "Set / ilość",
        ["1 Stück", "2er Set", "3er Set", "4er Set", "6er Set", "8er Set"],
        index=3,
    )
    material = st.text_input("Materiał", placeholder="np. Cord, Filz, Baumwolle")

    st.markdown("---")
    st.subheader("Cechy produktu")
    features_text = st.text_area(
        "Cechy (jedna per linia)",
        placeholder="waschbar\nrutschfest\nabnehmbar",
        height=100,
    )

    st.subheader("Synonimy / typy produktu")
    synonyms_text = st.text_area(
        "Synonimy (jedna per linia)",
        placeholder="Stuhlkissen\nBodenkissen\nSitzpolster\nStuhlauflage",
        height=120,
        help="Skopiuj z tytułów konkurencji - jakie inne nazwy używają dla tego produktu",
    )

    st.subheader("Konteksty użycia")
    contexts_text = st.text_area(
        "Konteksty (jeden per linia)",
        placeholder="Esszimmer\nWohnzimmer\nKüche\nSitzbank\nHocker",
        height=120,
        help="Gdzie/do czego klienci wyszukują ten produkt",
    )

    st.markdown("---")
    color = st.text_input(
        "Kolor (opcjonalnie)",
        placeholder="np. Anthrazit",
        help="Pomiń jeśli to PARENT-tytuł wariantu kolorystycznego",
    )

    st.markdown("---")
    no_commas = st.toggle("Bez przecinków w tytule", value=True)
    is_parent = st.toggle(
        "To jest PARENT-tytuł (wariant kolorystyczny)",
        value=False,
        help="Parent NIE może mieć koloru w tytule",
    )

# === Główne pole ===
col1, col2 = st.columns([1, 1])

with col1:
    st.subheader("1. Plik Cerebro (Helium10)")
    cerebro_file = st.file_uploader(
        "Wgraj plik .xlsx z eksportu Cerebro",
        type=["xlsx"],
        help="Plik z keywordami, Search Volume, Keyword Sales",
    )

with col2:
    st.subheader("2. Tytuły konkurencji")
    competitor_text = st.text_area(
        "Wklej tytuły konkurencji (jeden per linia)",
        placeholder=(
            "Lsjoaw Sitzkissen Rund 35cm Dicke 6cm 1 Stück Bodenkissen...\n"
            "heimtexland® Sitzkissen Set Filz Stuhlkissen Type 631 35 cm Orange...\n"
            "..."
        ),
        height=150,
    )

# Backend keywords
st.subheader("3. (Opcjonalnie) Słowa kluczowe do backendu")
backend_text = st.text_area(
    "Słowa do pola Schlüsselwörter (do 249 zn.) - nie wpisuj słów, które są już w tytule",
    placeholder="kissen rund cordstoff sitzauflage rutschfest abnehmbar terrasse balkon...",
    height=80,
)

# === Generowanie ===
st.markdown("---")

if st.button("🚀 Generuj tytuły", type="primary", use_container_width=True):
    # Walidacja wejścia
    if not brand or not main_kw:
        st.error("❌ Marka i Główna fraza są wymagane.")
        st.stop()

    if not cerebro_file:
        st.warning("⚠️ Brak pliku Cerebro - generuję bez analizy keywordów (Excel będzie pusty w arkuszu top).")
        cerebro_df = pd.DataFrame(columns=["Keyword Phrase", "Search Volume", "Keyword Sales"])
    else:
        try:
            cerebro_df = parse_cerebro(cerebro_file)
        except ValueError as e:
            st.error(f"❌ {e}")
            st.stop()

    # Parsowanie
    features = [f.strip() for f in features_text.split("\n") if f.strip()]
    synonyms = [s.strip() for s in synonyms_text.split("\n") if s.strip()]
    contexts = [c.strip() for c in contexts_text.split("\n") if c.strip()]
    competitor_titles = parse_competitor_titles(competitor_text)

    if is_parent:
        color = ""

    common_args = {
        "brand": brand,
        "main_kw": main_kw,
        "size": size,
        "set_size": set_size,
        "material": material,
        "features": features,
        "synonyms": synonyms,
        "contexts": contexts,
        "color": color,
        "no_commas": no_commas,
    }

    # Generuj 3 wersje z auto-trimowaniem
    titles_data = {}
    for style in ["A", "B", "C"]:
        title = trim_to_fit(build_title, style=style, **common_args)
        v = validate_title(title, no_commas=no_commas)
        titles_data[style] = {
            "title": title,
            "length": v["length"],
            "valid": v["valid"],
            "issues": v["issues"],
        }

    # === Wyświetl wyniki ===
    st.success("✅ Wygenerowano tytuły!")

    for style, label in [
        ("A", "🅰️ A – Klasyczny zlepek (max keywords)"),
        ("B", "🅱️ B – Hybryda (REKOMENDOWANE)"),
        ("C", "🅲 C – Zdaniowa naturalna"),
    ]:
        data = titles_data[style]
        with st.expander(f"{label} ({data['length']} zn.)", expanded=(style == "B")):
            st.code(data["title"], language=None)
            if data["valid"]:
                st.success("✅ Spełnia wszystkie reguły Amazon DE")
            else:
                st.error(f"❌ Problemy: {'; '.join(data['issues'])}")
            st.caption(f"Długość: **{data['length']}** / {TITLE_LIMIT} znaków")

    # === Generowanie Excel ===
    st.markdown("---")
    st.subheader("📥 Pobierz analizę")

    product_params = {
        "Marka": brand,
        "Główna fraza": main_kw,
        "Wymiar": size,
        "Set": set_size,
        "Materiał": material,
        "Cechy": ", ".join(features),
        "Synonimy": ", ".join(synonyms),
        "Konteksty": ", ".join(contexts),
        "Kolor": color if color else "(parent - brak)",
        "Bez przecinków": "Tak" if no_commas else "Nie",
        "Parent-tytuł": "Tak" if is_parent else "Nie",
    }

    excel_bytes = build_excel(
        titles=titles_data,
        cerebro_df=cerebro_df,
        competitor_titles=competitor_titles,
        backend_terms=backend_text,
        product_params=product_params,
    )

    safe_name = re.sub(r"[^A-Za-z0-9_-]+", "_", main_kw)[:30] or "tytul"
    st.download_button(
        label="⬇️ Pobierz plik Excel",
        data=excel_bytes,
        file_name=f"Amazon_DE_tytul_{safe_name}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

    # === Dodatkowa analiza ===
    if competitor_titles:
        with st.expander("🔍 Słowa najczęściej używane przez konkurencję", expanded=False):
            kws = extract_keywords_from_titles(competitor_titles)
            df_kw = pd.DataFrame(kws, columns=["Słowo", "Liczba wystąpień"])
            st.dataframe(df_kw, use_container_width=True, hide_index=True)

    if not cerebro_df.empty:
        with st.expander("📊 Top 20 keywordów z Cerebro", expanded=False):
            display = cerebro_df.head(20)[
                ["Keyword Phrase", "Search Volume", "Keyword Sales"]
            ]
            st.dataframe(display, use_container_width=True, hide_index=True)

else:
    st.info(
        "👈 Wypełnij parametry produktu w panelu po lewej, wgraj plik Cerebro "
        "i wklej tytuły konkurencji, a następnie kliknij **Generuj tytuły**."
    )
