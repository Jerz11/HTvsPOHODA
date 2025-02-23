# main.py

import streamlit as st
import pandas as pd
import io
import zipfile
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import PatternFill, Font

st.set_page_config(
    page_title="HotelTime vs Pohoda - Porovnání sestav",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Vlastní CSS pro vylepšení vzhledu
st.markdown("""
    <style>
        /* Hlavní nadpis */
        .main-header {
            color: #9C27B0;
            font-size: 2.5rem;
            font-weight: 600;
            margin-bottom: 2rem;
            text-align: center;
            padding: 1rem;
            background: rgba(156, 39, 176, 0.1);
            border-radius: 10px;
        }
        
        /* Sidebar */
        .css-1d391kg {
            padding: 2rem 1rem;
        }
        
        /* Upload area */
        .uploadedFile {
            background: rgba(156, 39, 176, 0.1) !important;
            border-radius: 10px !important;
            padding: 1rem !important;
        }
        
        /* Tabulka */
        .stDataFrame {
            background: rgba(0, 0, 0, 0.2);
            padding: 1rem;
            border-radius: 10px;
        }
        
        /* Export button */
        .stDownloadButton {
            background-color: #9C27B0 !important;
            color: white !important;
            padding: 0.5rem 1rem !important;
            border-radius: 5px !important;
            border: none !important;
            cursor: pointer !important;
            margin-top: 1rem !important;
        }
        
        /* Celkový rozdíl */
        .total-diff {
            background: rgba(156, 39, 176, 0.1);
            padding: 1rem;
            border-radius: 5px;
            margin-top: 1rem;
            font-size: 1.2rem;
        }
        
        /* Info messages */
        .stInfo {
            background-color: rgba(156, 39, 176, 0.1) !important;
            color: #9C27B0 !important;
            padding: 1rem !important;
            border-radius: 5px !important;
        }
    </style>
""", unsafe_allow_html=True)

def get_report_type(file):
    """
    Určí typ reportu na základě názvu souboru.
    Pokud název obsahuje 'HT', vrací 'HT'.
    Pokud název obsahuje 'FV', vrací 'FV'.
    """
    if "HT" in file.name:
        return "HT"
    elif "FV" in file.name:
        return "FV"
    else:
        return None

def remove_styles_from_excel(file_bytes):
    """
    Odstraní soubor xl/styles.xml z excelovského archivu.
    Vrací BytesIO se "očistěným" obsahem, který lze načíst bez chyb způsobených neplatnými styly.
    """
    input_io = io.BytesIO(file_bytes)
    output_io = io.BytesIO()
    with zipfile.ZipFile(input_io, 'r') as zin:
        with zipfile.ZipFile(output_io, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename == 'xl/styles.xml':
                    # Nepřidáváme styles.xml do nového archivu.
                    continue
                data = zin.read(item.filename)
                zout.writestr(item, data)
    output_io.seek(0)
    return output_io

def read_raw_excel(file):
    """
    Načte Excel soubor a vrátí DataFrame s původními daty (nezkrácený),
    se stejnou logikou odstranění stylů jako v process_excel.
    """
    try:
         file_bytes = file.read()
         # Nepotřebujeme reset pointer, neboť používáme kopii pomocí BytesIO
    except Exception as e:
         st.error(f"Chyba při čtení souboru: {e}")
         return None

    file_io = io.BytesIO(file_bytes)
    try:
         df = pd.read_excel(
             file_io,
             sheet_name=0,
             engine_kwargs={'read_only': True, 'ignore_styles': True}
         )
    except (TypeError, ValueError) as e:
         if "ignore_styles" in str(e) or "stylesheet" in str(e):
             cleaned_io = remove_styles_from_excel(file_bytes)
             try:
                 df = pd.read_excel(
                     cleaned_io,
                     sheet_name=0,
                     engine_kwargs={'read_only': True}
                 )
             except Exception as e2:
                 st.error(f"Chyba při načítání excel souboru po odstranění stylů: {e2}")
                 return None
         else:
             st.error(f"Chyba při načítání excel souboru: {e}")
             return None
    except Exception as e:
         st.error(f"Chyba při načítání excel souboru: {e}")
         return None

    return df

def process_excel(file, report_type):
    """
    Načte excel, seskupí doklady a sečte částky.
    Reporty mají rozdílné názvy sloupců, proto jsou definovány podmíněně.
    """
    # Načteme celý obsah souboru do bytu (file stream se přečte jen jednou).
    try:
        file_bytes = file.read()
    except Exception as e:
        st.error(f"Chyba při čtení souboru: {e}")
        return None

    # Původní pokus načtení s ignore_styles.
    file_io = io.BytesIO(file_bytes)
    try:
         df = pd.read_excel(
             file_io,
             sheet_name=0,
             engine_kwargs={'read_only': True, 'ignore_styles': True}
         )
    except (TypeError, ValueError) as e:
         if "ignore_styles" in str(e) or "stylesheet" in str(e):
             # Odstraníme stylovací list a pokusíme se načíst data z "očistěného" souboru.
             cleaned_io = remove_styles_from_excel(file_bytes)
             try:
                 df = pd.read_excel(
                     cleaned_io,
                     sheet_name=0,
                     engine_kwargs={'read_only': True}
                 )
             except Exception as e2:
                 st.error(f"Chyba při načítání excel souboru po odstranění stylů: {e2}")
                 return None
         else:
             st.error(f"Chyba při načítání excel souboru: {e}")
             return None
    except Exception as e:
         st.error(f"Chyba při načítání excel souboru: {e}")
         return None
    
    if report_type == "HT":
        invoice_col = "Číslo dokladu"
        amount_col = "Celkem s DPH"
    elif report_type == "FV":
        invoice_col = "Číslo"
        amount_col = "Celkem"
    else:
        st.error("Neznámý typ reportu")
        return None
    
    # Nejprve vyčistíme sloupec s dokladem (odstraníme oddělovače)
    df[invoice_col] = df[invoice_col].astype(str).str.replace(" ", "").str.strip()

    # Vyčistíme a převedeme na číslo hodnoty ve sloupci s částkou: odstraníme mezery a nahradíme čárku tečkou
    if df[amount_col].dtype == object:
        df[amount_col] = df[amount_col].astype(str).str.replace(" ", "").str.replace(",", ".")
    df[amount_col] = pd.to_numeric(df[amount_col], errors='coerce')

    # Pro report typu HT navíc získáme sloupce "Odběratel" a "DUZP", pokud existují.
    if report_type == "HT":
        extra_cols = {}
        if "Odběratel" in df.columns:
            extra_cols["Odběratel"] = "first"
        if "DUZP" in df.columns:
            extra_cols["DUZP"] = "first"
        if extra_cols:
            aggregated = df.groupby(invoice_col, as_index=False).agg({amount_col: "sum", **extra_cols})
        else:
            aggregated = df.groupby(invoice_col, as_index=False)[amount_col].sum()
    else:
        aggregated = df.groupby(invoice_col, as_index=False)[amount_col].sum()

    # Přejmenujeme sloupce pro sjednocení výstupu
    if report_type == "HT":
         aggregated.rename(columns={invoice_col: "Doklad", amount_col: "Částka HT"}, inplace=True)
    else:
         aggregated.rename(columns={invoice_col: "Doklad", amount_col: "Částka FV"}, inplace=True)

    return aggregated

def compare_reports(df_ht, df_fv):
    """
    Porovná reporty na základě společného sloupce 'Doklad'.
    Přidá sloupec 'Rozdíl', který ukáže rozdíl částek pro doklady,
    které se nachází v obou reportech.
    
    TODO: Oddělit doklady společné a nespárované, přidat řádek se součtem částek.
    """
    # Použijeme outer join pro získání všech dokladů
    merged = pd.merge(df_ht, df_fv, on="Doklad", how="outer")
    
    # Převedeme sloupce s částkami na typ číslo
    merged["Částka HT"] = pd.to_numeric(merged["Částka HT"], errors='coerce')
    merged["Částka FV"] = pd.to_numeric(merged["Částka FV"], errors='coerce')
    
    # Vektorově spočteme rozdíl obou částek: Pohoda minus HotelTime
    merged["Rozdíl"] = merged["Částka FV"] - merged["Částka HT"]
    
    # Přidáme sloupec pro označení typu řádku
    merged["Status"] = "Nespárovaný"
    merged.loc[merged["Částka HT"].notna() & merged["Částka FV"].notna(), "Status"] = "Spárovaný"
    
    # Použijeme pomocný sloupec pro numerické řazení
    merged["Doklad_sort"] = pd.to_numeric(merged["Doklad"], errors='coerce')
    
    # Seřadíme nejdřív podle statusu (spárované nahoře) a pak podle čísla dokladu
    merged.sort_values(by=["Status", "Doklad_sort"], ascending=[True, True], inplace=True)
    merged.drop(columns=["Doklad_sort"], inplace=True)
    
    # Změníme názvy sloupců
    merged.rename(columns={
        "Částka HT": "Částka HotelTime",
        "Částka FV": "Částka Pohoda"
    }, inplace=True)
    
    # Odebereme sloupec DUZP, pokud existuje
    if "DUZP" in merged.columns:
        merged = merged.drop(columns=["DUZP"])
    
    # Uspořádáme sloupce
    cols_order = []
    if "Doklad" in merged.columns:
        cols_order.append("Doklad")
    if "Odběratel" in merged.columns:
        cols_order.append("Odběratel")
    if "Částka HotelTime" in merged.columns:
        cols_order.append("Částka HotelTime")
    if "Částka Pohoda" in merged.columns:
        cols_order.append("Částka Pohoda")
    if "Rozdíl" in merged.columns:
        cols_order.append("Rozdíl")
    cols_order.append("Status")
    merged = merged[cols_order]
    
    return merged

def export_reports(df_ht_raw, df_fv_raw, differences):
    """
    Vytvoří Excel report se třemi listy:
      - "HotelTime report": původní HT data (nezkrácená, se všemi řádky),
      - "Pohoda report": původní FV data (nezkrácená),
      - "Rozdíly": obsahující tabulku porovnání.
    """
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
         df_ht_raw.to_excel(writer, sheet_name="HotelTime report", index=False)
         df_fv_raw.to_excel(writer, sheet_name="Pohoda report", index=False)
         differences.to_excel(writer, sheet_name="Rozdíly", index=False)
         
         # Získáme worksheet pro formátování
         worksheet = writer.sheets["Rozdíly"]
         
         # Najdeme sloupec s rozdíly
         rozdil_col = None
         for idx, col in enumerate(differences.columns):
             if col == "Rozdíl":
                 rozdil_col = idx
                 break
         
         if rozdil_col is not None:
             # Převedeme číslo sloupce na písmeno (0=A, 1=B, atd.)
             col_letter = chr(65 + rozdil_col)
             last_col = chr(65 + len(differences.columns) - 1)
             # Definujeme styly
             red_fill = PatternFill(start_color='990000', end_color='990000', fill_type='solid')
             white_font = Font(color='FFFFFF')
             
             # Aplikujeme podmíněné formátování na všechny řádky kromě hlavičky
             # Pro hodnoty větší než 1
             worksheet.conditional_formatting.add(
                 f'{col_letter}2:{col_letter}{len(differences) + 1}',
                 CellIsRule(
                     operator='greaterThan',
                     formula=['1'],
                     fill=red_fill,
                     font=white_font
                 )
             )
             # Pro hodnoty menší než -1
             worksheet.conditional_formatting.add(
                 f'{col_letter}2:{col_letter}{len(differences) + 1}',
                 CellIsRule(
                     operator='lessThan',
                     formula=['-1'],
                     fill=red_fill,
                     font=white_font
                 )
             )
    return output.getvalue()

def main():
    st.markdown('<h1 class="main-header">HotelTime vs Pohoda</h1>', unsafe_allow_html=True)
    
    # Sidebar styling
    st.sidebar.markdown("""
        <div style='text-align: center; padding: 1rem; background: rgba(156, 39, 176, 0.1); border-radius: 10px; margin-bottom: 2rem;'>
            <h3 style='color: #9C27B0;'>Nahrání souborů</h3>
        </div>
    """, unsafe_allow_html=True)
    st.sidebar.write("Nahrajte prosím oba excel reporty (jeden s 'HT' a jeden s 'FV' v názvu souboru).")
    uploaded_files = st.sidebar.file_uploader("Nahrát excel reporty", type=["xlsx", "xls"], accept_multiple_files=True)
    
    df_ht = None
    df_fv = None
    df_ht_orig = None
    df_fv_orig = None

    if uploaded_files:
        for file in uploaded_files:
            report_type = get_report_type(file)
            if not report_type:
                st.error(f"Nelze určit typ pro soubor: {file.name}")
                continue
            file_bytes = file.getvalue()
            if report_type == "HT":
                df_ht = process_excel(io.BytesIO(file_bytes), "HT")
                df_ht_orig = read_raw_excel(io.BytesIO(file_bytes))
            elif report_type == "FV":
                df_fv = process_excel(io.BytesIO(file_bytes), "FV")
                df_fv_orig = read_raw_excel(io.BytesIO(file_bytes))
        
        if df_ht is not None and df_fv is not None:
            result = compare_reports(df_ht, df_fv)
            st.write("Výsledek porovnání:")
            # Použijeme stylování pro zvýraznění řádků s nenulovým Rozdíl:
            def highlight_nonzero(row):
                return ['background-color: #990000; color: white' if (row['Rozdíl'] > 1 or row['Rozdíl'] < -1) else '' for _ in row]

            result_styled = result.style.apply(highlight_nonzero, axis=1)
            # Naformátujeme sloupce s částkami a rozdílem: max dvě desetinná místa, oddělovač tisíců
            result_styled = result_styled.format({
                "Částka HotelTime": "{:,.2f}",
                "Částka Pohoda": "{:,.2f}",
                "Rozdíl": "{:,.2f}"
            })

            st.dataframe(result_styled, use_container_width=True)

            total_diff = result["Rozdíl"].sum()
            st.markdown(
                f"""<div class='total-diff' style='text-align: right;'>
                    <span style='color: {"#ff4444" if abs(total_diff) > 1 else "#4CAF50"}'>
                        <b>Celkový rozdíl: {total_diff:,.2f}</b>
                    </span>
                </div>""",
                unsafe_allow_html=True
            )

            excel_data = export_reports(df_ht_orig, df_fv_orig, result)
            st.download_button(
                label="Exportuj rozdíly",
                data=excel_data,
                file_name="HT_vs_Pohoda.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.info("Prosím nahrajte oba reporty (HT a FV).")

if __name__ == '__main__':
    main()
