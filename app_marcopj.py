from datetime import datetime
from typing_extensions import Buffer
import streamlit as st
import pandas as pd
import os

st.set_page_config(page_title="App Unificata", layout="wide")

menu = st.sidebar.selectbox("📋 Seleziona una funzione", [
    "Importa Tracciato",
    "Importa CSV"
])


    
    # AUTO-GENERATED UNIFIED SCRIPT
    
    
    # Deve essere la prima chiamata

    
if menu == "Importa Tracciato":
        
        import zipfile
        import os
        import io
        
        FIELDS = [
            (1, 6, 'Codice Cliente Mittente'),
            (7, 12, 'Codice Tariffa Arco'),
            (13, 22, 'Codice Raggruppamento Fatture'),
            (23, 32, 'Codice Marcatura Colli'),
            (33, 67, 'Mittente - Ragione Sociale'),
            (68, 97, 'Mittente - Indirizzo'),
            (98, 102, 'Mittente - CAP'),
            (103, 132, 'Mittente - Localita'),
            (133, 134, 'Mittente - Provincia'),
            (135, 149, 'Bolla / Fattura Numero'),
            (150, 157, 'Bolla / Fattura Data'),
            (158, 172, 'Numero Ordine / Consegna'),
            (173, 180, 'Data Ordine / Consegna'),
            (181, 195, 'Destinatario - Codice Cliente'),
            (196, 240, 'Destinatario - Ragione Sociale'),
            (276, 310, 'Destinatario - Indirizzo'),
            (311, 318, 'Destinatario - CAP'),
            (319, 348, 'Destinatario - Localita'),
            (349, 350, 'Destinatario - Provincia'),
            (351, 353, 'Destinatario - Nazione'),
            (354, 354, 'Tipo Porto'),
            (355, 359, 'Totale Colli'),
            (360, 362, 'Totale Etichette'),
            (363, 365, 'Bancali da rendere'),
            (366, 368, 'Bancali a perdere'),
            (369, 375, 'Peso in Kg'),
            (380, 387, 'Importo Contrassegno'),
            (918, 1099, 'Informazioni su etichetta'),
            (1100, 1100, 'Flag Fine Record')
        ]
        
        def estrai_record_da_file(file, nome_file):
            records = []
            scarti = []
            lines = file.getvalue().decode("utf-8").splitlines()
            for i, line in enumerate(lines):
                riga_corrente = i + 1
                if len(line) >= 1100:
                    record = {"File Origine Tra": nome_file}
                    for start, end, name in FIELDS:
                        value = line[start-1:end].strip()
                        if name in ['Totale Colli', 'Totale Etichette', 'Bancali da rendere', 'Bancali a perdere']:
                            value = int(value) if value.isdigit() else None
                        elif name == 'Peso in Kg':
                            value = f"{float(value)/10:.1f}".replace('.', ',') if value.isdigit() else ""
                        elif name == 'Importo Contrassegno':
                            value = f"{float(value)/100:.2f}".replace('.', ',') if value.isdigit() else ""
                        record[name] = value
                    raw_date = record.get('Bolla / Fattura Data')
                    if raw_date and raw_date.isdigit() and len(raw_date) == 8:
                        try:
                            formatted_date = f"{raw_date[6:8]}/{raw_date[4:6]}/{raw_date[0:4]}"
                        except:
                            formatted_date = ""
                    else:
                        formatted_date = ""
                    record['Data calc'] = formatted_date
                    record["Riga Tra"] = riga_corrente  # 👈 AGGIUNTA QUI
                    record = {'File Origine Tra': record.pop('File Origine Tra'), 'Riga Tra' : record.pop ('Riga Tra' ), 'Data calc': record.pop('Data calc'), **record}
                    records.append(record)
                else:
                    scarti.append({
                    'Numero Riga Tra': riga_corrente,
                        "File Origine Tra": nome_file,
                        "Numero Riga": i + 1,
                        "Lunghezza Riga": len(line),
                        "Contenuto": line
                    })
            return records, scarti
        
        def process_file(uploaded_file):
            all_records = []
            all_scarti = []
            summary = []
        
            if uploaded_file.name.endswith(".zip"):
                with zipfile.ZipFile(uploaded_file) as zip_ref:
                    for name in zip_ref.namelist():
                        if name.endswith(".txt"):
                            with zip_ref.open(name) as file:
                                content = io.BytesIO(file.read())
                                recs, bads = estrai_record_da_file(content, name)
                                all_records.extend(recs)
                                all_scarti.extend(bads)
                                summary.append({"File": name, "Record Validi": len(recs), "Scartati": len(bads)})
            elif uploaded_file.name.endswith(".txt"):
                recs, bads = estrai_record_da_file(uploaded_file, uploaded_file.name)
                all_records.extend(recs)
                all_scarti.extend(bads)
                summary.append({"File": uploaded_file.name, "Record Validi": len(recs), "Scartati": len(bads)})
            
            return pd.DataFrame(all_records), pd.DataFrame(all_scarti), pd.DataFrame(summary)
        st.header("Importazione Fixed Format Ordini a Arco")
        st.title("📦 Convertitore Tracciato per Maroil")
        st.markdown("""
        ✅ Carica un `.zip` con file a "larghezza fissa"  
        ✅ Crea antiprima dei risultati  
        ✅ Salva in Excel:
        - Dati validi → foglio **Dati_Tra**
        - Scarti → foglio **Scarti_Tra**
        - Conteggio → fogli **Riepilogo_Tra**, ecc*
        - Conteggio righe
        """)
        
        uploaded_file = st.file_uploader("Carica un file .txt o .zip", type=["txt", "zip"])
        
        if uploaded_file:
            with st.spinner("⏳ Elaborazione in corso..."):
                df_dati, df_scarti, df_summary = process_file(uploaded_file)
        
                if not df_dati.empty:
                    st.subheader("✅ Record Validi")
                    st.dataframe(df_dati.head(50))
        
                if not df_scarti.empty:
                    st.subheader("⚠️ Record Scartati")
                    st.dataframe(df_scarti.head(50))
        
                st.subheader("📊 Riepilogo")
                st.dataframe(df_summary)
        
                # Raggruppamenti
                df_summary_by_date = df_dati.groupby('Data calc').size().reset_index(name='Totale Record')
                df_summary_by_dest = df_dati.groupby('Destinatario - Ragione Sociale').size().reset_index(name='Totale Record')
        
                # Grafico settimanale robusto
                st.subheader("📈 Grafico settimanale per Data calc (con date)")
                try:
                    df_dati['Data calc parsed'] = pd.to_datetime(df_dati['Data calc'], format="%d/%m/%Y", errors='coerce')
                    df_week = df_dati.dropna(subset=['Data calc parsed'])
                    df_week['Week Start'] = df_week['Data calc parsed'].dt.to_period('W').apply(lambda r: r.start_time.strftime('%d/%m/%Y'))
                    df_week['Week End'] = df_week['Data calc parsed'].dt.to_period('W').apply(lambda r: r.end_time.strftime('%d/%m/%Y'))
                    df_week['Settimana'] = df_week['Week Start'] + " - " + df_week['Week End']
                    df_summary_by_week = df_week.groupby('Settimana').size().reset_index(name='Totale Record')
                    st.bar_chart(df_summary_by_week.set_index('Settimana'))
                except Exception as e:
                    st.error(f"Errore nella generazione del grafico settimanale: {e}")
        
                st.subheader("🏢 Grafico per Destinatario - Ragione Sociale (Top 20)")
                top_dest = df_summary_by_dest.sort_values(by='Totale Record', ascending=False).head(20)
                st.bar_chart(top_dest.set_index('Destinatario - Ragione Sociale'))
        
        # Calcolo avanzato: pulizia e conversioni numeriche
                df_dati['Importo Contrassegno (float)'] = df_dati['Importo Contrassegno'].str.replace(',', '.', regex=False).astype(float, errors='ignore')
                df_dati['Peso in Kg (float)'] = df_dati['Peso in Kg'].str.replace(',', '.', regex=False).astype(float, errors='ignore')
        
                # Top 10 destinatari per importo contrassegno
                st.subheader("💰 Top 10 Destinatari per Importo Contrassegno")
                top_contrassegno = df_dati.groupby('Destinatario - Ragione Sociale')['Importo Contrassegno (float)'].sum(numeric_only=True).reset_index()
                top_contrassegno = top_contrassegno.sort_values(by='Importo Contrassegno (float)', ascending=False).head(10)
                st.bar_chart(top_contrassegno.set_index('Destinatario - Ragione Sociale'))
        
                # Top 10 destinatari per peso totale
                st.subheader("⚖️ Top 10 Destinatari per Peso Totale")
                top_peso = df_dati.groupby('Destinatario - Ragione Sociale')['Peso in Kg (float)'].sum(numeric_only=True).reset_index()
                top_peso = top_peso.sort_values(by='Peso in Kg (float)', ascending=False).head(10)
                st.bar_chart(top_peso.set_index('Destinatario - Ragione Sociale'))
        
                # Trend settimanale di peso e contrassegno
                st.subheader("📈 Trend Settimanale - Peso e Contrassegno")
                if 'Data calc parsed' in df_week.columns:
                    df_trend = df_week.copy()
                    df_trend['Peso (float)'] = df_trend['Peso in Kg'].str.replace(',', '.', regex=False).astype(float, errors='ignore')
                    df_trend['Contrassegno (float)'] = df_trend['Importo Contrassegno'].str.replace(',', '.', regex=False).astype(float, errors='ignore')
                    df_trend_grouped = df_trend.groupby('Settimana')[['Peso (float)', 'Contrassegno (float)']].sum().reset_index()
                    st.line_chart(df_trend_grouped.set_index('Settimana'))
        
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    if not df_dati.empty:
                        df_dati.to_excel(writer, index=False, sheet_name="Dati_Tra")
                    if not df_scarti.empty:
                        df_scarti.to_excel(writer, index=False, sheet_name="Scarti_Tra")
                    df_summary.to_excel(writer, index=False, sheet_name="Riepilogo_Tra")
                    df_summary_by_date.to_excel(writer, index=False, sheet_name="Riep_DataCalc_Tra")
                    df_summary_by_dest.to_excel(writer, index=False, sheet_name="Riep_Destinatario_Tra")
                output.seek(0)
        
                st.download_button(
                    label="📥 Scarica Excel",
                    data=output.getvalue(),
                    file_name=f"export_tracciato_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
    
elif menu == "Importa CSV":
        st.header("Importazione CSV Fatture Arco")
        
        import io
        import zipfile
        from pathlib import Path
        from datetime import datetime
        from collections import defaultdict
        
        
        st.title("📄 Convertitore CSV/ZIP in Excel per Maroil")
        st.markdown("""
        ✅ Carica un `.zip` con file `.csv` separati da `;`   
        ✅ Crea antiprima dei risultati   
        ✅ Salva in Excel:
        - Dati validi → foglio **Dati Fat**
        - Scarti → foglio **Scarti**
        - Conteggio → fogli **Riepilogo Fat** e **Scarti Fat**
        - Rimozione virgolette (`"`) e intestazione unica
        """)
        
        uploaded_file = st.file_uploader("Carica un file ZIP con CSV", type="zip")
        
        if uploaded_file:
            dati_records = []
            scarti_records = []
            header_fields = None
        
            try:
                with zipfile.ZipFile(uploaded_file) as archive:
                    csv_list = [name for name in archive.namelist() if name.endswith(".csv")]
                    if not csv_list:
                        st.warning("⚠️ Nessun file CSV trovato nello ZIP.")
        
                    for csv_name in csv_list:
                        try:
                            with archive.open(csv_name) as csv_file:
                                raw_lines = csv_file.read().decode("utf-8", errors="ignore").splitlines()
                                if not raw_lines:
                                    continue
        
                                if header_fields is None:
                                    header_fields = [h.strip().replace('"', '') for h in raw_lines[0].split(";")]
                                num_colonne_attese = len(header_fields)
        
                                for i, line in enumerate(raw_lines[1:], start=2):
                                    line_clean = line.replace('"', '')
                                    cols = line_clean.split(";")
                                    row_info = {"File origine Fat": csv_name, "Riga CSV": i}
        
                                    if len(cols) == num_colonne_attese:
                                        record = {**row_info, **{header_fields[j]: cols[j].strip() for j in range(num_colonne_attese)}}
                                        if "DT_FATT" in record and record["DT_FATT"].isdigit() and len(record["DT_FATT"]) == 8:
                                           record["Data Calc CSV"] = f"{record['DT_FATT'][6:8]}/{record['DT_FATT'][4:6]}/{record['DT_FATT'][0:4]}"
                                        else:
                                            record["Data Calc CSV"] = ""
                                        dati_records.append(record)
                                    else:
                                        scarti_records.append({**row_info, "Contenuto": line, "Errore": "Colonne non conformi"})
        
                            st.success(f"✅ Elaborato: {csv_name}")
                        except Exception as e:
                            st.error(f"Errore nel file {csv_name}: {e}")
            except Exception as e:
                st.error(f"Errore nell’apertura dello ZIP: {e}")
        
            # Solo 3 righe per anteprima veloce
            if dati_records:
                st.subheader("👁️ Anteprima Dati Fat")
                st.dataframe(pd.DataFrame(dati_records[:5]))
            if scarti_records:
                st.subheader("🚫 Anteprima Scarti")
                st.dataframe(pd.DataFrame(scarti_records[:5]))
        
            if dati_records or scarti_records:
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
                    # Dati Fat
                    pd.DataFrame(dati_records).to_excel(writer, sheet_name="Dati_Fat", index=False)
                    # Scarti
                    # Riepilogo Fat
                    df_riepilogo = pd.DataFrame(dati_records).groupby("File origine Fat").size().reset_index(name="Record Validi")
                    df_riepilogo.to_excel(writer, sheet_name="Riepilogo_Fat", index=False)
        
                    # Scarti Fat
                    df_scarti_fat = pd.DataFrame([
                        {
                            "File origine Fat": r.get("File origine Fat", ""),
                            "Riga CSV": r.get("Riga CSV", ""),
                            "Errore": r.get("Errore", "Errore sconosciuto")
                        } for r in scarti_records
                    ])
                    df_scarti_fat.to_excel(writer, sheet_name="Scarti_Fat", index=False)
        
                st.download_button(
                    label="📥 Scarica Excel completo",
                    data=buffer.getvalue(),
                    file_name=f"export_fatture_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
    
elif menu == "Confronta Base":
        st.header("Confronto Dati")
        
        st.title("✅ Verifica coerenza tra Tracciato e Fattura")
        
        file1 = st.file_uploader("Carica il file Excel del Tracciato", type=["xlsx"])
        file2 = st.file_uploader("Carica il file Excel della Fattura", type=["xlsx"])
        
        def normalizza_colli(val):
            try:
                return int(str(val).lstrip("0"))
            except:
                return None
        
        def normalizza_peso(val):
            try:
                return float(str(val).replace(",", "."))
            except:
                return None
        
        if file1 and file2:
            try:
                tracciato = pd.read_excel(file1, sheet_name="Dati_Tra")
                fattura = pd.read_excel(file2, sheet_name="Dati_Fat")
        
                tracciato.columns = tracciato.columns.str.strip().str.lower()
                fattura.columns = fattura.columns.str.strip().str.lower()
        
                # 🔧 Uniforma i tipi delle colonne chiave a stringa
                tracciato["bolla / fattura numero"] = tracciato["bolla / fattura numero"].astype(str).str.strip()
                fattura["rif_cli"] = fattura["rif_cli"].astype(str).str.strip()
        
                trac = tracciato[["bolla / fattura numero", "totale colli", "peso in kg"]].copy()
                fat = fattura[["rif_cli", "num_coll", "tot_peso"]].copy()
        
                trac["totale colli norm"] = trac["totale colli"].apply(normalizza_colli)
                fat["totale colli norm"] = fat["num_coll"].apply(normalizza_colli)
        
                trac["peso norm"] = trac["peso in kg"].apply(normalizza_peso)
                fat["peso norm"] = fat["tot_peso"].apply(normalizza_peso)
        
                merged = pd.merge(
                    fat, trac,
                    how="left",
                    left_on="rif_cli",
                    right_on="bolla / fattura numero",
                    suffixes=("_fat", "_trac")
                )
        
                # ✅ Corretto: riferimenti ai nomi giusti dopo il merge
                merged["ok_colli"] = merged["totale colli norm_fat"] == merged["totale colli norm_trac"]
                merged["ok_peso"] = merged["peso norm_fat"].round(2) == merged["peso norm_trac"].round(2)
        
                errori = merged[~(merged["ok_colli"] & merged["ok_peso"])]
        
                st.subheader("🔎 Risultati Confronto")
                st.write(f"Totale righe confrontate: {len(merged)}")
                st.write(f"✅ Righe corrette: {(merged['ok_colli'] & merged['ok_peso']).sum()}")
                st.write(f"❌ Righe con errori: {len(errori)}")
        
                st.subheader("❌ Dettaglio Errori")
                if errori.empty:
                    st.success("Tutti i record sono coerenti!")
                else:
                    st.dataframe(errori)
        
                import io
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                    merged.to_excel(writer, sheet_name="Confronto Completo", index=False)
                    errori.to_excel(writer, sheet_name="Errori", index=False)
        
                st.download_button(
                    "📥 Scarica risultati in Excel",
                    data=output.getvalue(),
                    file_name=f"verifica_fattura_tracciato_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                )
        
            except Exception as e:
                st.error(f"Errore durante l'elaborazione: {e}")
        else:
            st.info("📂 Carica entrambi i file per eseguire la validazione.")

elif menu == "Importa CSV":
    st.header("📥 Importazione CSV")
    
    # AUTO-GENERATED UNIFIED SCRIPT
    
    
    # Deve essere la prima chiamata
    
    # Sidebar menu
    menu = st.sidebar.selectbox("Scegli la funzione", ["Importa Tracciato", "Importa CSV"])
    
    if menu == "Importa Tracciato":
        st.header("Importazione Tracciato")
        
        import zipfile
        import os
        import io
        
        FIELDS = [
            (1, 6, 'Codice Cliente Mittente'),
            (7, 12, 'Codice Tariffa Arco'),
            (13, 22, 'Codice Raggruppamento Fatture'),
            (23, 32, 'Codice Marcatura Colli'),
            (33, 67, 'Mittente - Ragione Sociale'),
            (68, 97, 'Mittente - Indirizzo'),
            (98, 102, 'Mittente - CAP'),
            (103, 132, 'Mittente - Localita'),
            (133, 134, 'Mittente - Provincia'),
            (135, 149, 'Bolla / Fattura Numero'),
            (150, 157, 'Bolla / Fattura Data'),
            (158, 172, 'Numero Ordine / Consegna'),
            (173, 180, 'Data Ordine / Consegna'),
            (181, 195, 'Destinatario - Codice Cliente'),
            (196, 240, 'Destinatario - Ragione Sociale'),
            (276, 310, 'Destinatario - Indirizzo'),
            (311, 318, 'Destinatario - CAP'),
            (319, 348, 'Destinatario - Localita'),
            (349, 350, 'Destinatario - Provincia'),
            (351, 353, 'Destinatario - Nazione'),
            (354, 354, 'Tipo Porto'),
            (355, 359, 'Totale Colli'),
            (360, 362, 'Totale Etichette'),
            (363, 365, 'Bancali da rendere'),
            (366, 368, 'Bancali a perdere'),
            (369, 375, 'Peso in Kg'),
            (380, 387, 'Importo Contrassegno'),
            (918, 1099, 'Informazioni su etichetta'),
            (1100, 1100, 'Flag Fine Record')
        ]
        
        def estrai_record_da_file(file, nome_file):
            records = []
            scarti = []
            lines = file.getvalue().decode("utf-8").splitlines()
            for i, line in enumerate(lines):
                riga_corrente = i + 1
                if len(line) >= 1100:
                    record = {"File Origine Tra": nome_file}
                    for start, end, name in FIELDS:
                        value = line[start-1:end].strip()
                        if name in ['Totale Colli', 'Totale Etichette', 'Bancali da rendere', 'Bancali a perdere']:
                            value = int(value) if value.isdigit() else None
                        elif name == 'Peso in Kg':
                            value = f"{float(value)/10:.1f}".replace('.', ',') if value.isdigit() else ""
                        elif name == 'Importo Contrassegno':
                            value = f"{float(value)/100:.2f}".replace('.', ',') if value.isdigit() else ""
                        record[name] = value
                    raw_date = record.get('Bolla / Fattura Data')
                    if raw_date and raw_date.isdigit() and len(raw_date) == 8:
                        try:
                            formatted_date = f"{raw_date[6:8]}/{raw_date[4:6]}/{raw_date[0:4]}"
                        except:
                            formatted_date = ""
                    else:
                        formatted_date = ""
                    record['Data calc'] = formatted_date
                    record["Riga Tra"] = riga_corrente  # 👈 AGGIUNTA QUI
                    record = {'File Origine Tra': record.pop('File Origine Tra'), 'Riga Tra' : record.pop ('Riga Tra' ), 'Data calc': record.pop('Data calc'), **record}
                    records.append(record)
                else:
                    scarti.append({
                    'Numero Riga': riga_corrente,
                        "File Origine Tra": nome_file,
                        "Numero Riga": i + 1,
                        "Lunghezza Riga": len(line),
                        "Contenuto": line
                    })
            return records, scarti
        
        def process_file(uploaded_file):
            all_records = []
            all_scarti = []
            summary = []
        
            if uploaded_file.name.endswith(".zip"):
                with zipfile.ZipFile(uploaded_file) as zip_ref:
                    for name in zip_ref.namelist():
                        if name.endswith(".txt"):
                            with zip_ref.open(name) as file:
                                content = io.BytesIO(file.read())
                                recs, bads = estrai_record_da_file(content, name)
                                all_records.extend(recs)
                                all_scarti.extend(bads)
                                summary.append({"File": name, "Record Validi": len(recs), "Scartati": len(bads)})
            elif uploaded_file.name.endswith(".txt"):
                recs, bads = estrai_record_da_file(uploaded_file, uploaded_file.name)
                all_records.extend(recs)
                all_scarti.extend(bads)
                summary.append({"File": uploaded_file.name, "Record Validi": len(recs), "Scartati": len(bads)})
            
            return pd.DataFrame(all_records), pd.DataFrame(all_scarti), pd.DataFrame(summary)
        
        st.title("📦 Tracciato Converter per Maroil")
        
        uploaded_file = st.file_uploader("Carica un file .txt o .zip", type=["txt", "zip"])
        
        if uploaded_file:
            with st.spinner("⏳ Elaborazione in corso..."):
                df_dati, df_scarti, df_summary = process_file(uploaded_file)
        
                if not df_dati.empty:
                    st.subheader("✅ Record Validi")
                    st.dataframe(df_dati.head(50))
        
                if not df_scarti.empty:
                    st.subheader("⚠️ Record Scartati")
                    st.dataframe(df_scarti.head(50))
        
                st.subheader("📊 Riepilogo")
                st.dataframe(df_summary)
        
                # Raggruppamenti
                df_summary_by_date = df_dati.groupby('Data calc').size().reset_index(name='Totale Record')
                df_summary_by_dest = df_dati.groupby('Destinatario - Ragione Sociale').size().reset_index(name='Totale Record')
        
                # Grafico settimanale robusto
                st.subheader("📈 Grafico settimanale per Data calc (con date)")
                try:
                    df_dati['Data calc parsed'] = pd.to_datetime(df_dati['Data calc'], format="%d/%m/%Y", errors='coerce')
                    df_week = df_dati.dropna(subset=['Data calc parsed'])
                    df_week['Week Start'] = df_week['Data calc parsed'].dt.to_period('W').apply(lambda r: r.start_time.strftime('%d/%m/%Y'))
                    df_week['Week End'] = df_week['Data calc parsed'].dt.to_period('W').apply(lambda r: r.end_time.strftime('%d/%m/%Y'))
                    df_week['Settimana'] = df_week['Week Start'] + " - " + df_week['Week End']
                    df_summary_by_week = df_week.groupby('Settimana').size().reset_index(name='Totale Record')
                    st.bar_chart(df_summary_by_week.set_index('Settimana'))
                except Exception as e:
                    st.error(f"Errore nella generazione del grafico settimanale: {e}")
        
                st.subheader("🏢 Grafico per Destinatario - Ragione Sociale (Top 20)")
                top_dest = df_summary_by_dest.sort_values(by='Totale Record', ascending=False).head(20)
                st.bar_chart(top_dest.set_index('Destinatario - Ragione Sociale'))
        
        # Calcolo avanzato: pulizia e conversioni numeriche
                df_dati['Importo Contrassegno (float)'] = df_dati['Importo Contrassegno'].str.replace(',', '.', regex=False).astype(float, errors='ignore')
                df_dati['Peso in Kg (float)'] = df_dati['Peso in Kg'].str.replace(',', '.', regex=False).astype(float, errors='ignore')
        
                # Top 10 destinatari per importo contrassegno
                st.subheader("💰 Top 10 Destinatari per Importo Contrassegno")
                top_contrassegno = df_dati.groupby('Destinatario - Ragione Sociale')['Importo Contrassegno (float)'].sum(numeric_only=True).reset_index()
                top_contrassegno = top_contrassegno.sort_values(by='Importo Contrassegno (float)', ascending=False).head(10)
                st.bar_chart(top_contrassegno.set_index('Destinatario - Ragione Sociale'))
        
                # Top 10 destinatari per peso totale
                st.subheader("⚖️ Top 10 Destinatari per Peso Totale")
                top_peso = df_dati.groupby('Destinatario - Ragione Sociale')['Peso in Kg (float)'].sum(numeric_only=True).reset_index()
                top_peso = top_peso.sort_values(by='Peso in Kg (float)', ascending=False).head(10)
                st.bar_chart(top_peso.set_index('Destinatario - Ragione Sociale'))
        
                # Trend settimanale di peso e contrassegno
                st.subheader("📈 Trend Settimanale - Peso e Contrassegno")
                if 'Data calc parsed' in df_week.columns:
                    df_trend = df_week.copy()
                    df_trend['Peso (float)'] = df_trend['Peso in Kg'].str.replace(',', '.', regex=False).astype(float, errors='ignore')
                    df_trend['Contrassegno (float)'] = df_trend['Importo Contrassegno'].str.replace(',', '.', regex=False).astype(float, errors='ignore')
                    df_trend_grouped = df_trend.groupby('Settimana')[['Peso (float)', 'Contrassegno (float)']].sum().reset_index()
                    st.line_chart(df_trend_grouped.set_index('Settimana'))
        
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    if not df_dati.empty:
                        df_dati.to_excel(writer, index=False, sheet_name="Dati_Tra")
                    if not df_scarti.empty:
                        df_scarti.to_excel(writer, index=False, sheet_name="Scarti_Tra")
                    df_summary.to_excel(writer, index=False, sheet_name="Riepilogo_Tra")
                    df_summary_by_date.to_excel(writer, index=False, sheet_name="Riep_DataCalc_Tra")
                    df_summary_by_dest.to_excel(writer, index=False, sheet_name="Riep_Destinatario_Tra")
                output.seek(0)
        
                st.download_button(
                    label="📥 Scarica Excel",
                    data=output.getvalue(),
                    file_name=f"export_tracciato_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
    
    elif menu == "Importa CSV":
        
        import io
        import zipfile
        from pathlib import Path
        from datetime import datetime
        from collections import defaultdict
        
        
        st.title("📄 Convertitore Fatture ARCO CSV/ZIP in Excel per Maroil")
        st.markdown("""
        ✅ Carica un `.zip` con file `.csv` separati da `;`  
        ✅ Salva in Excel:
        - Dati validi → foglio **Dati Fat**
        - Scarti → foglio **Scarti**
        - Conteggio → fogli **Riepilogo Fat** e **Scarti Fat**
        - Rimozione virgolette (`"`) e intestazione unica
        """)
        
        uploaded_file = st.file_uploader("Carica un file ZIP con CSV", type="zip")
        
        if uploaded_file:
            dati_records = []
            scarti_records = []
            header_fields = None
        
            try:
                with zipfile.ZipFile(uploaded_file) as archive:
                    csv_list = [name for name in archive.namelist() if name.endswith(".csv")]
                    if not csv_list:
                        st.warning("⚠️ Nessun file CSV trovato nello ZIP.")
        
                    for csv_name in csv_list:
                        try:
                            with archive.open(csv_name) as csv_file:
                                raw_lines = csv_file.read().decode("utf-8", errors="ignore").splitlines()
                                if not raw_lines:
                                    continue
        
                                if header_fields is None:
                                    header_fields = [h.strip().replace('"', '') for h in raw_lines[0].split(";")]
                                num_colonne_attese = len(header_fields)
        
                                for i, line in enumerate(raw_lines[1:], start=2):
                                    line_clean = line.replace('"', '')
                                    cols = line_clean.split(";")
                                    row_info = {"File origine Fat": csv_name, "Riga CSV": i}
        
                                if len(cols) == num_colonne_attese:
                                        record = {**row_info, **{header_fields[j]: cols[j].strip() for j in range(num_colonne_attese)}}
                                        if "DT_FATT" in record and record["DT_FATT"].isdigit() and len(record["DT_FATT"]) == 8:
                                           record["Data Calc CSV"] = f"{record['DT_FATT'][6:8]}/{record['DT_FATT'][4:6]}/{record['DT_FATT'][0:4]}"
                                        else:
                                            record["Data Calc CSV"] = ""
                                        dati_records.append(record)
                                else:
                                        scarti_records.append({**row_info, "Contenuto": line, "Errore": "Colonne non conformi"})
        
        
                            st.success(f"✅ Elaborato: {csv_name}")
                        except Exception as e:
                            st.error(f"Errore nel file {csv_name}: {e}")
            except Exception as e:
                st.error(f"Errore nell’apertura dello ZIP: {e}")
        
            # Solo 3 righe per anteprima veloce
            if dati_records:
                st.subheader("👁️ Anteprima Dati Fat")
                st.dataframe(pd.DataFrame(dati_records[:5]))
            if scarti_records:
                st.subheader("🚫 Anteprima Scarti")
                st.dataframe(pd.DataFrame(scarti_records[:5]))
        
            if dati_records or scarti_records:
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
                    # Dati Fat
                    pd.DataFrame(dati_records).to_excel(writer, sheet_name="Dati_Fat", index=False)
                    # Scarti
                    # Riepilogo Fat
                    df_riepilogo = pd.DataFrame(dati_records).groupby("File origine CSV").size().reset_index(name="Record Validi")
                    df_riepilogo.to_excel(writer, sheet_name="Riepilogo_Fat", index=False)
        
                    # Scarti Fat
                    df_scarti_fat = pd.DataFrame([
                        {
                            "File origine CSV": r.get("File origine CSV", ""),
                            "Riga CSV": r.get("Riga CSV", ""),
                            "Errore": r.get("Errore", "Errore sconosciuto")
                        } for r in scarti_records
                    ])
                    df_scarti_fat.to_excel(writer, sheet_name="Scarti_Fat", index=False)
        
                st.download_button(
                    label="📥 Scarica Excel completo",
                    data=output.getvalue(),
                    file_name=f"export_tracciato_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
    
    elif menu == "Confronta":
        st.header("Confronto Dati")
        
        st.title("✅ Verifica coerenza tra Tracciato e Fattura")
        
        file1 = st.file_uploader("Carica il file Excel del Tracciato", type=["xlsx"])
        file2 = st.file_uploader("Carica il file Excel della Fattura", type=["xlsx"])
        
        def normalizza_colli(val):
            try:
                return int(str(val).lstrip("0"))
            except:
                return None
        
        def normalizza_peso(val):
            try:
                return float(str(val).replace(",", "."))
            except:
                return None
        
        if file1 and file2:
            try:
                tracciato = pd.read_excel(file1, sheet_name="Dati_Tra")
                fattura = pd.read_excel(file2, sheet_name="Dati_Fat")
        
                tracciato.columns = tracciato.columns.str.strip().str.lower()
                fattura.columns = fattura.columns.str.strip().str.lower()
        
                # 🔧 Uniforma i tipi delle colonne chiave a stringa
                tracciato["bolla / fattura numero"] = tracciato["bolla / fattura numero"].astype(str).str.strip()
                fattura["rif_cli"] = fattura["rif_cli"].astype(str).str.strip()
        
                trac = tracciato[["bolla / fattura numero", "totale colli", "peso in kg"]].copy()
                fat = fattura[["rif_cli", "num_coll", "tot_peso"]].copy()
        
                trac["totale colli norm"] = trac["totale colli"].apply(normalizza_colli)
                fat["totale colli norm"] = fat["num_coll"].apply(normalizza_colli)
        
                trac["peso norm"] = trac["peso in kg"].apply(normalizza_peso)
                fat["peso norm"] = fat["tot_peso"].apply(normalizza_peso)
        
                merged = pd.merge(
                    fat, trac,
                    how="left",
                    left_on="rif_cli",
                    right_on="bolla / fattura numero",
                    suffixes=("_fat", "_trac")
                )
        
                # ✅ Corretto: riferimenti ai nomi giusti dopo il merge
                merged["ok_colli"] = merged["totale colli norm_fat"] == merged["totale colli norm_trac"]
                merged["ok_peso"] = merged["peso norm_fat"].round(2) == merged["peso norm_trac"].round(2)
        
                errori = merged[~(merged["ok_colli"] & merged["ok_peso"])]
        
                st.subheader("🔎 Risultati Confronto")
                st.write(f"Totale righe confrontate: {len(merged)}")
                st.write(f"✅ Righe corrette: {(merged['ok_colli'] & merged['ok_peso']).sum()}")
                st.write(f"❌ Righe con errori: {len(errori)}")
        
                st.subheader("❌ Dettaglio Errori")
                if errori.empty:
                    st.success("Tutti i record sono coerenti!")
                else:
                    st.dataframe(errori)
        
                import io
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                    merged.to_excel(writer, sheet_name="Confronto Completo", index=False)
                    errori.to_excel(writer, sheet_name="Errori", index=False)
        
                st.download_button(
                    "📥 Scarica risultati in Excel",
                    data=output.getvalue(),
                    file_name=f"verifica_fattura_tracciato_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                )
        
            except Exception as e:
                st.error(f"Errore durante l'elaborazione: {e}")
        else:
            st.info("📂 Carica entrambi i file per eseguire la validazione.")

elif menu == "Confronta Base":
    st.header("📊 Confronto Dati Base")
    
    # AUTO-GENERATED UNIFIED SCRIPT
    
    
    # Deve essere la prima chiamata
    
    # Sidebar menu
    menu = st.sidebar.selectbox("Scegli la funzione", ["Importa Tracciato", "Importa CSV"])
    
    if menu == "Importa Tracciato":
        st.header("Importazione Tracciato")
        
        import zipfile
        import os
        import io
        
        FIELDS = [
            (1, 6, 'Codice Cliente Mittente'),
            (7, 12, 'Codice Tariffa Arco'),
            (13, 22, 'Codice Raggruppamento Fatture'),
            (23, 32, 'Codice Marcatura Colli'),
            (33, 67, 'Mittente - Ragione Sociale'),
            (68, 97, 'Mittente - Indirizzo'),
            (98, 102, 'Mittente - CAP'),
            (103, 132, 'Mittente - Localita'),
            (133, 134, 'Mittente - Provincia'),
            (135, 149, 'Bolla / Fattura Numero'),
            (150, 157, 'Bolla / Fattura Data'),
            (158, 172, 'Numero Ordine / Consegna'),
            (173, 180, 'Data Ordine / Consegna'),
            (181, 195, 'Destinatario - Codice Cliente'),
            (196, 240, 'Destinatario - Ragione Sociale'),
            (276, 310, 'Destinatario - Indirizzo'),
            (311, 318, 'Destinatario - CAP'),
            (319, 348, 'Destinatario - Localita'),
            (349, 350, 'Destinatario - Provincia'),
            (351, 353, 'Destinatario - Nazione'),
            (354, 354, 'Tipo Porto'),
            (355, 359, 'Totale Colli'),
            (360, 362, 'Totale Etichette'),
            (363, 365, 'Bancali da rendere'),
            (366, 368, 'Bancali a perdere'),
            (369, 375, 'Peso in Kg'),
            (380, 387, 'Importo Contrassegno'),
            (918, 1099, 'Informazioni su etichetta'),
            (1100, 1100, 'Flag Fine Record')
        ]
        
        def estrai_record_da_file(file, nome_file):
            records = []
            scarti = []
            lines = file.getvalue().decode("utf-8").splitlines()
            for i, line in enumerate(lines):
                riga_corrente = i + 1
                if len(line) >= 1100:
                    record = {"File Origine Tra": nome_file}
                    for start, end, name in FIELDS:
                        value = line[start-1:end].strip()
                        if name in ['Totale Colli', 'Totale Etichette', 'Bancali da rendere', 'Bancali a perdere']:
                            value = int(value) if value.isdigit() else None
                        elif name == 'Peso in Kg':
                            value = f"{float(value)/10:.1f}".replace('.', ',') if value.isdigit() else ""
                        elif name == 'Importo Contrassegno':
                            value = f"{float(value)/100:.2f}".replace('.', ',') if value.isdigit() else ""
                        record[name] = value
                    raw_date = record.get('Bolla / Fattura Data')
                    if raw_date and raw_date.isdigit() and len(raw_date) == 8:
                        try:
                            formatted_date = f"{raw_date[6:8]}/{raw_date[4:6]}/{raw_date[0:4]}"
                        except:
                            formatted_date = ""
                    else:
                        formatted_date = ""
                    record['Data calc'] = formatted_date
                    record["Riga Tra"] = riga_corrente  # 👈 AGGIUNTA QUI
                    record = {'File Origine Tra': record.pop('File Origine Tra'), 'Riga Tra' : record.pop ('Riga Tra' ), 'Data calc': record.pop('Data calc'), **record}
                    records.append(record)
                else:
                    scarti.append({
                    'Numero Riga': riga_corrente,
                        "File Origine Tra": nome_file,
                        "Numero Riga": i + 1,
                        "Lunghezza Riga": len(line),
                        "Contenuto": line
                    })
            return records, scarti
        
        def process_file(uploaded_file):
            all_records = []
            all_scarti = []
            summary = []
        
            if uploaded_file.name.endswith(".zip"):
                with zipfile.ZipFile(uploaded_file) as zip_ref:
                    for name in zip_ref.namelist():
                        if name.endswith(".txt"):
                            with zip_ref.open(name) as file:
                                content = io.BytesIO(file.read())
                                recs, bads = estrai_record_da_file(content, name)
                                all_records.extend(recs)
                                all_scarti.extend(bads)
                                summary.append({"File": name, "Record Validi": len(recs), "Scartati": len(bads)})
            elif uploaded_file.name.endswith(".txt"):
                recs, bads = estrai_record_da_file(uploaded_file, uploaded_file.name)
                all_records.extend(recs)
                all_scarti.extend(bads)
                summary.append({"File": uploaded_file.name, "Record Validi": len(recs), "Scartati": len(bads)})
            
            return pd.DataFrame(all_records), pd.DataFrame(all_scarti), pd.DataFrame(summary)
        
        st.title("📦 Tracciato Converter per Maroil")
        
        uploaded_file = st.file_uploader("Carica un file .txt o .zip", type=["txt", "zip"])
        
        if uploaded_file:
            with st.spinner("⏳ Elaborazione in corso..."):
                df_dati, df_scarti, df_summary = process_file(uploaded_file)
        
                if not df_dati.empty:
                    st.subheader("✅ Record Validi")
                    st.dataframe(df_dati.head(50))
        
                if not df_scarti.empty:
                    st.subheader("⚠️ Record Scartati")
                    st.dataframe(df_scarti.head(50))
        
                st.subheader("📊 Riepilogo")
                st.dataframe(df_summary)
        
                # Raggruppamenti
                df_summary_by_date = df_dati.groupby('Data calc').size().reset_index(name='Totale Record')
                df_summary_by_dest = df_dati.groupby('Destinatario - Ragione Sociale').size().reset_index(name='Totale Record')
        
                # Grafico settimanale robusto
                st.subheader("📈 Grafico settimanale per Data calc (con date)")
                try:
                    df_dati['Data calc parsed'] = pd.to_datetime(df_dati['Data calc'], format="%d/%m/%Y", errors='coerce')
                    df_week = df_dati.dropna(subset=['Data calc parsed'])
                    df_week['Week Start'] = df_week['Data calc parsed'].dt.to_period('W').apply(lambda r: r.start_time.strftime('%d/%m/%Y'))
                    df_week['Week End'] = df_week['Data calc parsed'].dt.to_period('W').apply(lambda r: r.end_time.strftime('%d/%m/%Y'))
                    df_week['Settimana'] = df_week['Week Start'] + " - " + df_week['Week End']
                    df_summary_by_week = df_week.groupby('Settimana').size().reset_index(name='Totale Record')
                    st.bar_chart(df_summary_by_week.set_index('Settimana'))
                except Exception as e:
                    st.error(f"Errore nella generazione del grafico settimanale: {e}")
        
                st.subheader("🏢 Grafico per Destinatario - Ragione Sociale (Top 20)")
                top_dest = df_summary_by_dest.sort_values(by='Totale Record', ascending=False).head(20)
                st.bar_chart(top_dest.set_index('Destinatario - Ragione Sociale'))
        
        # Calcolo avanzato: pulizia e conversioni numeriche
                df_dati['Importo Contrassegno (float)'] = df_dati['Importo Contrassegno'].str.replace(',', '.', regex=False).astype(float, errors='ignore')
                df_dati['Peso in Kg (float)'] = df_dati['Peso in Kg'].str.replace(',', '.', regex=False).astype(float, errors='ignore')
        
                # Top 10 destinatari per importo contrassegno
                st.subheader("💰 Top 10 Destinatari per Importo Contrassegno")
                top_contrassegno = df_dati.groupby('Destinatario - Ragione Sociale')['Importo Contrassegno (float)'].sum(numeric_only=True).reset_index()
                top_contrassegno = top_contrassegno.sort_values(by='Importo Contrassegno (float)', ascending=False).head(10)
                st.bar_chart(top_contrassegno.set_index('Destinatario - Ragione Sociale'))
        
                # Top 10 destinatari per peso totale
                st.subheader("⚖️ Top 10 Destinatari per Peso Totale")
                top_peso = df_dati.groupby('Destinatario - Ragione Sociale')['Peso in Kg (float)'].sum(numeric_only=True).reset_index()
                top_peso = top_peso.sort_values(by='Peso in Kg (float)', ascending=False).head(10)
                st.bar_chart(top_peso.set_index('Destinatario - Ragione Sociale'))
        
                # Trend settimanale di peso e contrassegno
                st.subheader("📈 Trend Settimanale - Peso e Contrassegno")
                if 'Data calc parsed' in df_week.columns:
                    df_trend = df_week.copy()
                    df_trend['Peso (float)'] = df_trend['Peso in Kg'].str.replace(',', '.', regex=False).astype(float, errors='ignore')
                    df_trend['Contrassegno (float)'] = df_trend['Importo Contrassegno'].str.replace(',', '.', regex=False).astype(float, errors='ignore')
                    df_trend_grouped = df_trend.groupby('Settimana')[['Peso (float)', 'Contrassegno (float)']].sum().reset_index()
                    st.line_chart(df_trend_grouped.set_index('Settimana'))
        
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    if not df_dati.empty:
                        df_dati.to_excel(writer, index=False, sheet_name="Dati_Tra")
                    if not df_scarti.empty:
                        df_scarti.to_excel(writer, index=False, sheet_name="Scarti_Tra")
                    df_summary.to_excel(writer, index=False, sheet_name="Riepilogo_Tra")
                    df_summary_by_date.to_excel(writer, index=False, sheet_name="Riep_DataCalc_Tra")
                    df_summary_by_dest.to_excel(writer, index=False, sheet_name="Riep_Destinatario_Tra")
                output.seek(0)
        
                st.download_button(
                    label="📥 Scarica Excel",
                    data=output.getvalue(),
                    file_name=f"export_tracciato_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
    
    elif menu == "Importa CSV":
        st.header("Importazione CSV")
        
        import io
        import zipfile
        from pathlib import Path
        from datetime import datetime
        from collections import defaultdict
        
        
        st.title("📄 Convertitore Fatture CSV/ZIP in Excel per Maroil")
        st.markdown("""
        ✅ Carica un `.zip` con file `.csv` separati da `;`  
        ✅ Salva in Excel:
        - Dati validi → foglio **Dati Fat**
        - Scarti → foglio **Scarti**
        - Conteggio → fogli **Riepilogo Fat** e **Scarti Fat**
        - Rimozione virgolette (`"`) e intestazione unica
        """)
        
        uploaded_file = st.file_uploader("Carica un file ZIP con CSV", type="zip")
        
        if uploaded_file:
            dati_records = []
            scarti_records = []
            header_fields = None
        
            try:
                with zipfile.ZipFile(uploaded_file) as archive:
                    csv_list = [name for name in archive.namelist() if name.endswith(".csv")]
                    if not csv_list:
                        st.warning("⚠️ Nessun file CSV trovato nello ZIP.")
        
                    for csv_name in csv_list:
                        try:
                            with archive.open(csv_name) as csv_file:
                                raw_lines = csv_file.read().decode("utf-8", errors="ignore").splitlines()
                                if not raw_lines:
                                    continue
        
                                if header_fields is None:
                                    header_fields = [h.strip().replace('"', '') for h in raw_lines[0].split(";")]
                                num_colonne_attese = len(header_fields)
        
                                for i, line in enumerate(raw_lines[1:], start=2):
                                    line_clean = line.replace('"', '')
                                    cols = line_clean.split(";")
                                    row_info = {"File origine Fat": csv_name, "Riga CSV": i}
              
                                    if len(cols) == num_colonne_attese:
                                        record = {**row_info, **{header_fields[j]: cols[j].strip() for j in range(num_colonne_attese)}}
                                        if "DT_FATT" in record and record["DT_FATT"].isdigit() and len(record["DT_FATT"]) == 8:
                                           record["Data Calc CSV"] = f"{record['DT_FATT'][6:8]}/{record['DT_FATT'][4:6]}/{record['DT_FATT'][0:4]}"
                                        else:
                                            record["Data Calc CSV"] = ""
                                        dati_records.append(record)
                                    else:
                                        scarti_records.append({**row_info, "Contenuto": line, "Errore": "Colonne non conformi"})
       
                            st.success(f"✅ Elaborato: {csv_name}")
                        except Exception as e:
                            st.error(f"Errore nel file {csv_name}: {e}")
            except Exception as e:
                st.error(f"Errore nell’apertura dello ZIP: {e}")
        
            # Solo 3 righe per anteprima veloce
            if dati_records:
                st.subheader("👁️ Anteprima Dati_Fat")
                st.dataframe(pd.DataFrame(dati_records[:5]))
            if scarti_records:
                st.subheader("🚫 Anteprima Scarti")
                st.dataframe(pd.DataFrame(scarti_records[:5]))
        
            if dati_records or scarti_records:
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
                    # Dati Fat
                    pd.DataFrame(dati_records).to_excel(writer, sheet_name="Dati_Fat", index=False)
                    # Scarti
                    # Riepilogo Fat
                    df_riepilogo = pd.DataFrame(dati_records).groupby("File origine Fat").size().reset_index(name="Record Validi")
                    df_riepilogo.to_excel(writer, sheet_name="Riepilogo_Fat", index=False)
        
                    # Scarti Fat
                    df_scarti_fat = pd.DataFrame([
                        {
                            "File origine Fat": r.get("File origine Fat", ""),
                            "Riga Fat": r.get("Riga Fat", ""),
                            "Errore": r.get("Errore", "Errore sconosciuto")
                        } for r in scarti_records
                    ])
                    df_scarti_fat.to_excel(writer, sheet_name="Scarti_Fat", index=False)
        
                st.download_button(
                    label="📥 Scarica Excel completo",
                    data=buffer.getvalue(),
                    file_name=f"export_fatture_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
    
    elif menu == "Confronta":
        
        st.title("✅ Verifica coerenza tra Tracciato e Fattura")
        
        file1 = st.file_uploader("Carica il file Excel del Tracciato", type=["xlsx"])
        file2 = st.file_uploader("Carica il file Excel della Fattura", type=["xlsx"])
        
        def normalizza_colli(val):
            try:
                return int(str(val).lstrip("0"))
            except:
                return None
        
        def normalizza_peso(val):
            try:
                return float(str(val).replace(",", "."))
            except:
                return None
        
        if file1 and file2:
            try:
                tracciato = pd.read_excel(file1, sheet_name="Dati_Tra")
                fattura = pd.read_excel(file2, sheet_name="Dati_Fat")
        
                tracciato.columns = tracciato.columns.str.strip().str.lower()
                fattura.columns = fattura.columns.str.strip().str.lower()
        
                # 🔧 Uniforma i tipi delle colonne chiave a stringa
                tracciato["bolla / fattura numero"] = tracciato["bolla / fattura numero"].astype(str).str.strip()
                fattura["rif_cli"] = fattura["rif_cli"].astype(str).str.strip()
        
                trac = tracciato[["bolla / fattura numero", "totale colli", "peso in kg"]].copy()
                fat = fattura[["rif_cli", "num_coll", "tot_peso"]].copy()
        
                trac["totale colli norm"] = trac["totale colli"].apply(normalizza_colli)
                fat["totale colli norm"] = fat["num_coll"].apply(normalizza_colli)
        
                trac["peso norm"] = trac["peso in kg"].apply(normalizza_peso)
                fat["peso norm"] = fat["tot_peso"].apply(normalizza_peso)
        
                merged = pd.merge(
                    fat, trac,
                    how="left",
                    left_on="rif_cli",
                    right_on="bolla / fattura numero",
                    suffixes=("_fat", "_trac")
                )
        
                # ✅ Corretto: riferimenti ai nomi giusti dopo il merge
                merged["ok_colli"] = merged["totale colli norm_fat"] == merged["totale colli norm_trac"]
                merged["ok_peso"] = merged["peso norm_fat"].round(2) == merged["peso norm_trac"].round(2)
        
                errori = merged[~(merged["ok_colli"] & merged["ok_peso"])]
        
                st.subheader("🔎 Risultati Confronto")
                st.write(f"Totale righe confrontate: {len(merged)}")
                st.write(f"✅ Righe corrette: {(merged['ok_colli'] & merged['ok_peso']).sum()}")
                st.write(f"❌ Righe con errori: {len(errori)}")
        
                st.subheader("❌ Dettaglio Errori")
                if errori.empty:
                    st.success("Tutti i record sono coerenti!")
                else:
                    st.dataframe(errori)
        
                import io
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                    merged.to_excel(writer, sheet_name="Confronto Completo", index=False)
                    errori.to_excel(writer, sheet_name="Errori", index=False)
        
                st.download_button(
                    "📥 Scarica risultati in Excel",
                    data=output.getvalue(),
                    file_name=f"verifica_fattura_tracciato_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                )
        
            except Exception as e:
                st.error(f"Errore durante l'elaborazione: {e}")
        else:
            st.info("📂 Carica entrambi i file per eseguire la validazione.")

elif menu == "Confronta Avanzato":
    st.header("🧠 Confronto Dati Avanzato")
    
    import os
    
    st.title("🧠 Confronto Avanzato tra Tracciato e Fatture")
    
    st.markdown("Carica i file CSV o Excel delle **Fatture** e del **Tracciato**, poi configura fino a 5 confronti campo per campo con logica AND/OR.")
    
    file_fatture = st.file_uploader("📄 Carica file Fatture", type=["csv", "xlsx"])
    file_tracciato = st.file_uploader("📄 Carica file Tracciato", type=["csv", "xlsx"])
    
    if file_fatture and file_tracciato:
        df_fatture = pd.read_excel(file_fatture) if file_fatture.name.endswith("xlsx") else pd.read_csv(file_fatture, sep=None, engine="python")
        df_tracciato = pd.read_excel(file_tracciato) if file_tracciato.name.endswith("xlsx") else pd.read_csv(file_tracciato, sep=None, engine="python")
    
        st.success("✅ File caricati correttamente.")
        st.markdown("### ⚙️ Configura fino a 5 confronti")
    
        config = []
        for i in range(5):
            st.markdown(f"#### Campo {i+1}")
            cols = st.columns([3, 3, 2, 2])
            with cols[0]:
                col_fatt = st.selectbox(f"Colonna Fattura {i+1}", df_fatture.columns, key=f"fatt_{i}")
            with cols[1]:
                col_trac = st.selectbox(f"Colonna Tracciato {i+1}", df_tracciato.columns, key=f"trac_{i}")
            with cols[2]:
                logic = st.selectbox("Logica", ["AND", "OR"], key=f"logica_{i}")
            with cols[3]:
                attivo = st.checkbox("Attivo", value=(i == 0), key=f"attivo_{i}")  # Primo attivo di default
    
            config.append({
                "attivo": attivo,
                "col_fattura": col_fatt,
                "col_tracciato": col_trac,
                "logica": logic
            })
    
        if st.button("▶️ Confronta"):
            errori = []
    
            for i_fattura, row_fattura in df_fatture.iterrows():
                condizioni_match = []
    
                for i_conf, conf in enumerate(config):
                    if not conf["attivo"]:
                        continue
                    val_fatt = row_fattura.get(conf["col_fattura"])
                    match_rows = df_tracciato[df_tracciato[conf["col_tracciato"]] == val_fatt]
                    condizioni_match.append((conf["logica"], match_rows))
    
                if not condizioni_match:
                    continue  # nessuna condizione attiva
    
                # Applica la logica combinata
                risultato = None
                for logic, subset in condizioni_match:
                    is_match = not subset.empty
                    if risultato is None:
                        risultato = is_match
                    elif logic == "AND":
                        risultato = risultato and is_match
                    elif logic == "OR":
                        risultato = risultato or is_match
    
                if not risultato:
                    errori.append({
                        "errore": "Nessuna corrispondenza coerente trovata con le condizioni configurate",
                        "riga_fattura": i_fattura + 1,
                        "riga_tracciato": i_tracciato + 1
                    })

    
            df_errori = pd.DataFrame(errori)
            output_path = "output_confronto_avanzato.xlsx"
    
            righe_fattura = []
            righe_tracciato = []
    
            for index, row in df_errori.iterrows():
                riga_fattura_num = row.get('riga_fattura', '')
                descrizione = row.get('errore', '')
    
                if riga_fattura_num and str(riga_fattura_num).isdigit():
                    idx = int(riga_fattura_num) - 1
                    if 0 <= idx < len(df_fatture):
                        riga_fattura = df_fatture.iloc[idx].copy()
                        riga_fattura['Riga Originale'] = riga_fattura_num
                        riga_fattura['Errore'] = descrizione
                        righe_fattura.append(riga_fattura)
    
            df_errori_fattura = pd.DataFrame(righe_fattura)
    
            with pd.ExcelWriter(output_path, engine='openpyxl', mode='w') as writer:
                if not df_errori_fattura.empty:
                    df_errori_fattura.to_excel(writer, sheet_name='Errori Fattura', index=False)
    
            with open(output_path, "rb") as f:
                st.download_button("⬇️ Scarica Report Errori", data=f, file_name="confronto_avanzato.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    else:
        st.info("📥 Carica entrambi i file per procedere al confronto.")