import os
import zipfile
import xml.etree.ElementTree as ET
import pandas as pd
import streamlit as st
import io
import tempfile

# Funzione gestione errori
def gestisci_errore_parsing(filename, errore):
    st.write(f"Errore nel file {filename}: {errore}. Passo al file successivo.")

# Funzione di esplorazione ricorsiva per il parsing dei dati
def parse_element(element, parsed_data, parent_tag=""):
    for child in element:
        tag_name = f"{parent_tag}/{child.tag.split('}')[-1]}" if parent_tag else child.tag.split('}')[-1]
        
        if list(child):  # Se ha figli, chiamata ricorsiva
            parse_element(child, parsed_data, tag_name)
        else:  # Altrimenti, aggiunge il testo alla struttura dei dati
            parsed_data[tag_name] = child.text

# Funzione per estrarre e parsare il file XML
def parse_xml_file(xml_file_path, includi_dettaglio_linee=True):
    tree = ET.parse(xml_file_path)
    root = tree.getroot()

    # Parsing dei dati generali della fattura
    header_data = {}
    header = root.find(".//FatturaElettronicaHeader")
    if header is not None:
        parse_element(header, header_data)

    general_data = {}
    dati_generali = root.find(".//FatturaElettronicaBody//DatiGenerali//DatiGeneraliDocumento")
    if dati_generali is not None:
        parse_element(dati_generali, general_data)

    riepilogo_dati = {}
    riepiloghi = root.findall(".//FatturaElettronicaBody//DatiBeniServizi//DatiRiepilogo")
    for riepilogo in riepiloghi:
        parse_element(riepilogo, riepilogo_dati)

    line_items = []
    descrizioni = []
    lines = root.findall(".//FatturaElettronicaBody//DettaglioLinee")
    for line in lines:
        line_data = {}
        parse_element(line, line_data)
        if "Descrizione" in line_data:
            descrizioni.append(line_data["Descrizione"])
        if includi_dettaglio_linee:
            line_items.append(line_data)

    all_data = []
    combined_data = {**header_data, **general_data, **riepilogo_dati}

    if not includi_dettaglio_linee and descrizioni:
        combined_data["Descrizione"] = " | ".join(descrizioni)
        all_data.append(combined_data)
    elif line_items:
        first_line_data = line_items[0]
        combined_data = {**combined_data, **first_line_data}
        all_data.append(combined_data)
        for line_data in line_items[1:]:
            line_row = {**{key: None for key in combined_data.keys()}, **line_data}
            all_data.append(line_row)
    else:
        all_data.append(combined_data)

    return all_data

# Funzione per estrarre i dati richiesti dal file XML
def extract_required_data_from_xml(xml_file_path):
    tree = ET.parse(xml_file_path)
    root = tree.getroot()

    # Dizionario per memorizzare i dati estratti
    extracted_data = {}

    # Estrazione dei dati richiesti dalla struttura XML
    header = root.find(".//CedentePrestatore/DatiAnagrafici/Anagrafica")
    if header is not None:
        extracted_data["Denominazione"] = header.find(".//Denominazione").text if header.find(".//Denominazione") is not None else None

    general_data = root.find(".//FatturaElettronicaBody//DatiGenerali//DatiGeneraliDocumento")
    if general_data is not None:
        extracted_data["Data"] = general_data.find(".//Data").text if general_data.find(".//Data") is not None else None
        extracted_data["Numero"] = general_data.find(".//Numero").text if general_data.find(".//Numero") is not None else None

    return extracted_data

# Funzione per decodificare e convertire i file .p7m in .xml
def converti_p7m_in_xml(fe_path):
    if not os.path.exists(fe_path):
        st.error("Il percorso specificato non esiste!")
        return

    file = os.listdir(fe_path)

    for x in range(len(file)):
        full_file_path = os.path.join(fe_path, file[x])
        if ".p7m" in file[x]:
            xml_output_path = os.path.join(fe_path, f"{x}.xml")
            os.system(f'openssl smime -verify -noverify -in "{full_file_path}" -inform DER -out "{xml_output_path}"')
            os.remove(full_file_path)  # Rimuovi il file .p7m originale
            st.write(f"File {file[x]} convertito in XML.")

# Funzione per estrarre e processare i file .zip
def estrai_zip(file):
    # Crea una cartella temporanea per estrarre il contenuto
    temp_dir = tempfile.mkdtemp()

    # Estrai il contenuto del file zip
    with zipfile.ZipFile(file, "r") as zip_ref:
        zip_ref.extractall(temp_dir)
    
    st.write(f"File ZIP estratto in: {temp_dir}")
    return temp_dir

# Funzione per processare tutti i file XML estratti
def process_all_files_from_zip(zip_file):
    # Estrai i file dal .zip
    extracted_folder = estrai_zip(zip_file)
    
    all_data_combined = []

    # Processa i file XML estratti
    for filename in os.listdir(extracted_folder):
        if filename.endswith('.xml'):
            xml_file_path = os.path.join(extracted_folder, filename)
            st.write(f"Elaborando il file: {filename}")
            try:
                file_data = parse_xml_file(xml_file_path)
                all_data_combined.extend(file_data)
            except ET.ParseError as e:
                gestisci_errore_parsing(filename, e)

    all_data_df = pd.DataFrame(all_data_combined)
    return all_data_df

# Caricamento del file ZIP
uploaded_zip = st.file_uploader("Carica il file ZIP contenente i file XML", type=["zip"])

if uploaded_zip:
    # Processa i file XML contenuti nel file ZIP
    extracted_data_df = process_all_files_from_zip(uploaded_zip)

    # Creazione del buffer per il file Excel
    output = io.BytesIO()
    extracted_data_df.to_excel(output, index=False)
    output.seek(0)

    # Pulsante per il download
    st.download_button(
        label="Scarica il file Excel",
        data=output,
        file_name="fattura_dati_combinati_selezionati.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
