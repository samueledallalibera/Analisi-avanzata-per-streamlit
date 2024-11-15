import os
import zipfile
import xml.etree.ElementTree as ET
import pandas as pd
import streamlit as st
import io
import tempfile
import shutil

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
def parse_xml_file(xml_file_path):
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
        if line_data:
            line_items.append(line_data)

    all_data = []
    combined_data = {**header_data, **general_data, **riepilogo_dati}

    if descrizioni:
        combined_data["Descrizione"] = " | ".join(descrizioni)
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

# Funzione per estrarre e processare i file .zip
def estrai_zip(file):
    # Crea una cartella temporanea per estrarre il contenuto
    temp_dir = tempfile.mkdtemp()

    # Estrai il contenuto del file zip
    with zipfile.ZipFile(file, "r") as zip_ref:
        zip_ref.extractall(temp_dir)
    
    st.write(f"File ZIP estratto in: {temp_dir}")
    return temp_dir

# Funzione per rinominare i file XML
def rinomina_file(xml_folder_path, extracted_data_df):
    renamed_folder = tempfile.mkdtemp()
    for filename in os.listdir(xml_folder_path):
        if filename.endswith('.xml'):
            xml_file_path = os.path.join(xml_folder_path, filename)
            # Estrai i dati del file XML
            data = extract_required_data_from_xml(xml_file_path)
            # Crea un nuovo nome per il file
            new_filename = f"{data.get('Numero', 'sconosciuto')}_{data.get('Denominazione', 'senza_nome')}.xml"
            new_file_path = os.path.join(renamed_folder, new_filename)
            # Copia il file con il nuovo nome
            shutil.copy(xml_file_path, new_file_path)
            st.write(f"File {filename} rinominato in {new_filename}")
    return renamed_folder

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
    return all_data_df, extracted_folder

# Caricamento del file ZIP
uploaded_zip = st.file_uploader("Carica il file ZIP contenente i file XML", type=["zip"])

if uploaded_zip:
    # Processa i file XML contenuti nel file ZIP
    extracted_data_df, extracted_folder = process_all_files_from_zip(uploaded_zip)

    # Rinominare i file XML estratti
    renamed_folder = rinomina_file(extracted_folder, extracted_data_df)

    # Creazione del buffer per il file Excel
    output = io.BytesIO()
    extracted_data_df.to_excel(output, index=False)
    output.seek(0)

    # Pulsante per il download dell'Excel
    st.download_button(
        label="Scarica il file Excel",
        data=output,
        file_name="fattura_dati_combinati_selezionati.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # Archiviazione dei file XML rinominati in un archivio ZIP
    zip_filename = tempfile.mktemp(suffix='.zip')
    shutil.make_archive(zip_filename.replace('.zip', ''), 'zip', renamed_folder)

    # Pulsante per il download del file ZIP rinominato
    with open(zip_filename, 'rb') as f:
        st.download_button(
            label="Scarica i file XML rinominati",
            data=f,
            file_name="fatture_rinominati.zip",
            mime="application/zip"
        )
