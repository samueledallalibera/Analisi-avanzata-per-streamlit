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

# Funzione per estrarre e parsare il file XML con possibilità di includere o meno il dettaglio delle linee
def parse_xml_file(xml_file_path, includi_dettaglio_linee=True):
    """
    Funzione che esegue il parsing di un file XML contenente una fattura elettronica.
    :param xml_file_path: Percorso del file XML da parsare.
    :param includi_dettaglio_linee: Se True, include anche i dettagli delle linee nella fattura.
    :return: Una lista di dizionari con i dati estratti dalla fattura.
    """
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

    # Gestione delle linee della fattura
    line_items = []
    descrizioni = []
    if includi_dettaglio_linee:
        lines = root.findall(".//FatturaElettronicaBody//DettaglioLinee")
        for line in lines:
            line_data = {}
            parse_element(line, line_data)
            if "Descrizione" in line_data:
                descrizioni.append(line_data["Descrizione"])
            line_items.append(line_data)

    # Combinazione dei dati estratti
    all_data = []
    combined_data = {**header_data, **general_data, **riepilogo_dati}

    # Se includiamo il dettaglio delle linee, lo aggiungiamo alla combinazione
    if line_items:
        first_line_data = line_items[0]
        combined_data = {**combined_data, **first_line_data}
        all_data.append(combined_data)
        for line_data in line_items[1:]:
            line_row = {**{key: None for key in combined_data.keys()}, **line_data}
            all_data.append(line_row)
    else:
        # Se non includiamo il dettaglio delle linee ma abbiamo descrizioni, le aggiungiamo come concatenazione
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

# Funzione per iterare su più file e compilare un unico DataFrame
def process_all_files(xml_folder_path, includi_dettaglio_linee=True):
    all_data_combined = []

    for filename in os.listdir(xml_folder_path):
        if filename.endswith('.xml'):
            xml_file_path = os.path.join(xml_folder_path, filename)
            print(f"Elaborando il file: {filename}")
            try:
                file_data = parse_xml_file(xml_file_path, includi_dettaglio_linee)
                all_data_combined.extend(file_data)
            except ET.ParseError as e:
                gestisci_errore_parsing(filename, e)

    all_data_df = pd.DataFrame(all_data_combined)
    return all_data_df

# Funzione per estrarre e processare i file .zip
def estrai_zip(file):
    # Crea una cartella visibile nella sessione corrente per estrarre il contenuto
    output_dir = "estratti_zip"
    
    # Se la cartella esiste già, la svuotiamo
    if os.path.exists(output_dir):
        for filename in os.listdir(output_dir):
            file_path = os.path.join(output_dir, filename)
            if os.path.isfile(file_path):
                os.remove(file_path)
    
    # Crea la cartella se non esiste
    os.makedirs(output_dir, exist_ok=True)

    # Estrai il contenuto del file zip
    with zipfile.ZipFile(file, "r") as zip_ref:
        zip_ref.extractall(output_dir)
    
    # Visualizza i file estratti
    file_names = os.listdir(output_dir)
    st.write(f"File ZIP estratto in: {output_dir}")
    st.write(f"File estratti: {file_names}")
    
    # Restituisci il percorso della cartella di estrazione
    return output_dir

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

    # Verifica se i dati sono stati estratti correttamente
    if extracted_data_df.empty:
        st.error("Nessun dato trovato nei file XML. Verifica i file contenuti nel file ZIP.")
    else:
        st.write(f"DataFrame estratto con successo: {extracted_data_df.shape[0]} righe.")

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

