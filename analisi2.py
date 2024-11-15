import os
import xml.etree.ElementTree as ET
import pandas as pd

# Funzione gestione errori
def gestisci_errore_parsing(filename, errore):
    print(f"Errore nel file {filename}: {errore}. Passo al file successivo.")

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

# Funzione per unire i dati nella forma "Denominazione FT Numero del Data"
# Limita la denominazione a 20 caratteri
def unisci_dati(df):
    df['Dati_Uniti'] = df.apply(lambda row: f"{(row['Denominazione'][:20] if row['Denominazione'] and len(row['Denominazione']) > 20 else row['Denominazione'])} FT {row['Numero']} del {row['Data']}", axis=1)
    return df[['Dati_Uniti', 'FileName']]

# Funzione per aggiornare il file Excel senza duplicati
def aggiorna_file_excel(output_path, nuovi_dati_df):
    if os.path.exists(output_path):
        dati_esistenti_df = pd.read_excel(output_path)
        df_combinato = pd.concat([dati_esistenti_df, nuovi_dati_df], ignore_index=True)
        df_combinato = df_combinato.drop_duplicates()
    else:
        df_combinato = nuovi_dati_df
    
    df_combinato.to_excel(output_path, index=False)
    print(f"Il file Excel è stato aggiornato o creato in '{output_path}'.")

# Funzione per decodificare e convertire i file .p7m in .xml
def converti_p7m_in_xml(fe_path):
    file = os.listdir(fe_path)

    for x in range(len(file)):
        full_file_path = os.path.join(fe_path, file[x])
        if ".p7m" in file[x]:
            xml_output_path = os.path.join(fe_path, f"{x}.xml")
            os.system(f'openssl smime -verify -noverify -in "{full_file_path}" -inform DER -out "{xml_output_path}"')
            os.remove(full_file_path)  # Rimuovi il file .p7m originale
            print(f"File {file[x]} convertito in XML.")

# Elenco delle colonne di default
colonne_default = [
    "CedentePrestatore/DatiAnagrafici/IdFiscaleIVA/IdPaese",
    "CedentePrestatore/DatiAnagrafici/IdFiscaleIVA/IdCodice",
    "CedentePrestatore/DatiAnagrafici/Anagrafica/Denominazione",
    "CedentePrestatore/DatiAnagrafici/RegimeFiscale",
    "CedentePrestatore/Sede/Indirizzo",
    "CedentePrestatore/Sede/NumeroCivico",
    "CedentePrestatore/Sede/CAP",
    "CedentePrestatore/Sede/Comune",
    "TipoDocumento",
    "Data",
    "Numero",
    "ImportoTotaleDocumento",
    "AliquotaIVA",
    "ImponibileImporto",
    "Imposta",
    "Descrizione",
    "PrezzoTotale"
]

# Funzione per rinominare i file con i dati uniti
def rinomina_file(xml_folder_path, df):
    for _, row in df.iterrows():
        old_file_path = os.path.join(xml_folder_path, row['FileName'])
        if row['Dati_Uniti']:  # Controlla che i dati uniti non siano vuoti
            # Creazione del nuovo nome del file, rimuovendo caratteri non validi per i nomi file
            new_file_name = "".join(c if c.isalnum() or c in " .-_()" else "_" for c in row['Dati_Uniti']) + ".xml"
            new_file_path = os.path.join(xml_folder_path, new_file_name)
            try:
                os.rename(old_file_path, new_file_path)
                print(f"Rinominato: {row['FileName']} -> {new_file_name}")
            except OSError as e:
                print(f"Errore nella rinomina di {row['FileName']}: {e}")

# Input percorso cartella XML
xml_folder_path = input("Inserisci il percorso della cartella contenente i file XML (converti prima se ci sono file .p7m): ")

# Chiede se includere il dettaglio delle linee
includi_dettaglio_linee = input("Vuoi includere il dettaglio delle linee? (sì/no): ").strip().lower() == 'sì'

# Se ci sono file .p7m, li converte in XML
converti_p7m_in_xml(xml_folder_path)

# Funzione per processare tutti i file XML nella cartella
def process_all_xml_files(xml_folder_path):
    all_extracted_data = []
    file_names = []  # Per salvare i nomi dei file

    for filename in os.listdir(xml_folder_path):
        if filename.endswith('.xml'):
            xml_file_path = os.path.join(xml_folder_path, filename)
            print(f"Elaborando il file: {filename}")
            try:
                # Estrai i dati richiesti dal file XML
                file_data = extract_required_data_from_xml(xml_file_path)
                file_data["FileName"] = filename  # Salva il nome originale del file
                all_extracted_data.append(file_data)
                file_names.append(filename)
            except ET.ParseError as e:
                print(f"Errore nel parsing del file {filename}: {e}. Passo al file successivo.")

    # Creazione di un DataFrame con i dati estratti
    return pd.DataFrame(all_extracted_data)

# Estrazione dei dati e creazione del DataFrame
extracted_data_df = process_all_xml_files(xml_folder_path)

# Unione dei dati nella forma desiderata
unified_data_df = unisci_dati(extracted_data_df)

# Rinominare i file XML
rinomina_file(xml_folder_path, unified_data_df)

# Input percorso di salvataggio
output_folder_path = input("Inserisci il percorso della cartella dove salvare il file Excel (premi invio per usare la cartella corrente): ").strip()
if not output_folder_path:
    output_folder_path = os.getcwd()
elif not os.path.exists(output_folder_path):
    print("Percorso non valido. Verrà utilizzata la cartella corrente.")
    output_folder_path = os.getcwd()

output_file_name = "fattura_dati_combinati_selezionati.xlsx"
output_path = os.path.join(output_folder_path, output_file_name)


# Parsing e creazione del DataFrame
all_data_df = process_all_files(xml_folder_path, includi_dettaglio_linee)

# Filtrare le colonne esistenti
colonne_esistenti = [col for col in colonne_default if col in all_data_df.columns]
if colonne_esistenti:
    dati_da_esportare_df = all_data_df[colonne_esistenti]
    aggiorna_file_excel(output_path, dati_da_esportare_df)
else:
    print("Nessuna colonna valida trovata per l'esportazione.")