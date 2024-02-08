import os
import PyPDF2
from docx import Document
import spacy
import sys
from collections import Counter
import json

nlp = spacy.load("custom_model")

def check_ent_folder_name(folder_path, ent):
    ent_to_list = {"LOC":"comuni_extended_list.json", "COMM":"Committente_list_extended.json", "TL": "Tipologia Lavori_list.json", "OPERA": "Opera_list.json", "T.INT": "Tipologia Intervento_list.json", "T.INC": "Tipologia Incarico_list.json", "STR": "struttura_list.json"}
    list_name = ent_to_list[ent]
    with open(f"Lists/{list_name}", "r") as f:
        l = json.load(f)
    folder_name = os.path.basename(folder_path)
    folder_name = folder_name.lower()
    folder_name = folder_name.split()
    for element in l:
        if element.lower() in folder_name:
            return element
    return False

def activation(folder_path):
    folder_name = os.path.basename(folder_path)
    commessa = folder_name[:8]
    entities_to_scan = []
    final_result = {}

    with open("commesse_azioni.json", "r") as f:
        actions = json.load(f)

    if "ID Committente" in actions[commessa]:
        if check_ent_folder_name(folder_path, "COMM") == False:
            entities_to_scan.append("COMM")
        else:
            final_result["COMM"] = check_ent_folder_name(folder_path, "COMM")

    elif "Ubicazione Opera" in actions[commessa]:
        if check_ent_folder_name(folder_path, "LOC") == False:
            entities_to_scan.append("LOC")
        else:
            final_result["LOC"] = check_ent_folder_name(folder_path, "LOC")
          
    elif "Tipologia Lavori" in actions[commessa]:
         if check_ent_folder_name(folder_path, "TL") == False:
            entities_to_scan.append("TL")
         else:
            final_result["TL"] = check_ent_folder_name(folder_path, "TL")

    # elif "Descrizione Breve" in actions[commessa]:
    #     return None

    elif "Opera" in actions[commessa]:
        if check_ent_folder_name(folder_path, "OPERA") == False:
            entities_to_scan.append("OPERA")
        else:
            final_result["OPERA"] = check_ent_folder_name(folder_path, "OPERA")

    # elif "Descrizione Estesa" in actions[commessa]:
    #     return None

    elif "Tipologia Incarico" in actions[commessa]:
        if check_ent_folder_name(folder_path, "T.INC") == False:
            entities_to_scan.append("T.INC")
        else:
            final_result["T.INC"] = check_ent_folder_name(folder_path, "T.INC")

    elif "Tipologia Intervento" in actions[commessa]:
        if check_ent_folder_name(folder_path, "T.INT") == False:
            entities_to_scan.append("T.INT")
        else:
            final_result["T.INT"] = check_ent_folder_name(folder_path, "T.INT")

    elif "Struttura" in actions[commessa]:
        if check_ent_folder_name(folder_path, "STR") == False:
            entities_to_scan.append("STR")
        else:
            final_result["STR"] = check_ent_folder_name(folder_path, "STR")

    return (entities_to_scan, final_result)

def scan(folder_path):
    entities_to_scan, final_result = activation(folder_path)
    for root, _, files in os.walk(folder_path):
        for file in files:
            file_path = os.path.join(root, file)
            file_extension = os.path.splitext(file)[1].lower()

            if file_extension == ".pdf":
                result = scan_pdf(file_path, entities_to_scan)
            elif file_extension == ".docx":
                try:
                    result = scan_word_doc(file_path, entities_to_scan)
                except Exception as e:
                    print(f"Error opening Word document: {file_path} - {str(e)}")
            
            for ents in result.keys():
                final_result[ents] = check(ents, result[ents], 4)

            return final_result


def extract_text_from_pdf(file_path):
    with open(file_path, 'rb') as file:
        pdf_reader = PyPDF2.PdfReader(file)
        text = ''
        for page in pdf_reader.pages:
            text += page.extract_text()
    return text

def extract_text_from_docx(file_path):
    doc = Document(file_path)
    text = ''
    for paragraph in doc.paragraphs:
        text += paragraph.text + '\n'
    return text

def scan_pdf(file_path, entities_to_scan):
    text = extract_text_from_pdf(file_path)
    result = {ent:[] for ent in entities_to_scan}
    # KEEP CAPITAL LETTERS TO SCAN FOR LOCATIONS
    if "LOC" in entities_to_scan:
        doc = nlp(text)
        for ent in doc.ents:
            if ent.label_== "LOC":
                result["LOC"].append(ent.text.encode(sys.stdout.encoding, errors='replace').decode(sys.stdout.encoding))
    # LOWER DOWN THE LETTERS TO IMPROVE THE EFFICIENCY
    text = text.lower()
    doc = nlp(text)
    for ent in doc.ents:
        if ent.label_ in result.keys and ent.label_ != "LOC":
            result[ent.label_].append(ent.text.encode(sys.stdout.encoding, errors='replace').decode(sys.stdout.encoding))
    return result

def scan_word_doc(file_path, entities_to_scan):
    text = extract_text_from_docx(file_path)
    result = {ent:[] for ent in entities_to_scan}
    # KEEP CAPITAL LETTERS TO SCAN FOR LOCATIONS
    if "LOC" in entities_to_scan:
        doc = nlp(text)
        for ent in doc.ents:
            if ent.label_== "LOC":
                result["LOC"].append(ent.text.encode(sys.stdout.encoding, errors='replace').decode(sys.stdout.encoding))
    # LOWER DOWN THE LETTERS TO IMPROVE THE EFFICIENCY
    text = text.lower()
    doc = nlp(text)
    for ent in doc.ents:
        if ent.label_ in result.keys and ent.label_ != "LOC":
            result[ent.label_].append(ent.text.encode(sys.stdout.encoding, errors='replace').decode(sys.stdout.encoding))
    return result

def check(ent, lst, gap):
    counter = Counter(lst)
    if len(counter) < 3:
        print(f"La lista delle entità \"{ent}\" ha troppi pochi elementi: \n {counter.most_common(3)}")
        return "NONE"
    most_common = counter.most_common(3)
    first, count1 = most_common[0]
    second, count2 = most_common[1]
    third, count3 = most_common[2]
    if count1 - count2 < gap:
        print(f"La selezione automatica del campo \"{ent}\" non è affidabile, ci serve un aiuto. ") #trasformare ent nel nome effettivo del campo
        print(f"Questi sono i risultati della ricerca: {most_common} \n Quale dei tre è quella giusta?")
        inference = int(input(f"Digita un numero da 1 a 4. \n 1. {first}     2. {second}     3. {third}     4. Nessuna delle tre"))
        if inference == 1:
            return first
        elif inference == 2:
            return second
        elif inference == 3:
            return third
        elif inference == 4:
            # scegliere dalla lista?
            result = str(input("Dicci quale è quella giusta: "))
            #spelling check
            # add result to the list of the entity if not there yet
            return result

# Usage
directory_path = r""
print(scan(directory_path))


#def find_most_frequent_element(lst):
#    counter = Counter(lst)
#    most_common = counter.most_common(3)
#    first, count1 = most_common[0]
#    second, count2 = most_common[1]
#    print(f"The three most common locations are {most_common}") # check with more folders how it's reliable
#    # # if count1 - count2 <= 5 or first == "Avezzano" or second == "Avezzano":
#    # #     check() # give a check through an input
#    print(f"The one selected is {selected}")
#    return most_common[0][0]
#
#print(find_most_frequent_element(scanned_locations))

