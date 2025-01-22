import base64
from io import BytesIO
import logging
import zipfile
from docx import Document
from fastapi import APIRouter, HTTPException
import os
from fastapi.responses import FileResponse
import openpyxl
import requests
from app.services.ects_service import get_ects_for_template
from app.services.prisma_service import get_template_from_prisma
from app.services.excel_service import match_template_and_get_word, process_excel_with_template
from app.services.word_service import generate_bulletins_from_excel
from prisma import Prisma
from datetime import datetime

router = APIRouter()

def calculate_weighted_average(notes):
    """
    Calcule la moyenne pondérée des notes avec leurs coefficients.
    Format des notes : "10 (0,25)" ou "10(0.25)" ou "10"
    """
    try:
        if not notes:
            return 0

        total_weighted_sum = 0
        total_coefficients = 0

        for note_str in notes:
            if note_str is None or str(note_str).strip() == "":
                continue

            note_str = str(note_str).strip()
            
            # Cas d'une note avec coefficient
            if "(" in note_str:
                try:
                    # Extraire la note et le coefficient
                    parts = note_str.split("(")
                    note = float(parts[0].strip())
                    coef = float(parts[1].replace(")", "").replace(",", ".").strip())
                    
                    total_weighted_sum += note * coef
                    total_coefficients += coef
                except (ValueError, IndexError):
                    continue
            
            # Cas d'une note simple (coefficient 1)
            else:
                try:
                    note = float(note_str)
                    total_weighted_sum += note
                    total_coefficients += 1
                except ValueError:
                    continue

        if total_coefficients == 0:
            return 0

        # Arrondir au centième
        return round(total_weighted_sum / total_coefficients, 2)

    except Exception as e:
        logging.error(f"Erreur lors du calcul de la moyenne pondérée: {str(e)}")
        return 0
    

def calculate_single_note_average(note_str):
    """
    Calcule la moyenne pondérée pour une note avec plusieurs coefficients.
    Format: "10 (0,25) - 15 (0,25) - 10,5 (0,5)" ou "17 - 16 - 17" ou "Absent au devoir (0,25) - 11 (0,25) - 11,5 (0,5)"
    """
    try:
        if not note_str or str(note_str).strip() == "":
            return ""

        note_str = str(note_str).strip()
        
        # Si la note contient des coefficients
        if "(" in note_str:
            try:
                # Séparer les différentes notes
                notes_parts = note_str.split("-")
                total_weighted_sum = 0
                total_coefficients = 0
                
                for part in notes_parts:
                    part = part.strip()
                    # Ignorer les parties contenant "Absent au devoir"
                    if "Absent au devoir" in part:
                        continue
                        
                    if "(" in part and ")" in part:
                        # Extraire la note et le coefficient
                        note_part = part.split("(")
                        # Remplacer la virgule par un point dans la note
                        note = float(note_part[0].strip().replace(",", "."))
                        # Remplacer la virgule par un point et enlever la parenthèse dans le coefficient
                        coeff = float(note_part[1].replace(",", ".").replace(")", "").strip())
                        
                        total_weighted_sum += note * coeff
                        total_coefficients += coeff
                
                if total_coefficients > 0:
                    weighted_average = total_weighted_sum / total_coefficients
                    return f"{weighted_average:.2f}"
                return ""
                    
            except (ValueError, IndexError) as e:
                logging.error(f"Erreur lors du calcul de la moyenne pondérée: {str(e)}")
                return ""
        
        # Cas d'une note simple ou multiple sans coefficient (ex: "17 - 16 - 17")
        else:
            try:
                # Séparer les notes s'il y en a plusieurs
                notes = [float(n.strip().replace(",", ".")) for n in note_str.split("-") if n.strip() and "Absent au devoir" not in n]
                if notes:
                    # Calculer la moyenne simple (coefficient 1 pour chaque note)
                    average = sum(notes) / len(notes)
                    return f"{average:.2f}"
                return ""
            except ValueError as e:
                logging.error(f"Erreur lors du calcul de la note simple: {str(e)}")
                return ""

    except Exception as e:
        logging.error(f"Erreur lors du calcul de la note: {str(e)}")
        return ""

def get_etat(note_str: str) -> str:
    """
    Détermine l'état en fonction de la note.
    - Si note >= 10 ou vide : ""
    - Si 8 <= note < 10 : "C"
    - Si note < 8 : "R"
    """
    if not note_str or str(note_str).strip() == "":
        return ""
        
    try:
        note = float(str(note_str).replace(",", "."))
        if note >= 10:
            return ""
        elif 8 <= note < 10:
            return "C"
        else:
            return "R"
    except (ValueError, TypeError):
        return ""

def get_etat_ue(etats: list, moyenne_ue: str = "") -> str:
    """
    Détermine l'état d'une UE en fonction des états des notes et de la moyenne de l'UE.
    
    Règles:
    - "VA" si:
        * tous les états sont vides ET moyenne_ue >= 10
        * OU au moins un état "C", les autres vides, ET moyenne_ue >= 10
    - "NV" si:
        * au moins un état "R"
        * OU moyenne_ue < 8
        * OU (au moins un état "C" ET moyenne_ue < 10)
    """
    try:
        moyenne = float(moyenne_ue.replace(",", ".")) if moyenne_ue else 0
    except (ValueError, TypeError):
        moyenne = 0

    has_r = False
    has_c = False
    all_empty = True

    for etat in etats:
        if etat == "R":
            has_r = True
        elif etat == "C":
            has_c = True
        if etat != "":
            all_empty = False

    # Cas où il y a au moins un R ou moyenne < 8
    if has_r or moyenne < 8:
        return "NV"
    
    # Cas où tous les états sont vides
    if all_empty and moyenne >= 10:
        return "VA"
    
    # Cas où il y a au moins un C
    if has_c:
        return "VA" if moyenne >= 10 else "NV"
    
    # Cas par défaut
    return "NV"


def calculate_ects_weighted_average(notes, ects_values):
    """
    Calcule la moyenne pondérée par les ECTS.
    Les notes avec ECTS = 0 sont complètement ignorées dans le calcul.
    """
    try:
        total_weighted_sum = 0
        total_ects = 0

        for note_str, ects_str in zip(notes, ects_values):
            if not note_str:
                continue

            try:
                note = float(str(note_str).replace(",", "."))
                ects = int(ects_str)

                # Si la note est inférieure à 8, l'ECTS devient 0
                if note < 8:
                    ects = 0

                # Ne prendre en compte que les notes avec ECTS > 0
                if ects > 0:
                    total_weighted_sum += note * ects
                    total_ects += ects

            except (ValueError, TypeError):
                continue

        if total_ects == 0:
            return ""

        return f"{(total_weighted_sum / total_ects):.2f}"

    except Exception as e:
        logging.error(f"Erreur lors du calcul de la moyenne pondérée ECTS: {str(e)}")
        return ""
    
def clean_temp_directory(temp_dir: str):
    for filename in os.listdir(temp_dir):
        file_path = os.path.join(temp_dir, filename)
        if os.path.isfile(file_path):
            os.remove(file_path)

def clean_except_specific_file(directory: str, keep_filename: str):
    """
    Supprime tous les fichiers dans un répertoire sauf le fichier spécifié.
    """
    for filename in os.listdir(directory):
        file_path = os.path.join(directory, filename)
        if filename != keep_filename and os.path.isfile(file_path):
            try:
                os.remove(file_path)
                logging.info(f"Fichier supprimé : {file_path}")
            except Exception as e:
                logging.error(f"Erreur lors de la suppression de {file_path}: {str(e)}")


# Calculer les totaux d'ECTS pour chaque UE en tenant compte de la règle des notes < 8
def calculate_ue_ects(notes, ects_values):
    total_ects = 0
    for note_str, ects_str in zip(notes, ects_values):
        try:
            note = float(note_str) if note_str else 0
            ects = int(ects_str)
            if note >= 8:  # On ne compte les ECTS que si la note est >= 8
                total_ects += ects
        except (ValueError, TypeError):
            continue
        return total_ects
    
def get_total_etat(etat_ue1: str, etat_ue2: str, etat_ue3: str, etat_ue4: str) -> str:
    """
    Détermine l'état total en fonction des états des UE.
    
    Règles:
    - "VA" si tous les états des UE sont "VA"
    - "NV" si au moins un état d'UE est "NV"
    """
    if all(etat == "VA" for etat in [etat_ue1, etat_ue2, etat_ue3, etat_ue4]):
        return "VA"
    return "NV"
    
@router.post("/process-excel")
async def process_excel(excel_url: str, word_url: str, user_id: str):
    try:
        output_dir = "./temp"

        logging.info(f"Début du traitement avec excel_url={excel_url}, word_url={word_url}")

        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
        else:
            clean_temp_directory(output_dir)

        # Télécharger le fichier Excel source pour lire le nom du groupe
        logging.info(f"Téléchargement du fichier Excel depuis {excel_url}")
        excel_response = requests.get(excel_url)
        if excel_response.status_code != 200:
            raise HTTPException(status_code=400, detail="Impossible de télécharger le fichier Excel.")
        excel_file = BytesIO(excel_response.content)
        
        # Lire le nom du groupe depuis B2
        wb = openpyxl.load_workbook(excel_file)
        ws = wb.active
        group_name = ws["B2"].value
        
        if not group_name:
            raise HTTPException(status_code=400, detail="Nom du groupe non trouvé dans la cellule B2")
            
        # Déterminer le template à utiliser
        from app.core.template_mapping import TEMPLATE_MAPPING
        prisma_template = TEMPLATE_MAPPING.get(group_name)
        if not prisma_template:
            raise HTTPException(status_code=400, detail=f"Aucun template trouvé pour le groupe : {group_name}")

        logging.info(f"Template sélectionné pour le groupe {group_name}: {prisma_template}")

        # Récupérer le template Excel depuis Prisma
        template_excel_path = await get_template_from_prisma(prisma_template, output_dir)
        logging.info(f"Template récupéré et sauvegardé dans : {template_excel_path}")

        clean_except_specific_file(output_dir, keep_filename=os.path.basename(template_excel_path))

        logging.info("Début du traitement des données entre fichier source et template")
        result = await process_excel_with_template(excel_url, output_dir, prisma_template, user_id)
        logging.info(f"Traitement terminé. Fichier mis à jour disponible : {result['excel_path']}")

        return {"message": "Fichier traité avec succès", "excel_id": result['excel_id']}

    except Exception as e:
        logging.error(f"Erreur pendant le traitement : {str(e)}")
        raise HTTPException(status_code=400, detail=f"Erreur : {str(e)}")



@router.post("/get-word-template")
async def get_word_template_endpoint():
    try:
        db = Prisma()
        await db.connect()

        excel_path = os.path.join("./temp", "updated_excel.xlsx")
        if not os.path.exists(excel_path):
            raise ValueError("Fichier Excel mis à jour non trouvé dans ./temp")

        # Déterminer le template à utiliser en comparant avec les modèles
        template_info = await match_template_and_get_word(excel_path)
        template_name = template_info["template_name"]
        ects_template = template_info["ects_template"]
        
        logging.info(f"Utilisation du template {template_name} avec ECTS {ects_template}")

        updated_wb = openpyxl.load_workbook(excel_path)
        updated_ws = updated_wb.active

        date_du_jour = datetime.utcnow().strftime("%d/%m/%Y")

        # Mapping des colonnes selon le modèle
        if template_name == "modeleBGALT2.docx":
            ue_matieres = {
                "UE1_Title": str(updated_ws["C1"].value or ""),
                "matiere1": str(updated_ws["D1"].value or ""),
                "matiere2": str(updated_ws["E1"].value or ""),
                "matiere3": str(updated_ws["F1"].value or ""),
                "matiere4": str(updated_ws["G1"].value or ""),
                "UE2_Title": str(updated_ws["H1"].value or ""),
                "matiere5": str(updated_ws["I1"].value or ""),
                "matiere6": str(updated_ws["J1"].value or ""),
                "matiere7": str(updated_ws["K1"].value or ""),
                "UE3_Title": str(updated_ws["L1"].value or ""),
                "matiere8": str(updated_ws["M1"].value or ""),
                "UE4_Title": str(updated_ws["N1"].value or ""),
                "matiere9": str(updated_ws["O1"].value or ""),
                "matiere10": str(updated_ws["P1"].value or ""),
                "matiere11": str(updated_ws["Q1"].value or ""),
                "matiere12": str(updated_ws["R1"].value or ""),
                "matiere13": str(updated_ws["S1"].value or ""),
                "matiere14": str(updated_ws["T1"].value or ""),
                "matiere15": str(updated_ws["U1"].value or "")
            }
        else:  # modeleBGALT3.docx
            ue_matieres = {
                "UE1_Title": str(updated_ws["C1"].value or ""),
                "matiere1": str(updated_ws["D1"].value or ""),
                "matiere2": str(updated_ws["E1"].value or ""),
                "matiere3": str(updated_ws["F1"].value or ""),
                "matiere4": str(updated_ws["G1"].value or ""),
                "matiere5": str(updated_ws["H1"].value or ""),
                "UE2_Title": str(updated_ws["I1"].value or ""),
                "matiere6": str(updated_ws["J1"].value or ""),
                "matiere7": str(updated_ws["K1"].value or ""),
                "UE3_Title": str(updated_ws["L1"].value or ""),
                "matiere8": str(updated_ws["M1"].value or ""),
                "UE4_Title": str(updated_ws["N1"].value or ""),
                "matiere9": str(updated_ws["O1"].value or ""),
                "matiere10": str(updated_ws["P1"].value or ""),
                "matiere11": str(updated_ws["Q1"].value or ""),
                "matiere12": str(updated_ws["R1"].value or ""),
                "matiere13": str(updated_ws["S1"].value or "")
            }

        # Récupérer le template Word
        word_template = await db.generatedfile.find_first(
            where={
                "filename": template_name,
                "isTemplate": True
            }
        )
        
        if not word_template:
            raise ValueError(f"Template Word {template_name} non trouvé dans Prisma")

        # Récupérer les ECTS selon le template
        ects_data = await get_ects_for_template(ects_template)
        logging.info(f"ECTS data for {ects_template}: {ects_data}")
        
        if template_name == "modeleBGALT2.docx":
            required_ects = [f"ECTS{i}" for i in range(1, 16)]  # ECTS1 à ECTS15
            missing_ects = [ects for ects in required_ects if ects not in ects_data]
            if missing_ects:
                raise ValueError(f"Missing ECTS values for {template_name}: {missing_ects}")
        else:  # modeleBGALT3.docx
            required_ects = [f"ECTS{i}" for i in range(1, 14)]  # ECTS1 à ECTS13
            missing_ects = [ects for ects in required_ects if ects not in ects_data]
            if missing_ects:
                raise ValueError(f"Missing ECTS values for {template_name}: {missing_ects}")

        bulletins_dir = os.path.join("./temp", "bulletins")
        if not os.path.exists(bulletins_dir):
            os.makedirs(bulletins_dir)

        for row in range(3, updated_ws.max_row + 1):
            if not updated_ws[f"B{row}"].value:
                continue

            # Récupérer les notes selon le modèle
            if template_name == "modeleBGALT2.docx":
                ue1_notes = [
                    calculate_single_note_average(updated_ws[f"D{row}"].value),
                    calculate_single_note_average(updated_ws[f"E{row}"].value),
                    calculate_single_note_average(updated_ws[f"F{row}"].value),
                    calculate_single_note_average(updated_ws[f"G{row}"].value)
                ]
                ue2_notes = [
                    calculate_single_note_average(updated_ws[f"I{row}"].value),
                    calculate_single_note_average(updated_ws[f"J{row}"].value),
                    calculate_single_note_average(updated_ws[f"K{row}"].value)
                ]
                ue3_notes = [
                    calculate_single_note_average(updated_ws[f"M{row}"].value)
                ]
                ue4_notes = [
                    calculate_single_note_average(updated_ws[f"O{row}"].value),
                    calculate_single_note_average(updated_ws[f"P{row}"].value),
                    calculate_single_note_average(updated_ws[f"Q{row}"].value),
                    calculate_single_note_average(updated_ws[f"R{row}"].value),
                    calculate_single_note_average(updated_ws[f"S{row}"].value),
                    calculate_single_note_average(updated_ws[f"T{row}"].value),
                    calculate_single_note_average(updated_ws[f"U{row}"].value)
                ]
            else:  # modeleBGALT3.docx
                ue1_notes = [
                    calculate_single_note_average(updated_ws[f"D{row}"].value),
                    calculate_single_note_average(updated_ws[f"E{row}"].value),
                    calculate_single_note_average(updated_ws[f"F{row}"].value),
                    calculate_single_note_average(updated_ws[f"G{row}"].value),
                    calculate_single_note_average(updated_ws[f"H{row}"].value)
                ]
                ue2_notes = [
                    calculate_single_note_average(updated_ws[f"J{row}"].value),
                    calculate_single_note_average(updated_ws[f"K{row}"].value)
                ]
                ue3_notes = [
                    calculate_single_note_average(updated_ws[f"M{row}"].value)
                ]
                ue4_notes = [
                    calculate_single_note_average(updated_ws[f"O{row}"].value),
                    calculate_single_note_average(updated_ws[f"P{row}"].value),
                    calculate_single_note_average(updated_ws[f"Q{row}"].value),
                    calculate_single_note_average(updated_ws[f"R{row}"].value),
                    calculate_single_note_average(updated_ws[f"S{row}"].value)
                ]

            # Calculer les moyennes avec ECTS
            if template_name == "modeleBGALT2.docx":
                moyUE1 = calculate_ects_weighted_average(ue1_notes, [
                    ects_data["ECTS1"], ects_data["ECTS2"], ects_data["ECTS3"],
                    ects_data["ECTS4"]
                ])
                moyUE2 = calculate_ects_weighted_average(ue2_notes, [
                    ects_data["ECTS5"], ects_data["ECTS6"], ects_data["ECTS7"]
                ])
                moyUE3 = calculate_ects_weighted_average(ue3_notes, [
                    ects_data["ECTS8"]
                ])
                # Fix: Only use ECTS9 through ECTS15 for BG-ALT-S2
                moyUE4 = calculate_ects_weighted_average(ue4_notes[:7], [  # Limit to 7 notes
                    ects_data["ECTS9"], ects_data["ECTS10"], ects_data["ECTS11"],
                    ects_data["ECTS12"], ects_data["ECTS13"], ects_data["ECTS14"],
                    ects_data["ECTS15"]
                ])
            else:  # modeleBGALT3.docx
                moyUE1 = calculate_ects_weighted_average(ue1_notes, [
                    ects_data["ECTS1"], ects_data["ECTS2"], ects_data["ECTS3"],
                    ects_data["ECTS4"], ects_data["ECTS5"]
                ])
                moyUE2 = calculate_ects_weighted_average(ue2_notes, [
                    ects_data["ECTS6"], ects_data["ECTS7"]
                ])
                moyUE3 = calculate_ects_weighted_average(ue3_notes, [
                    ects_data["ECTS8"]
                ])
                moyUE4 = calculate_ects_weighted_average(ue4_notes, [
                    ects_data["ECTS9"], ects_data["ECTS10"], ects_data["ECTS11"],
                    ects_data["ECTS12"], ects_data["ECTS13"]
                ])

            # Calculer les totaux d'ECTS pour chaque UE
            if template_name == "modeleBGALT2.docx":
                ects_ue1 = sum(int(ects_data[f"ECTS{i}"]) for i in range(1, 5))
                ects_ue2 = sum(int(ects_data[f"ECTS{i}"]) for i in range(5, 8))
                ects_ue3 = int(ects_data["ECTS8"])
                ects_ue4 = sum(int(ects_data[f"ECTS{i}"]) for i in range(9, 16))
            else:  # modeleBGALT3.docx
                ects_ue1 = sum(int(ects_data[f"ECTS{i}"]) for i in range(1, 6))
                ects_ue2 = sum(int(ects_data[f"ECTS{i}"]) for i in range(6, 8))
                ects_ue3 = int(ects_data["ECTS8"])
                ects_ue4 = sum(int(ects_data[f"ECTS{i}"]) for i in range(9, 14))

            moyenne_ects = ects_ue1 + ects_ue2 + ects_ue3 + ects_ue4

            try:
                if moyenne_ects > 0:
                    moyenne_ponderee = (
                        float(moyUE1 or 0) * ects_ue1 +
                        float(moyUE2 or 0) * ects_ue2 +
                        float(moyUE3 or 0) * ects_ue3 +
                        float(moyUE4 or 0) * ects_ue4
                    ) / moyenne_ects
                    moyenne_ponderee_str = f"{moyenne_ponderee:.2f}"
                else:
                    moyenne_ponderee_str = ""
            except (ValueError, TypeError, ZeroDivisionError):
                moyenne_ponderee_str = ""

            word_bytes = base64.b64decode(str(word_template.fileData))
            doc = Document(BytesIO(word_bytes))

            # Préparer les données de l'étudiant selon le template
            if template_name == "modeleBGALT2.docx":
                student_data = {
                    "CodeApprenant": str(updated_ws[f"A{row}"].value or ""),
                    "nomApprenant": str(updated_ws[f"B{row}"].value or ""),
                    "note1": calculate_single_note_average(updated_ws[f"D{row}"].value),
                    "note2": calculate_single_note_average(updated_ws[f"E{row}"].value),
                    "note3": calculate_single_note_average(updated_ws[f"F{row}"].value),
                    "note4": calculate_single_note_average(updated_ws[f"G{row}"]. value),
                    "note5": calculate_single_note_average(updated_ws[f"I{row}"].value),
                    "note6": calculate_single_note_average(updated_ws[f"J{row}"].value),
                    "note7": calculate_single_note_average(updated_ws[f"K{row}"].value),
                    "note8": calculate_single_note_average(updated_ws[f"M{row}"].value),
                    "note9": calculate_single_note_average(updated_ws[f"O{row}"].value),
                    "note10": calculate_single_note_average(updated_ws[f"P{row}"].value),
                    "note11": calculate_single_note_average(updated_ws[f"Q{row}"].value),
                    "note12": calculate_single_note_average(updated_ws[f"R{row}"].value),
                    "note13": calculate_single_note_average(updated_ws[f"S{row}"].value),
                    "note14": calculate_single_note_average(updated_ws[f"T{row}"].value),
                    "note15": calculate_single_note_average(updated_ws[f"U{row}"].value),
                    "moyUE1": moyUE1,
                    "moyUE2": moyUE2,
                    "moyUE3": moyUE3,
                    "moyUE4": moyUE4,
                    "moyenne": moyenne_ponderee_str,
                    "dateNaissance": str(updated_ws[f"V{row}"].value or ""),
                    "campus": str(updated_ws[f"W{row}"].value or ""),
                    "groupe": str(updated_ws[f"Y{row}"].value or ""),
                    "etendugroupe": str(updated_ws[f"Z{row}"].value or ""),
                    "justifiee": str(updated_ws[f"AA{row}"].value or ""),
                    "injustifiee": str(updated_ws[f"AB{row}"].value or ""),
                    "retard": str(updated_ws[f"AC{row}"].value or ""),
                    "APPRECIATIONS": str(updated_ws[f"AD{row}"].value or ""),
                    "datedujour": date_du_jour,
                    "ECTSUE1": str(ects_ue1),
                    "ECTSUE2": str(ects_ue2),
                    "ECTSUE3": str(ects_ue3),
                    "ECTSUE4": str(ects_ue4),
                    "moyenneECTS": str(moyenne_ects),
                    **ue_matieres,
                    **ects_data
                }
            else:  # modeleBGALT3.docx
                student_data = {
                    "CodeApprenant": str(updated_ws[f"A{row}"].value or ""),
                    "nomApprenant": str(updated_ws[f"B{row}"].value or ""),
                    "note1": calculate_single_note_average(updated_ws[f"D{row}"].value),
                    "note2": calculate_single_note_average(updated_ws[f"E{row}"].value),
                    "note3": calculate_single_note_average(updated_ws[f"F{row}"].value),
                    "note4": calculate_single_note_average(updated_ws[f"G{row}"].value),
                    "note5": calculate_single_note_average(updated_ws[f"H{row}"].value),
                    "note6": calculate_single_note_average(updated_ws[f"J{row}"].value),
                    "note7": calculate_single_note_average(updated_ws[f"K{row}"].value),
                    "note8": calculate_single_note_average(updated_ws[f"M{row}"].value),
                    "note9": calculate_single_note_average(updated_ws[f"O{row}"].value),
                    "note10": calculate_single_note_average(updated_ws[f"P{row}"].value),
                    "note11": calculate_single_note_average(updated_ws[f"Q{row}"].value),
                    "note12": calculate_single_note_average(updated_ws[f"R{row}"].value),
                    "note13": calculate_single_note_average(updated_ws[f"S{row}"].value),
                    "moyUE1": moyUE1,
                    "moyUE2": moyUE2,
                    "moyUE3": moyUE3,
                    "moyUE4": moyUE4,
                    "moyenne": moyenne_ponderee_str,
                    "dateNaissance": str(updated_ws[f"T{row}"].value or ""),
                    "campus": str(updated_ws[f"U{row}"].value or ""),
                    "groupe": str(updated_ws[f"W{row}"].value or ""),
                    "etendugroupe": str(updated_ws[f"X{row}"].value or ""),
                    "justifiee": str(updated_ws[f"Y{row}"].value or ""),
                    "injustifiee": str(updated_ws[f"Z{row}"].value or ""),
                    "retard": str(updated_ws[f"AA{row}"].value or ""),
                    "APPRECIATIONS": str(updated_ws[f"AB{row}"].value or ""),
                    "datedujour": date_du_jour,
                    "ECTSUE1": str(ects_ue1),
                    "ECTSUE2": str(ects_ue2),
                    "ECTSUE3": str(ects_ue3),
                    "ECTSUE4": str(ects_ue4),
                    "moyenneECTS": str(moyenne_ects),
                    **ue_matieres,
                    **ects_data
                }

            # Calculer les états pour chaque note
            etats = {}
            if template_name == "modeleBGALT2.docx":
                for i in range(1, 16):
                    etats[f"etat{i}"] = get_etat(student_data[f"note{i}"])
            else:
                for i in range(1, 14):
                    etats[f"etat{i}"] = get_etat(student_data[f"note{i}"])

            student_data.update(etats)

            # Calculer les états des UE
            if template_name == "modeleBGALT2.docx":
                etats_ue = {
                    "etatUE1": get_etat_ue([etats[f"etat{i}"] for i in range(1, 5)], student_data["moyUE1"]),
                    "etatUE2": get_etat_ue([etats[f"etat{i}"] for i in range(5, 8)], student_data["moyUE2"]),
                    "etatUE3": get_etat_ue([etats["etat8"]], student_data["moyUE3"]),
                    "etatUE4": get_etat_ue([etats[f"etat{i}"] for i in range(9, 16)], student_data["moyUE4"])
                }
            else:
                etats_ue = {
                    "etatUE1": get_etat_ue([etats[f"etat{i}"] for i in range(1, 6)], student_data["moyUE1"]),
                    "etatUE2": get_etat_ue([etats[f"etat{i}"] for i in range(6, 8)], student_data["moyUE2"]),
                    "etatUE3": get_etat_ue([etats["etat8"]], student_data["moyUE3"]),
                    "etatUE4": get_etat_ue([etats[f"etat{i}"] for i in range(9, 14)], student_data["moyUE4"])
                }

            student_data.update(etats_ue)

            # Calculer le total des états
            student_data["totaletat"] = get_total_etat(
                etats_ue["etatUE1"],
                etats_ue["etatUE2"],
                etats_ue["etatUE3"],
                etats_ue["etatUE4"]
            )

            # Ajuster les ECTS en fonction des notes
            def adjust_ects(note_str, original_ects):
                try:
                    note = float(note_str) if note_str else 0
                    return "0" if note < 8 else str(original_ects)
                except (ValueError, TypeError):
                    return str(original_ects)

            # Ajuster les ECTS pour chaque matière
            if template_name == "modeleBGALT2.docx":
                for i in range(1, 16):
                    student_data[f"ECTS{i}"] = adjust_ects(student_data[f"note{i}"], ects_data[f"ECTS{i}"])
            else:
                for i in range(1, 14):
                    student_data[f"ECTS{i}"] = adjust_ects(student_data[f"note{i}"], ects_data[f"ECTS{i}"])

            # Recalculer les totaux d'ECTS avec les ECTS ajustés
            if template_name == "modeleBGALT2.docx":
                student_data["ECTSUE1"] = str(sum(int(student_data[f"ECTS{i}"]) for i in range(1, 5)))
                student_data["ECTSUE2"] = str(sum(int(student_data[f"ECTS{i}"]) for i in range(5, 8)))
                student_data["ECTSUE3"] = student_data["ECTS8"]
                student_data["ECTSUE4"] = str(sum(int(student_data[f"ECTS{i}"]) for i in range(9, 16)))
            else:
                student_data["ECTSUE1"] = str(sum(int(student_data[f"ECTS{i}"]) for i in range(1, 6)))
                student_data["ECTSUE2"] = str(sum(int(student_data[f"ECTS{i}"]) for i in range(6, 8)))
                student_data["ECTSUE3"] = student_data["ECTS8"]
                student_data["ECTSUE4"] = str(sum(int(student_data[f"ECTS{i}"]) for i in range(9, 14)))

            student_data["moyenneECTS"] = str(sum(int(student_data[f"ECTSUE{i}"]) for i in range(1, 5)))

            # Remplacer les variables dans le document
            for paragraph in doc.paragraphs:
                for key, value in student_data.items():
                    placeholder = f"{{{{{key}}}}}"
                    if placeholder in paragraph.text:
                        paragraph.text = paragraph.text.replace(placeholder, str(value))

            # Remplacer dans les tableaux
            for table in doc.tables:
                for table_row in table.rows:
                    for cell in table_row.cells:
                        for key, value in student_data.items():
                            placeholder = f"{{{{{key}}}}}"
                            if placeholder in cell.text:
                                cell.text = cell.text.replace(placeholder, str(value))

            # Sauvegarder le bulletin
            safe_nom = "".join(c for c in student_data["nomApprenant"] if c.isalnum() or c in (' ', '-', '_')).strip()
            bulletin_path = os.path.join(bulletins_dir, f"bulletin_{safe_nom}.docx")
            doc.save(bulletin_path)
            logging.info(f"Bulletin créé pour {student_data['nomApprenant']}")

        await db.disconnect()
        return {
            "message": "Bulletins générés avec succès",
            "bulletins_directory": bulletins_dir
        }

    except Exception as e:
        logging.error(f"Erreur lors de la génération des bulletins : {str(e)}")
        raise HTTPException(status_code=400, detail=str(e))
