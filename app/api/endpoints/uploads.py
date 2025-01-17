import base64
import datetime
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
from app.services.excel_service import process_excel_with_template
from app.services.word_service import generate_bulletins_from_excel
from prisma import Prisma

router = APIRouter()

def calculate_weighted_average(notes_list):
    """
    Calcule la moyenne pondérée des notes
    notes_list: liste des notes de l'UE
    """
    if not notes_list:
        return ""
        
    try:
        total_weighted_sum = 0
        total_coefficients = 0
        
        for note in notes_list:
            if note is None or note == "":
                continue
                
            # Convertir en string et nettoyer
            note_str = str(note).strip()
            
            try:
                # Cas des notes multiples séparées par des tirets
                if " - " in note_str:
                    sub_notes = note_str.split(" - ")
                    for sub_note in sub_notes:
                        sub_note = sub_note.strip()
                        # Ignorer les mentions "Absent au devoir"
                        if "Absent" in sub_note:
                            continue
                            
                        if "(" in sub_note and ")" in sub_note:
                            # Format: "note (coeff)"
                            note_parts = sub_note.split("(")
                            note_value = float(note_parts[0].strip().replace(",", "."))
                            coeff = float(note_parts[1].split(")")[0].strip().replace(",", "."))
                            total_weighted_sum += note_value * coeff
                            total_coefficients += coeff
                        else:
                            # Format: simple note (coefficient 1)
                            note_value = float(sub_note.replace(",", "."))
                            total_weighted_sum += note_value
                            total_coefficients += 1
                # Cas d'une seule note avec coefficient
                elif "(" in note_str and ")" in note_str:
                    # Ignorer les mentions "Absent au devoir"
                    if "Absent" not in note_str:
                        note_parts = note_str.split("(")
                        note_value = float(note_parts[0].strip().replace(",", "."))
                        coeff = float(note_parts[1].split(")")[0].strip().replace(",", "."))
                        total_weighted_sum += note_value * coeff
                        total_coefficients += coeff
                else:
                    # Ignorer les mentions "Absent au devoir"
                    if "Absent" not in note_str:
                        note_value = float(note_str.replace(",", "."))
                        total_weighted_sum += note_value
                        total_coefficients += 1
            except (ValueError, IndexError) as e:
                logging.error(f"Erreur lors du traitement de la note {note_str}: {str(e)}")
                continue
        
        # Calculer la moyenne pondérée
        if total_coefficients > 0:
            average = total_weighted_sum / total_coefficients
            # Arrondir au centième
            return f"{average:.2f}".replace(".", ",")
        return ""
        
    except Exception as e:
        logging.error(f"Erreur dans le calcul de la moyenne: {str(e)}")
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
        # Connexion à Prisma
        db = Prisma()
        await db.connect()

        # Charger le fichier Excel mis à jour
        excel_path = os.path.join("./temp", "updated_excel.xlsx")
        if not os.path.exists(excel_path):
            raise ValueError("Fichier Excel mis à jour non trouvé dans ./temp")

        updated_wb = openpyxl.load_workbook(excel_path)
        updated_ws = updated_wb.active

        # Récupérer les titres des UE et matières depuis la première ligne
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
                "filename": "modeleBGALT3.docx",
                "isTemplate": True
            }
        )
        
        if not word_template:
            raise ValueError("Template Word modeleBGALT3.docx non trouvé dans Prisma")

        # Créer le dossier bulletins s'il n'existe pas
        bulletins_dir = os.path.join("./temp", "bulletins")
        if not os.path.exists(bulletins_dir):
            os.makedirs(bulletins_dir)

        # Pour chaque étudiant, créer un bulletin personnalisé
        for row in range(3, updated_ws.max_row + 1):
            if not updated_ws[f"B{row}"].value:
                continue

            # Récupérer les notes pour chaque UE
            ue1_notes = [
                updated_ws[f"D{row}"].value,
                updated_ws[f"E{row}"].value,
                updated_ws[f"F{row}"].value,
                updated_ws[f"G{row}"].value,
                updated_ws[f"H{row}"].value
            ]

            ue2_notes = [
                updated_ws[f"J{row}"].value,
                updated_ws[f"K{row}"].value
            ]

            ue3_notes = [
                updated_ws[f"M{row}"].value
            ]

            ue4_notes = [
                updated_ws[f"O{row}"].value,
                updated_ws[f"P{row}"].value,
                updated_ws[f"Q{row}"].value,
                updated_ws[f"R{row}"].value,
                updated_ws[f"S{row}"].value
            ]

            # Calculer les moyennes
            moyenne_ue1 = calculate_weighted_average(ue1_notes)
            moyenne_ue2 = calculate_weighted_average(ue2_notes)
            moyenne_ue3 = calculate_weighted_average(ue3_notes)
            moyenne_ue4 = calculate_weighted_average(ue4_notes)
            
            # Calculer la moyenne générale
            all_notes = ue1_notes + ue2_notes + ue3_notes + ue4_notes
            moyenne_generale = calculate_weighted_average(all_notes)

            # Charger le template Word pour chaque étudiant
            word_bytes = base64.b64decode(str(word_template.fileData))
            doc = Document(BytesIO(word_bytes))

            # Préparer les données de l'étudiant
            student_data = {
                "CodeApprenant": str(updated_ws[f"A{row}"].value or ""),
                "nomApprenant": str(updated_ws[f"B{row}"].value or ""),
                "note1": str(updated_ws[f"D{row}"].value or ""),
                "note2": str(updated_ws[f"E{row}"].value or ""),
                "note3": str(updated_ws[f"F{row}"].value or ""),
                "note4": str(updated_ws[f"G{row}"].value or ""),
                "note5": str(updated_ws[f"H{row}"].value or ""),
                "note6": str(updated_ws[f"J{row}"].value or ""),
                "note7": str(updated_ws[f"K{row}"].value or ""),
                "note8": str(updated_ws[f"M{row}"].value or ""),
                "note9": str(updated_ws[f"O{row}"].value or ""),
                "note10": str(updated_ws[f"P{row}"].value or ""),
                "note11": str(updated_ws[f"Q{row}"].value or ""),
                "note12": str(updated_ws[f"R{row}"].value or ""),
                "note13": str(updated_ws[f"S{row}"].value or ""),
                "moyenne_ue1": str(moyenne_ue1),
                "moyenne_ue2": str(moyenne_ue2),
                "moyenne_ue3": str(moyenne_ue3),
                "moyenne_ue4": str(moyenne_ue4),
                "moyenne_generale": str(moyenne_generale),
                "dateNaissance": str(updated_ws[f"T{row}"].value or ""),
                "campus": str(updated_ws[f"U{row}"].value or ""),
                "groupe": str(updated_ws[f"W{row}"].value or ""),
                "etendugroupe": str(updated_ws[f"X{row}"].value or ""),
                "justifiee": str(updated_ws[f"Y{row}"].value or ""),
                "injustifiee": str(updated_ws[f"Z{row}"].value or ""),
                "retard": str(updated_ws[f"AA{row}"].value or ""),
                "APPRECIATIONS": str(updated_ws[f"AB{row}"].value or ""),
                **ue_matieres
            }

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