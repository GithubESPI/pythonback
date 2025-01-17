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
from app.services.excel_service import process_excel_with_template
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
            return "0.00"

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
                return "0.00"
                    
            except (ValueError, IndexError) as e:
                logging.error(f"Erreur lors du calcul de la moyenne pondérée: {str(e)}")
                return "0.00"
        
        # Cas d'une note simple ou multiple sans coefficient (ex: "17 - 16 - 17")
        else:
            try:
                # Séparer les notes s'il y en a plusieurs
                notes = [float(n.strip().replace(",", ".")) for n in note_str.split("-") if n.strip() and "Absent au devoir" not in n]
                if notes:
                    # Calculer la moyenne simple (coefficient 1 pour chaque note)
                    average = sum(notes) / len(notes)
                    return f"{average:.2f}"
                return "0.00"
            except ValueError as e:
                logging.error(f"Erreur lors du calcul de la note simple: {str(e)}")
                return "0.00"

    except Exception as e:
        logging.error(f"Erreur lors du calcul de la note: {str(e)}")
        return "0.00"


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

        # Obtenir la date du jour au format français
        date_du_jour = datetime.utcnow().strftime("%d/%m/%Y")

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
                str(updated_ws[f"D{row}"].value or ""),
                str(updated_ws[f"E{row}"].value or ""),
                str(updated_ws[f"F{row}"].value or ""),
                str(updated_ws[f"G{row}"].value or ""),
                str(updated_ws[f"H{row}"].value or "")
            ]

            ue2_notes = [
                str(updated_ws[f"J{row}"].value or ""),
                str(updated_ws[f"K{row}"].value or "")
            ]

            ue3_notes = [
                str(updated_ws[f"M{row}"].value or "")
            ]

            ue4_notes = [
                str(updated_ws[f"O{row}"].value or ""),
                str(updated_ws[f"P{row}"].value or ""),
                str(updated_ws[f"Q{row}"].value or ""),
                str(updated_ws[f"R{row}"].value or ""),
                str(updated_ws[f"S{row}"].value or "")
            ]

            # Calculer les moyennes avec la nouvelle fonction
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

            # Préparer les données de l'étudiant avec les notes originales
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
                "moyenne_ue1": f"{moyenne_ue1:.2f}",
                "moyenne_ue2": f"{moyenne_ue2:.2f}",
                "moyenne_ue3": f"{moyenne_ue3:.2f}",
                "moyenne_ue4": f"{moyenne_ue4:.2f}",
                "moyenne_generale": f"{moyenne_generale:.2f}",
                "dateNaissance": str(updated_ws[f"T{row}"].value or ""),
                "campus": str(updated_ws[f"U{row}"].value or ""),
                "groupe": str(updated_ws[f"W{row}"].value or ""),
                "etendugroupe": str(updated_ws[f"X{row}"].value or ""),
                "justifiee": str(updated_ws[f"Y{row}"].value or ""),
                "injustifiee": str(updated_ws[f"Z{row}"].value or ""),
                "retard": str(updated_ws[f"AA{row}"].value or ""),
                "APPRECIATIONS": str(updated_ws[f"AB{row}"].value or ""),
                "datedujour": date_du_jour,
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