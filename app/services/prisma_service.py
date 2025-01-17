import base64
import logging
import os
from prisma import Prisma

async def fetch_template_from_prisma(template_name: str) -> bytes:
    """
    Récupère le fichier Excel depuis Prisma et retourne directement les données en Bytes.
    """
    try:
        logging.info(f"Connexion à Prisma pour récupérer le fichier {template_name}...")
        db = Prisma()
        await db.connect()

        file_record = await db.generatedfile.find_first(where={"filename": template_name, "isTemplate": True})
        
        if file_record and file_record.fileData:
            logging.info(f"Type de fileData récupéré: {type(file_record.fileData)}")
            
            try:
                # Convertir directement l'objet Base64 en string
                decoded_data = base64.b64decode(str(file_record.fileData))
                logging.info(f"Décodage base64 réussi, taille: {len(decoded_data)} bytes")
                return decoded_data
            except Exception as decode_error:
                logging.error(f"Erreur lors du décodage base64: {str(decode_error)}")
                raise ValueError(f"Erreur lors du décodage base64: {str(decode_error)}")
        else:
            raise ValueError(f"Template {template_name} introuvable dans Prisma.")
    except Exception as e:
        logging.error(f"Erreur lors de la récupération du template: {str(e)}")
        raise
    finally:
        await db.disconnect()

async def save_file_to_prisma(filename: str, file_data: bytes, is_template: bool = False) -> None:
    """
    Sauvegarde un fichier dans Prisma en le convertissant en base64.
    """
    try:
        db = Prisma()
        await db.connect()

        # Convertir les bytes en base64
        base64_data = base64.b64encode(file_data).decode('utf-8')
        logging.info(f"Conversion en base64 réussie, taille: {len(base64_data)}")

        await db.generatedfile.create({
            'filename': filename,
            'fileType': filename.split('.')[-1],
            'fileData': base64_data,
            'isTemplate': is_template
        })

        logging.info(f"Fichier {filename} sauvegardé avec succès dans Prisma")
    except Exception as e:
        logging.error(f"Erreur lors de la sauvegarde du fichier dans Prisma: {str(e)}")
        raise
    finally:
        await db.disconnect()

async def get_word_template(template_name: str = "modeleBGALT3.docx") -> str:
    """
    Récupère le template Word depuis Prisma et le sauvegarde dans ./temp
    Retourne le chemin du fichier sauvegardé.
    """
    try:
        # Créer le dossier temp s'il n'existe pas
        temp_dir = "./temp"
        if not os.path.exists(temp_dir):
            os.makedirs(temp_dir)
            logging.info(f"Dossier temp créé : {temp_dir}")

        # Récupérer le template
        template_data = await fetch_template_from_prisma(template_name)
        
        # Sauvegarder le fichier
        template_path = os.path.join(temp_dir, template_name)
        with open(template_path, 'wb') as f:
            f.write(template_data)
        
        logging.info(f"Template Word sauvegardé avec succès : {template_path}")
        return template_path
        
    except Exception as e:
        logging.error(f"Erreur lors de la récupération du template Word : {str(e)}")
        raise
        
async def get_template_from_prisma(template_name: str, output_dir: str) -> str:
    """
    Récupère un fichier Excel depuis Prisma, le sauvegarde localement, et retourne le chemin.
    """
    file_bytes = await fetch_template_from_prisma(template_name)
    template_path = os.path.join(output_dir, template_name)
    with open(template_path, "wb") as f:
        f.write(file_bytes)
    logging.info(f"Template sauvegardé temporairement à : {template_path}")
    return template_path

async def get_template_from_prisma(template_name: str, output_dir: str) -> str:
    """
    Récupère un fichier Excel depuis Prisma, le sauvegarde localement, et retourne le chemin.
    """
    try:
        file_bytes = await fetch_template_from_prisma(template_name)

        # Assurez-vous que le répertoire temporaire existe
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
            logging.info(f"Répertoire temporaire créé : {output_dir}")

        # Sauvegarde du fichier
        template_path = os.path.join(output_dir, template_name)
        with open(template_path, "wb") as f:
            f.write(file_bytes)

        logging.info(f"Template sauvegardé temporairement à : {template_path}")
        return template_path
    except Exception as e:
        logging.error(f"Erreur lors de la sauvegarde du fichier depuis Prisma : {str(e)}")
        raise ValueError(f"Erreur lors de la récupération du fichier depuis Prisma : {e}")
    
async def save_excel_to_prisma(file_path: str, user_id: str) -> int:
    """
    Sauvegarde un fichier Excel dans Prisma en l'encodant en base64.
    """
    try:
        # Lire le fichier en binaire et encoder en base64
        with open(file_path, 'rb') as file:
            file_content = file.read()
            encoded_content = base64.b64encode(file_content)
            
        # Connexion à Prisma et sauvegarde
        db = Prisma()
        await db.connect()
        
        # Créer l'enregistrement GeneratedExcel
        generated_excel = await db.generatedexcel.create({
            'userId': user_id,
            'templateId': 2,  # ID du template Excel dans GeneratedFile
            'data': encoded_content,
        })
        
        await db.disconnect()
        return generated_excel.id

    except Exception as e:
        logging.error(f"Erreur lors de la sauvegarde de l'Excel dans Prisma : {str(e)}")
        raise ValueError(f"Erreur lors de la sauvegarde de l'Excel dans Prisma : {str(e)}")

async def get_excel_from_prisma(excel_id: int) -> bytes:
    """
    Récupère un fichier Excel depuis Prisma par son ID.
    """
    try:
        db = Prisma()
        await db.connect()
        
        excel_record = await db.generatedexcel.find_unique(
            where={'id': excel_id}
        )
        
        await db.disconnect()
        
        if excel_record and excel_record.data:
            return excel_record.data
        else:
            raise ValueError(f"Excel {excel_id} non trouvé dans Prisma")
            
    except Exception as e:
        logging.error(f"Erreur lors de la récupération de l'Excel depuis Prisma : {str(e)}")
        raise

