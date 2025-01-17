import logging
import json

# Définition des ECTS en tant que constante
ECTS_DATA = {
    "M1-S1": [{"ECTS1": 2, "ECTS2": 2, "ECTS3": 2, "ECTS4": 3, "ECTS5": 2, "ECTS6": 2, "ECTS7": 2, "ECTS8": 0, "ECTS9": 0, "ECTS10": 0, "ECTS11": 9, "ECTS12": 0, "ECTS13": 2, "ECTS14": 2, "ECTS15": 2}],
    "M1-S2": [{"ECTS1": 2, "ECTS2": 2, "ECTS3": 2, "ECTS4": 2, "ECTS5": 1, "ECTS6": 9, "ECTS7": 2, "ECTS8": 2, "ECTS9": 0, "ECTS10": 0, "ECTS11": 0, "ECTS12": 0, "ECTS13": 2, "ECTS14": 2, "ECTS15": 2, "ECTS16": 2}],
    "M2-S3-MAGI": [{"ECTS1": 2, "ECTS2": 2, "ECTS3": 2, "ECTS4": 0, "ECTS5": 9, "ECTS6": 2, "ECTS7": 2, "ECTS8": 0, "ECTS9": 0, "ECTS10": 2, "ECTS11": 3, "ECTS12": 3, "ECTS13": 3}],
    "M2-S3-MEFIM": [{"ECTS1": 2, "ECTS2": 2, "ECTS3": 2, "ECTS4": 0, "ECTS5": 9, "ECTS6": 2, "ECTS7": 2, "ECTS8": 0, "ECTS9": 0, "ECTS10": 3, "ECTS11": 2, "ECTS12": 3, "ECTS13": 3}],
    "M2-S3-MAPI": [{"ECTS1": 2, "ECTS2": 2, "ECTS3": 2, "ECTS4": 0, "ECTS5": 9, "ECTS6": 2, "ECTS7": 2, "ECTS8": 0, "ECTS9": 0, "ECTS10": 2, "ECTS11": 3, "ECTS12": 2, "ECTS13": 2, "ECTS14": 2}],
    "M2-S4": [{"ECTS1": 2, "ECTS2": 2, "ECTS3": 2, "ECTS4": 2, "ECTS5": 5, "ECTS6": 0, "ECTS7": 0, "ECTS8": 9, "ECTS9": 4, "ECTS10": 2, "ECTS11": 2}],
    "BG_ALT_1": [{"ECTS1": 3, "ECTS2": 3, "ECTS3": 3, "ECTS4": 3, "ECTS5": 3, "ECTS6": 2, "ECTS7": 9, "ECTS8": 2, "ECTS9": 0, "ECTS10": 0, "ECTS11": 0, "ECTS12": 2, "ECTS13": 0, "ECTS14": 0}],
    "BG_ALT_2": [{"ECTS1": 3, "ECTS2": 2, "ECTS3": 3, "ECTS4": 2, "ECTS5": 2, "ECTS6": 2, "ECTS7": 2, "ECTS8": 3, "ECTS9": 9, "ECTS10": 2, "ECTS11": 0, "ECTS12": 0, "ECTS13": 0, "ECTS14": 0, "ECTS15": 0}],
    "BG_ALT_3": [{"ECTS1": 3, "ECTS2": 3, "ECTS3": 2, "ECTS4": 0, "ECTS5": 3, "ECTS6": 3, "ECTS7": 3, "ECTS8": 2, "ECTS9": 9, "ECTS10": 2, "ECTS11": 0, "ECTS12": 0, "ECTS13": 0}],
    "BG_ALT_4": [{"ECTS1": 2, "ECTS2": 4, "ECTS3": 3, "ECTS4": 2, "ECTS5": 2, "ECTS6": 2, "ECTS7": 2, "ECTS8": 2, "ECTS9": 9, "ECTS10": 2, "ECTS11": 0, "ECTS12": 0, "ECTS13": 0}],
    "BG_ALT_5": [{"ECTS1": 3, "ECTS2": 3, "ECTS3": 3, "ECTS4": 2, "ECTS5": 2, "ECTS6": 3, "ECTS7": 3, "ECTS8": 7, "ECTS9": 2, "ECTS10": 0, "ECTS11": 0, "ECTS12": 0, "ECTS13": 2, "ECTS14": 0}],
    "BG_ALT_6": [{"ECTS1": 2, "ECTS2": 2, "ECTS3": 3, "ECTS4": 2, "ECTS5": 2, "ECTS6": 2, "ECTS7": 7, "ECTS8": 4, "ECTS9": 0, "ECTS10": 4, "ECTS11": 2, "ECTS12": 0}],
    "BG_TP_1": [{"ECTS1": 3, "ECTS2": 3, "ECTS3": 3, "ECTS4": 3, "ECTS5": 2, "ECTS6": 3, "ECTS7": 2, "ECTS8": 3, "ECTS9": 3, "ECTS10": 2, "ECTS11": 2, "ECTS12": 2, "ECTS13": 2, "ECTS14": 3, "ECTS15": 2, "ECTS16": 0, "ECTS17": 0, "ECTS18": 0, "ECTS19": 2, "ECTS20": 0, "ECTS21": 0, "ECTS22": 0}],
    "BG_TP_2": [{"ECTS1": 18, "ECTS2": 2}],
    "BG_TP_3": [{"ECTS1": 3, "ECTS2": 3, "ECTS3": 2, "ECTS4": 3, "ECTS5": 3, "ECTS6": 3, "ECTS7": 2, "ECTS8": 2, "ECTS9": 2, "ECTS10": 2, "ECTS11": 3, "ECTS12": 2, "ECTS13": 0, "ECTS14": 0, "ECTS15": 0}],
    "BG_TP_4": [{"ECTS1": 30}],
    "BG_TP_5": [{"ECTS1": 3, "ECTS2": 3, "ECTS3": 3, "ECTS4": 2, "ECTS5": 2, "ECTS6": 2, "ECTS7": 2, "ECTS8": 3, "ECTS9": 3, "ECTS10": 2, "ECTS11": 3, "ECTS12": 2, "ECTS13": 2, "ECTS14": 2, "ECTS15": 0, "ECTS16": 0, "ECTS17": 2, "ECTS18": 0, "ECTS19": 4}],
    "BG_TP_6": [{"ECTS1": 14, "ECTS2": 4, "ECTS3": 2}]
}

async def get_ects_for_template(template_name: str):
    """
    Récupère les ECTS pour un template donné depuis les données statiques
    """
    try:
        # Normaliser le nom du template (enlever le _S s'il existe)
        normalized_name = template_name.replace("_S", "")
        
        if normalized_name not in ECTS_DATA:
            logging.error(f"Template ECTS non trouvé : {normalized_name}")
            raise ValueError(f"Template ECTS non trouvé : {normalized_name}")
            
        # Récupérer les données ECTS
        ects_data = ECTS_DATA[normalized_name][0]
        
        # Convertir toutes les valeurs en chaînes
        result = {k: str(v) for k, v in ects_data.items()}
        logging.info(f"ECTS trouvés pour {normalized_name}: {result}")
        
        return result
        
    except Exception as e:
        logging.error(f"Erreur lors de la récupération des ECTS : {str(e)}")
        raise