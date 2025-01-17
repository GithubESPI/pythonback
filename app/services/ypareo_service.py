import logging
import os
import requests
from dotenv import load_dotenv
from datetime import datetime

# Charger les variables d'environnement
load_dotenv()

class YpareoService:
    BASE_URL = os.getenv("YPAERO_BASE_URL")
    API_TOKEN = os.getenv("YPAERO_API_TOKEN")

    @staticmethod
    def fetch_json(endpoint: str):
        if not YpareoService.BASE_URL or not YpareoService.API_TOKEN:
            raise ValueError("Environment variables for Ypareo are not set.")

        url = f"{YpareoService.BASE_URL}{endpoint}"
        headers = {"X-Auth-Token": YpareoService.API_TOKEN}
        response = requests.get(url, headers=headers)
        if response.status_code != 200:
            raise Exception(f"Erreur API Yparéo: {response.status_code} - {response.text}")
        return response.json()

    @staticmethod
    def get_periode_2023_2024():
        return next((p for p in YpareoService.fetch_json("/r/v1/periodes").values() if p["codePeriode"] == 2), None)

    @staticmethod
    def get_frequentes():
        return list(YpareoService.fetch_json("/r/v1/apprenants/frequentes?codesPeriode=2").values())

    @staticmethod
    def get_apprenants():
        return list(YpareoService.fetch_json("/r/v1/formation-longue/apprenants?codesPeriode=2").values())

    @staticmethod
    def get_groupes():
        return list(YpareoService.fetch_json("/r/v1/formation-longue/groupes?codesPeriode=2").values())
        
    @staticmethod
    def get_absences():
        absences_data = YpareoService.fetch_json("/r/v1/absences/01-09-2023/15-09-2024")
        # Organiser les absences par code apprenant
        absences_by_apprenant = {}
        for absence in absences_data.values():
            code_apprenant = str(absence.get("codeApprenant"))
            if code_apprenant:
                if code_apprenant not in absences_by_apprenant:
                    absences_by_apprenant[code_apprenant] = []
                absences_by_apprenant[code_apprenant].append(absence)
        return absences_by_apprenant # Convertir en liste comme les autres méthodes