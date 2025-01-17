from typing import List, Optional

class Site:
    def __init__(self, codeSite: int, nomSite: str, etenduSite: str):
        self.codeSite = codeSite
        self.nomSite = nomSite
        self.etenduSite = etenduSite
        
class Absence:
    def __init__(self, code_apprenant=None, duree=None, is_justifie=None, is_retard=None):
        self.code_apprenant = code_apprenant
        self.duree = duree
        self.is_justifie = is_justifie
        self.is_retard = is_retard

class Inscription:
    def __init__(self, codeSite: int, site: Site):
        self.codeSite = codeSite
        self.site = site

class Apprenant:
    def __init__(self, code_apprenant, nom_apprenant, prenom_apprenant, date_naissance, inscriptions, code_groupe=None):
        self.code_apprenant = code_apprenant
        self.nom_apprenant = nom_apprenant
        self.prenom_apprenant = prenom_apprenant
        self.date_naissance = date_naissance
        self.inscriptions = inscriptions
        self.code_groupe = code_groupe

class Groupe:
    def __init__(self, code_groupe, nom_groupe, etendu_groupe):
        self.code_groupe = code_groupe
        self.nom_groupe = nom_groupe
        self.etendu_groupe = etendu_groupe

class Periode:
    def __init__(self, codePeriode: int, nomPeriode: str, dateDeb: str, dateFin: str):
        self.codePeriode = codePeriode
        self.nomPeriode = nomPeriode
        self.dateDeb = dateDeb
        self.dateFin = dateFin

class Frequente:
    def __init__(self, codeFrequente: int, codeGroupe: int, codeApprenant: int):
        self.codeFrequente = codeFrequente
        self.codeGroupe = codeGroupe
        self.codeApprenant = codeApprenant
