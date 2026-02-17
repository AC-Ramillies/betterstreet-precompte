# BetterStreet – Extraction planning précompte

Outil interne développé par le Service Gouvernance Numerique de la Commune de Ramillies.

Ce script Python permet de transformer un export CSV issu de BetterStreet (Inforius) en un fichier Excel structuré et exploitable pour le traitement du précompte et l’analyse du planning du Service Travaux.

## Contexte

L’export CSV standard de BetterStreet présente certaines limitations :

- Lignes parfois mal structurées (retours à la ligne dans les cellules)
- Colonnes décalées
- Ambiguïtés horaires (format 12h / 24h)
- Interventions sans planification
- Absence éventuelle d’agents affectés

Ce script a été développé afin de :

- Fiabiliser les données extraites
- Identifier les incohérences
- Produire un fichier conforme au modèle interne du Service Travaux
- Assurer la reproductibilité annuelle du traitement

## Fonctionnalités principales

### Reconstruction des enregistrements
Reconstitution correcte des interventions même lorsque l’export CSV est imparfait.

### Filtrage par année
Filtrage automatique des interventions selon l’année cible (ex. 2025).

### Correction horaire encadrée
Détection des cas où l’heure de fin est inférieure à l’heure de début et application d’une correction +12h avec traçabilité.

### Détection d’anomalies
Création d’un onglet "Anomalies" reprenant notamment :

- Début manquant
- Fin manquante
- Agents/Équipes manquants
- Corrections horaires appliquées
- Erreurs de parsing

### Mise en forme Excel
- Tri chronologique
- Surlignage des cas à vérifier
- Format compatible avec le modèle précompte communal

## Utilisation

### Prérequis

- Python 3.10 ou supérieur
- Modules :
  - pandas
  - openpyxl

Installation des dépendances :

```bash
pip install pandas openpyxl
```

### Exécution

Dans le dossier dans lequel se trouve les fichiers **betterstreet_to_precompte_v5.py** et **run_precompte.ps1**, ajoutez votre export CSV issu de BetterStreet. **Attention**, votre export doit porter le nom **"Export_Planning_BetterStreet"**.

Dans PowerShell, lancez ensuite ces commandes :

```bash
.\run_precompte.ps1
```

OU

```bash
python betterstreet_to_precompte_v5.py Export_Planning_BetterStreet.csv 2025
```
OU avec nom de fichier de sortie :

```bash
python betterstreet_to_precompte_v5.py Export_Planning_BetterStreet.csv 2025 Planning_Ouvriers_2025.xlsx

```

