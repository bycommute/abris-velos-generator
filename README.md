# Calculateur Prix Camflex - Documentation Complète

## 🎯 Objectif Global du Projet

Ce calculateur de prix est un système automatisé complet qui permet de :

1. **Générer tous les fichiers Excel** pour chaque variant d'abrivélo existant chez ByCommute
2. **Calculer automatiquement les prix** de chaque variant via les formules Excel
3. **Extraire les prix et les listes de composants** de chaque fichier Excel
4. **Générer les URLs SharePoint Drive** pour héberger les fichiers
5. **Préparer les données pour Odoo** : prix unitaire après réduction, URLs des fichiers, et listes de composants

**Le processus complet** : Fichier de base → Génération Excel → Calcul prix → Extraction → URLs → Données pour Odoo

---

## 📋 Vue d'Ensemble du Processus

```
┌─────────────────────────────────────────────────────────────────┐
│                    PROCESSUS COMPLET                             │
└─────────────────────────────────────────────────────────────────┘

1. FICHIER DE BASE (nepastoucher.xlsx)
   ↓
   [Contient toutes les hypothèses et formules de calcul Camflex]
   ↓

2. GÉNÉRATION DES EXCEL (Scripts Python par type d'abri)
   ↓
   [Un script Python = Un type d'abri vélo]
   [Chaque script génère tous les variants de ce type]
   ↓

3. DOSSIER RÉSULTATS (résultats/)
   ↓
   [Tous les Excel générés, organisés par type d'abri]
   ↓

4. CALCUL DES FORMULES (extract_prices_and_components.py)
   ↓
   [Ouvre chaque Excel dans Microsoft Excel pour calculer les formules]
   ↓

5. EXTRACTION DES DONNÉES
   ↓
   [Prix avant/après réduction + Liste des composants]
   ↓

6. GÉNÉRATION DES URLs (generate_drive_urls.py)
   ↓
   [URLs SharePoint Drive pour chaque fichier]
   ↓

7. DONNÉES FINALES POUR ODOO
   ↓
   [resultats_tous.json + urls_drive.csv/xlsx]
   ↓
   [Upload dans Odoo : Prix + URLs + Composants]
```

---

## 📁 Structure Complète du Projet

```
.
├── fichier de base/
│   └── nepastoucher.xlsx          # ⭐ FICHIER SOURCE (voir section dédiée)
│
├── résultats/                     # ⭐ TOUS LES EXCEL GÉNÉRÉS (voir section dédiée)
│   ├── carport/
│   │   ├── CAR-2.5M-N-200-G.xlsx
│   │   ├── CAR-6M-P-250-PT.xlsx
│   │   └── ... (80 fichiers Excel)
│   ├── bosquet_ferme/
│   │   └── ... (200 fichiers Excel)
│   ├── bosquet_ferme_compact/
│   ├── bosquet_ouvert/
│   ├── domino_ferme/
│   ├── domino_ferme_compact/
│   ├── domino_ouvert/
│   ├── metallique_ferme/
│   ├── metallique_ferme_compact/
│   ├── metallique_ouvert/
│   └── neve_ouvert/
│
├── composant/                     # Composants détaillés extraits (JSON)
│   ├── carport/
│   ├── bosquet_ferme/
│   └── ...
│
├── calculateur_prix_camflex.py     # ⭐ SCRIPT PRINCIPAL (guide interactif)
├── extract_prices_and_components.py # Extraction des prix et composants
├── generate_drive_urls.py        # ⭐ GÉNÉRATEUR D'URLs SharePoint
│
├── generate_*.py                  # ⭐ SCRIPTS DE GÉNÉRATION (voir section dédiée)
│   ├── generate_carport.py
│   ├── generate_bosquet_ferme.py
│   ├── generate_bosquet_ferme_compact.py
│   ├── generate_bosquet_ouvert.py
│   ├── generate_domino_ferme.py
│   ├── generate_domino_ferme_compact.py
│   ├── generate_domino_ouvert.py
│   ├── generate_metallique_ferme.py
│   ├── generate_metallique_ferme_compact.py
│   ├── generate_metallique_ouvert.py
│   └── generate_neve_ouvert.py
│
├── resultats_tous.json            # ⭐ FICHIER FINAL (tous les prix)
├── urls_drive.csv                 # ⭐ URLs SharePoint (CSV)
├── urls_drive.xlsx                # ⭐ URLs SharePoint (Excel)
└── README.md                      # Ce fichier
```

---

## 🔑 Composants Clés du Système

### 1. Le Fichier de Base (`fichier de base/nepastoucher.xlsx`)

**Rôle :** C'est le fichier Excel source fourni par Camflex qui contient :
- Toutes les **hypothèses de calcul** (coûts des matériaux, main d'œuvre, etc.)
- Toutes les **formules Excel** qui calculent les prix en fonction des paramètres
- La structure de base qui sera copiée pour chaque variant

**⚠️ IMPORTANT :**
- **NE JAMAIS MODIFIER DIRECTEMENT** ce fichier
- C'est le fichier source de référence fourni par Camflex
- Tous les fichiers Excel générés sont des **copies** de ce fichier avec des paramètres différents

**Quand mettre à jour le fichier de base :**
- Quand Camflex fournit un nouveau fichier avec des prix mis à jour
- Quand les formules de calcul changent
- Quand de nouvelles hypothèses sont ajoutées

**⚠️ CONSÉQUENCE D'UN CHANGEMENT :**
Si vous remplacez le fichier de base par un nouveau fichier :
- **TOUS les fichiers Excel doivent être régénérés** (étape 2)
- **TOUS les prix doivent être recalculés** (étape 4)
- **TOUTES les données doivent être réextraites** (étape 5)

Le script principal (`calculateur_prix_camflex.py`) vous demandera confirmation avant de régénérer tout.

**Comment mettre à jour :**
1. Placez le nouveau fichier Excel Camflex dans `fichier de base/`
2. Renommez-le en `nepastoucher.xlsx`
3. Lancez `python calculateur_prix_camflex.py`
4. Le script détectera le changement et vous proposera de régénérer tout

---

### 2. Le Dossier Résultats (`résultats/`)

**Rôle :** Contient **TOUS les fichiers Excel générés** pour chaque variant d'abrivélo.

**Structure :**
- Un sous-dossier par **type d'abri vélo** (carport, bosquet_ferme, etc.)
- Dans chaque sous-dossier, un fichier Excel par **variant** (ex: `CAR-2.5M-N-200-G.xlsx`)

**Contenu de chaque fichier Excel :**
- Copie du fichier de base avec des paramètres spécifiques au variant
- Formules Excel qui calculent le prix en fonction des paramètres
- Feuille "PRC import" qui contient :
  - Prix avant réduction (cellule H7)
  - Prix après réduction (cellule H9)
  - Liste des composants (lignes A2:E110)

**Utilisation :**
- Ces fichiers permettent de **vérifier manuellement** chaque variant
- Ils servent de **source pour l'extraction** des prix et composants
- Ils seront **hébergés sur SharePoint Drive** pour être accessibles depuis Odoo

**⚠️ IMPORTANT :**
- Ces fichiers doivent être **ouverts dans Excel** pour que les formules se calculent
- Le script `extract_prices_and_components.py` fait cela automatiquement
- Ne modifiez pas manuellement ces fichiers, ils sont régénérés automatiquement

---

### 3. Les Scripts Python de Génération (`generate_*.py`)

**Principe fondamental :** **Un script Python = Un type d'abri vélo**

**Rôle de chaque script :**
- Prend le fichier de base (`nepastoucher.xlsx`)
- Génère tous les variants possibles pour ce type d'abri
- Crée un fichier Excel par variant dans `résultats/{type_abri}/`

**Scripts disponibles :**
- `generate_carport.py` → Génère tous les variants Carport
- `generate_bosquet_ferme.py` → Génère tous les variants Bosquet Fermé
- `generate_bosquet_ferme_compact.py` → Génère tous les variants Bosquet Fermé Compact
- `generate_bosquet_ouvert.py` → Génère tous les variants Bosquet Ouvert
- `generate_domino_ferme.py` → Génère tous les variants Domino Fermé
- `generate_domino_ferme_compact.py` → Génère tous les variants Domino Fermé Compact
- `generate_domino_ouvert.py` → Génère tous les variants Domino Ouvert
- `generate_metallique_ferme.py` → Génère tous les variants Métallique Fermé
- `generate_metallique_ferme_compact.py` → Génère tous les variants Métallique Fermé Compact
- `generate_metallique_ouvert.py` → Génère tous les variants Métallique Ouvert
- `generate_neve_ouvert.py` → Génère tous les variants Neve Ouvert

**Comment fonctionne un script de génération :**
1. Lit le fichier de base
2. Définit tous les paramètres possibles pour ce type d'abri :
   - Longueurs (2M, 2.5M, 4M, 5M, 6M, etc.)
   - Types (N = Normal, P = Premium)
   - Largeurs (200, 250, 400, etc.)
   - Couleurs (G = Gris, PT = Peinture, etc.)
3. Pour chaque combinaison de paramètres :
   - Crée une copie du fichier de base
   - Modifie les paramètres dans les cellules appropriées
   - Sauvegarde dans `résultats/{type_abri}/{NOM_FICHIER}.xlsx`

**Pour créer un nouveau type d'abri :**
1. Copiez un script existant (ex: `generate_carport.py`)
2. Renommez-le (ex: `generate_nouveau_type.py`)
3. Modifiez les paramètres dans le script :
   - Les longueurs possibles
   - Les types possibles
   - Les largeurs possibles
   - Les couleurs possibles
   - Le nom du dossier de sortie
4. Ajoutez le script à la liste dans `calculateur_prix_camflex.py` (variable `GENERATION_SCRIPTS`)

**Pour modifier les variants d'un type existant :**
1. Ouvrez le script correspondant (ex: `generate_carport.py`)
2. Modifiez les listes de paramètres :
   ```python
   LONGUEURS = ['2M', '2.5M', '4M', '5M', '6M', ...]  # Ajoutez/supprimez des longueurs
   TYPES = ['N', 'P']  # Ajoutez/supprimez des types
   LARGEURS = [200, 250, 400, ...]  # Ajoutez/supprimez des largeurs
   COULEURS = ['G', 'PT']  # Ajoutez/supprimez des couleurs
   ```
3. Relancez le script ou le script principal

**⚠️ IMPORTANT :**
- Chaque modification d'un script nécessite de **régénérer tous les fichiers** de ce type
- Le script principal vous proposera de régénérer automatiquement

---

### 4. Le Script Principal (`calculateur_prix_camflex.py`)

**Rôle :** Guide interactif qui automatise tout le processus.

**Ce qu'il fait :**
1. **Vérifie le fichier de base** et demande confirmation
2. **Génère tous les fichiers Excel** en lançant tous les scripts `generate_*.py`
3. **Extrait les prix et composants** en lançant `extract_prices_and_components.py`
4. **Affiche un résumé** des résultats finaux

**Utilisation :**
```bash
python calculateur_prix_camflex.py
```

Le script vous pose des questions à chaque étape :
- Voulez-vous utiliser ce fichier de base ?
- Voulez-vous régénérer tous les fichiers Excel ?
- Voulez-vous réextraire tous les prix ?

**Avantages :**
- Processus guidé, pas besoin de connaître tous les scripts
- Détection automatique des fichiers déjà générés
- Possibilité de reprendre après interruption

---

### 5. Le Script d'Extraction (`extract_prices_and_components.py`)

**Rôle :** Extrait les prix et composants depuis tous les fichiers Excel générés.

**Ce qu'il fait :**
1. Parcourt tous les fichiers Excel dans `résultats/`
2. Pour chaque fichier :
   - Ouvre le fichier dans Microsoft Excel (nécessaire pour calculer les formules)
   - Force le recalcul de toutes les formules
   - Lit les prix depuis la feuille "PRC import" :
     - Prix avant réduction : cellule H7
     - Prix après réduction : cellule H9
   - Lit les composants : lignes A2:E110 de la feuille "PRC import"
   - Sauvegarde et ferme le fichier
3. Génère deux types de fichiers :
   - `resultats_tous.json` : Tous les prix de tous les abrivélos
   - `composant/{type_abri}/{fichier}.json` : Composants détaillés par fichier

**⚠️ IMPORTANT :**
- **Microsoft Excel doit être installé** sur le système
- Cette étape peut prendre **plusieurs heures** (2-4h pour ~1600 fichiers)
- Les fichiers sont traités en parallèle pour accélérer

**Format des résultats :**

`resultats_tous.json` :
```json
{
  "date": "2024-01-01 12:00:00",
  "date_derniere_maj": "2024-01-01 12:00:00",
  "total": 1600,
  "resultats": [
    {
      "fichier": "CAR-2.5M-N-200-G.xlsx",
      "chemin_complet": "résultats/carport/CAR-2.5M-N-200-G.xlsx",
      "type_abri": "carport",
      "prix_avant_reduction": 1234.56,
      "prix_apres_reduction": 802.46,
      "date_extraction": "2024-01-01 12:00:00"
    }
  ]
}
```

`composant/{type_abri}/{fichier}.json` :
```json
{
  "fichier_source": "CAR-2.5M-N-200-G.xlsx",
  "chemin_source": "résultats/carport/CAR-2.5M-N-200-G.xlsx",
  "date_extraction": "2024-01-01 12:00:00",
  "composants": [
    ["Composant 1", "Référence", "Quantité", "Prix unitaire", "Prix total"],
    ["Composant 2", "REF-002", "5", "10.50", "52.50"],
    ...
  ]
}
```

---

### 6. Le Générateur d'URLs SharePoint (`generate_drive_urls.py`)

**Rôle :** Génère les URLs SharePoint Drive pour tous les fichiers Excel hébergés.

**Pourquoi c'est important :**
- Les fichiers Excel doivent être **hébergés sur SharePoint Drive** pour être accessibles depuis Odoo
- Odoo a besoin de l'**URL de chaque fichier** pour y accéder
- Ce script génère automatiquement toutes les URLs selon la structure SharePoint

**Comment ça fonctionne :**

La logique SharePoint Drive suit ce schéma :
- **Base dossiers** : `https://camflexsystems.sharepoint.com/:f:/r/sites/agentportal/ByCommute/Domino%20Tool/Tous_les_variants_bycommute/`
- **Base fichiers** : `https://camflexsystems.sharepoint.com/:x:/r/sites/agentportal/ByCommute/Domino%20Tool/Tous_les_variants_bycommute/`
- **URL dossier** : `{base_dossiers}{nom_dossier}?web=1`
- **URL fichier** : `{base_fichiers}{nom_dossier}/{nom_fichier}?web=1`

**Utilisation :**
```bash
python generate_drive_urls.py
```

**Ce qu'il fait :**
1. Parcourt le dossier `résultats/` et tous ses sous-dossiers
2. Pour chaque fichier trouvé :
   - Génère l'URL du dossier SharePoint
   - Génère l'URL du fichier SharePoint
3. Génère deux fichiers :
   - `urls_drive.csv` : Tableau CSV avec colonnes : Nom du dossier, Nom du fichier, URL du dossier, URL du fichier
   - `urls_drive.xlsx` : Même chose en format Excel, avec les bases d'URL en colonnes F et G

**⚠️ IMPORTANT - Vérification de l'URL de base :**

**Si les URLs ne fonctionnent pas :**
1. Vérifiez que l'URL de base dans le script correspond à la structure SharePoint actuelle
2. Ouvrez `generate_drive_urls.py`
3. Vérifiez les lignes 30-31 :
   ```python
   BASE_DOSSIERS = 'https://camflexsystems.sharepoint.com/:f:/r/sites/agentportal/ByCommute/Domino%20Tool/Tous_les_variants_bycommute/'
   BASE_FICHIERS = 'https://camflexsystems.sharepoint.com/:x:/r/sites/agentportal/ByCommute/Domino%20Tool/Tous_les_variants_bycommute/'
   ```
4. Si l'URL a changé dans SharePoint :
   - Modifiez ces deux lignes avec la nouvelle URL
   - Relancez le script pour régénérer les URLs

**Si l'URL n'a pas changé :**
- Ne modifiez rien, laissez les URLs telles quelles
- Le script fonctionne correctement

**Format de sortie :**

`urls_drive.csv` :
```csv
Nom du dossier;Nom du fichier;URL du dossier;URL du fichier
carport;CAR-2.5M-N-200-G.xlsx;https://camflexsystems.sharepoint.com/:f:/r/sites/agentportal/ByCommute/Domino%20Tool/Tous_les_variants_bycommute/carport?web=1;https://camflexsystems.sharepoint.com/:x:/r/sites/agentportal/ByCommute/Domino%20Tool/Tous_les_variants_bycommute/carport/CAR-2.5M-N-200-G.xlsx?web=1
```

---

## 🔄 Processus Complet : Du Fichier de Base aux Données Odoo

### Étape 1 : Préparation du Fichier de Base

1. Placez le fichier Excel Camflex dans `fichier de base/`
2. Renommez-le en `nepastoucher.xlsx`
3. Vérifiez que le fichier contient bien :
   - Les formules de calcul
   - La feuille "PRC import" avec les cellules H7 et H9 pour les prix

### Étape 2 : Génération de Tous les Excel

```bash
python calculateur_prix_camflex.py
```

Le script va :
1. Vérifier le fichier de base
2. Lancer tous les scripts `generate_*.py`
3. Créer tous les fichiers Excel dans `résultats/`

**Résultat :** ~1600 fichiers Excel générés, organisés par type d'abri

### Étape 3 : Calcul des Formules Excel

Le script `extract_prices_and_components.py` (lancé automatiquement) va :
1. Ouvrir chaque fichier Excel dans Microsoft Excel
2. Forcer le recalcul de toutes les formules
3. Sauvegarder et fermer chaque fichier

**⚠️ Cette étape prend 2-4 heures** (nécessite Excel installé)

### Étape 4 : Extraction des Prix et Composants

Toujours via `extract_prices_and_components.py` :
1. Lit les prix depuis chaque fichier Excel
2. Extrait les composants détaillés
3. Génère `resultats_tous.json` et les fichiers dans `composant/`

**Résultat :**
- `resultats_tous.json` : Tous les prix (avant/après réduction)
- `composant/{type_abri}/` : Composants détaillés par fichier

### Étape 5 : Upload des Fichiers sur SharePoint Drive

**⚠️ Action manuelle requise :**

1. Upload tous les fichiers Excel de `résultats/` sur SharePoint Drive
2. Structure à respecter :
   ```
   SharePoint/ByCommute/Domino Tool/Tous_les_variants_bycommute/
   ├── carport/
   │   ├── CAR-2.5M-N-200-G.xlsx
   │   └── ...
   ├── bosquet_ferme/
   │   └── ...
   └── ...
   ```
3. Vérifiez que la structure correspond exactement aux dossiers dans `résultats/`

### Étape 6 : Génération des URLs SharePoint

```bash
python generate_drive_urls.py
```

**⚠️ Vérifiez d'abord l'URL de base :**
- Ouvrez `generate_drive_urls.py`
- Vérifiez que les URLs en lignes 30-31 correspondent à votre SharePoint
- Si l'URL a changé, modifiez-la avant de lancer

Le script génère :
- `urls_drive.csv` : Tableau avec toutes les URLs
- `urls_drive.xlsx` : Même chose en Excel

### Étape 7 : Préparation des Données pour Odoo

**Données nécessaires pour Odoo :**
1. **Prix unitaire après réduction** → Disponible dans `resultats_tous.json` (champ `prix_apres_reduction`)
2. **URL du fichier Excel** → Disponible dans `urls_drive.csv/xlsx` (colonne "URL du fichier")
3. **Liste des composants** → Disponible dans `composant/{type_abri}/{fichier}.json`

**Format pour Odoo :**
- Pour chaque variant d'abrivélo :
  - Nom du variant (ex: "CAR-2.5M-N-200-G")
  - Prix unitaire après réduction
  - URL du fichier Excel sur SharePoint
  - Liste des composants (référence, quantité, prix unitaire, prix total)

### Étape 8 : Upload dans Odoo

**Action manuelle requise :**

1. Utilisez les données de `resultats_tous.json` et `urls_drive.csv`
2. Pour chaque variant :
   - Créez/Modifiez l'enregistrement dans Odoo
   - Ajoutez le prix unitaire après réduction
   - Ajoutez l'URL du fichier Excel
   - Ajoutez la liste des composants

**⚠️ IMPORTANT :**
- Vérifiez que tous les fichiers sont bien uploadés sur SharePoint avant d'ajouter les URLs dans Odoo
- Testez quelques URLs pour vérifier qu'elles fonctionnent
- Si une URL ne fonctionne pas, vérifiez l'URL de base dans `generate_drive_urls.py`

---

## 📊 Cas d'Usage Détaillés

### Cas d'Usage 1 : Générer Tous les Prix pour la Première Fois

**Objectif :** Partir du fichier de base et obtenir tous les prix et URLs pour Odoo.

**Étapes :**
1. Placez `nepastoucher.xlsx` dans `fichier de base/`
2. Lancez `python calculateur_prix_camflex.py`
3. Répondez "Oui" à toutes les questions
4. Attendez la fin du processus (plusieurs heures)
5. Upload tous les fichiers de `résultats/` sur SharePoint Drive
6. Lancez `python generate_drive_urls.py`
7. Vérifiez les URLs générées
8. Utilisez `resultats_tous.json` et `urls_drive.csv` pour uploader dans Odoo

**Résultat :** Tous les prix, URLs et composants prêts pour Odoo

---

### Cas d'Usage 2 : Mettre à Jour les Prix (Nouveau Fichier de Base)

**Objectif :** Quand Camflex fournit un nouveau fichier avec des prix mis à jour.

**Étapes :**
1. Remplacez `fichier de base/nepastoucher.xlsx` par le nouveau fichier
2. Lancez `python calculateur_prix_camflex.py`
3. Le script détectera le changement et vous demandera confirmation
4. Choisissez de régénérer tous les fichiers Excel
5. Le script va :
   - Régénérer tous les Excel (étape 2)
   - Recalculer tous les prix (étape 3)
   - Réextraire tous les prix (étape 4)
6. Upload les nouveaux fichiers sur SharePoint (remplacez les anciens)
7. Relancez `python generate_drive_urls.py` pour régénérer les URLs
8. Mettez à jour Odoo avec les nouveaux prix

**⚠️ IMPORTANT :**
- Tous les fichiers Excel seront régénérés
- Tous les prix seront recalculés
- Les URLs resteront les mêmes (si la structure SharePoint n'a pas changé)

---

### Cas d'Usage 3 : Ajouter un Nouveau Type d'Abri

**Objectif :** Créer un nouveau type d'abri vélo (ex: "nouveau_type").

**Étapes :**
1. Copiez un script existant (ex: `generate_carport.py`)
2. Renommez-le (ex: `generate_nouveau_type.py`)
3. Modifiez le script :
   - Changez le nom du dossier de sortie
   - Modifiez les paramètres (longueurs, types, largeurs, couleurs)
   - Adaptez la logique de génération si nécessaire
4. Ajoutez le script à `calculateur_prix_camflex.py` :
   ```python
   GENERATION_SCRIPTS = [
       'generate_carport.py',
       ...
       'generate_nouveau_type.py',  # Ajoutez cette ligne
   ]
   ```
5. Lancez `python calculateur_prix_camflex.py`
6. Le script générera les nouveaux fichiers Excel
7. Suivez les étapes 3-8 du processus complet

**Résultat :** Nouveau type d'abri avec tous ses variants générés

---

### Cas d'Usage 4 : Modifier les Variants d'un Type Existant

**Objectif :** Ajouter/supprimer des variants pour un type d'abri existant.

**Exemple :** Ajouter la longueur "15M" au type "carport".

**Étapes :**
1. Ouvrez `generate_carport.py`
2. Trouvez la liste des longueurs :
   ```python
   LONGUEURS = ['2M', '2.5M', '4M', '5M', '6M', ...]
   ```
3. Ajoutez '15M' :
   ```python
   LONGUEURS = ['2M', '2.5M', '4M', '5M', '6M', ..., '15M']
   ```
4. Sauvegardez le fichier
5. Lancez `python calculateur_prix_camflex.py`
6. Choisissez de régénérer tous les fichiers Excel
7. Le script générera les nouveaux variants
8. Suivez les étapes 3-8 du processus complet

**⚠️ IMPORTANT :**
- Tous les fichiers Excel de ce type seront régénérés
- Les anciens variants resteront, les nouveaux seront ajoutés

---

### Cas d'Usage 5 : Générer les URLs SharePoint (Après Upload)

**Objectif :** Générer les URLs SharePoint après avoir uploadé les fichiers.

**Étapes :**
1. **Vérifiez d'abord l'URL de base :**
   - Ouvrez `generate_drive_urls.py`
   - Vérifiez les lignes 30-31
   - Si l'URL SharePoint a changé, modifiez-la
   - Si l'URL n'a pas changé, ne modifiez rien
2. Lancez `python generate_drive_urls.py`
3. Vérifiez les fichiers générés :
   - `urls_drive.csv` : Ouvrez dans Excel/LibreOffice
   - `urls_drive.xlsx` : Ouvrez dans Excel
4. Testez quelques URLs manuellement :
   - Ouvrez une URL dans un navigateur
   - Vérifiez qu'elle pointe vers le bon fichier
5. Si les URLs ne fonctionnent pas :
   - Vérifiez que la structure SharePoint correspond
   - Vérifiez que l'URL de base est correcte
   - Modifiez l'URL de base si nécessaire et relancez

**⚠️ IMPORTANT :**
- Les URLs doivent être générées **après** l'upload sur SharePoint
- Si la structure SharePoint change, il faut mettre à jour l'URL de base
- Testez toujours quelques URLs avant d'utiliser le fichier complet

---

### Cas d'Usage 6 : Reconstruire la Logique Odoo

**Objectif :** Après avoir uploadé tous les fichiers sur SharePoint, reconstruire la logique dans Odoo.

**Données disponibles :**
1. `resultats_tous.json` : Tous les prix (avant/après réduction)
2. `urls_drive.csv` : Toutes les URLs SharePoint
3. `composant/{type_abri}/` : Tous les composants détaillés

**Étapes :**
1. Parsez `resultats_tous.json` pour obtenir les prix
2. Parsez `urls_drive.csv` pour obtenir les URLs
3. Pour chaque variant :
   - Récupérez le prix après réduction depuis `resultats_tous.json`
   - Récupérez l'URL depuis `urls_drive.csv`
   - Récupérez les composants depuis `composant/{type_abri}/{fichier}.json`
   - Créez/Modifiez l'enregistrement dans Odoo avec ces données

**Format des données pour Odoo :**
```json
{
  "variant": "CAR-2.5M-N-200-G",
  "type_abri": "carport",
  "prix_unitaire_apres_reduction": 802.46,
  "url_fichier_excel": "https://camflexsystems.sharepoint.com/:x:/r/sites/agentportal/ByCommute/Domino%20Tool/Tous_les_variants_bycommute/carport/CAR-2.5M-N-200-G.xlsx?web=1",
  "composants": [
    {
      "nom": "Composant 1",
      "reference": "REF-001",
      "quantite": 5,
      "prix_unitaire": 10.50,
      "prix_total": 52.50
    },
    ...
  ]
}
```

---

## ⚠️ Points d'Attention Critiques

### 1. Le Fichier de Base

- **NE JAMAIS MODIFIER DIRECTEMENT** le fichier de base
- Si vous le remplacez, **TOUT doit être régénéré**
- Vérifiez toujours que le nouveau fichier a la même structure

### 2. Les URLs SharePoint

- **Vérifiez l'URL de base** avant de générer les URLs
- Si l'URL SharePoint change, modifiez-la dans `generate_drive_urls.py`
- **Testez toujours quelques URLs** avant d'utiliser le fichier complet
- Les URLs doivent être générées **après** l'upload sur SharePoint

### 3. Les Scripts Python

- **Un script = Un type d'abri** : Ne modifiez pas un script pour changer un autre type
- Pour créer un nouveau type, **copiez un script existant** et modifiez-le
- Pour modifier les variants, **modifiez les listes de paramètres** dans le script

### 4. Le Calcul des Formules Excel

- **Microsoft Excel doit être installé** pour que les formules se calculent
- Cette étape prend **plusieurs heures** (2-4h)
- Ne fermez pas Excel pendant le processus
- Si le processus est interrompu, vous pouvez le relancer (il reprend où il s'est arrêté)

### 5. L'Upload sur SharePoint

- **Respectez la structure exacte** des dossiers
- La structure SharePoint doit correspondre à `résultats/`
- Upload tous les fichiers **avant** de générer les URLs

---

## 🔧 Dépannage

### Problème : Excel n'est pas installé

**Erreur :** `❌ Excel n'est pas installé`

**Solution :**
- Installez Microsoft Excel
- Le script nécessite Excel pour calculer les formules
- Alternative : Utilisez LibreOffice (mais peut nécessiter des modifications du script)

---

### Problème : Aucun prix calculé

**Erreur :** `⚠️  Aucun prix n'a été calculé`

**Solution :**
1. Vérifiez que les fichiers Excel ont bien été ouverts dans Excel
2. Vérifiez que les formules se sont bien calculées
3. Relancez l'extraction : `python extract_prices_and_components.py`
4. Vérifiez que la feuille "PRC import" existe dans les fichiers Excel
5. Vérifiez que les cellules H7 et H9 contiennent bien les prix

---

### Problème : Fichier de base introuvable

**Erreur :** `❌ Le fichier de base n'existe pas`

**Solution :**
1. Vérifiez que le fichier est bien dans `fichier de base/nepastoucher.xlsx`
2. Vérifiez l'orthographe du nom du fichier
3. Vérifiez que le fichier n'est pas corrompu

---

### Problème : Les URLs SharePoint ne fonctionnent pas

**Symptôme :** Les URLs générées ne pointent pas vers les bons fichiers

**Solution :**
1. Ouvrez `generate_drive_urls.py`
2. Vérifiez les URLs de base (lignes 30-31)
3. Vérifiez que la structure SharePoint correspond à la structure dans `résultats/`
4. Si l'URL SharePoint a changé, modifiez-la dans le script
5. Relancez `python generate_drive_urls.py`
6. Testez quelques URLs manuellement

---

### Problème : Script de génération échoue

**Erreur :** Un script `generate_*.py` échoue

**Solution :**
1. Vérifiez que le fichier de base existe et n'est pas corrompu
2. Vérifiez les paramètres dans le script (longueurs, types, etc.)
3. Vérifiez que le dossier de sortie existe
4. Vérifiez les permissions d'écriture
5. Regardez les logs d'erreur pour plus de détails

---

## 📝 Notes Finales

Ce calculateur de prix est un système complet qui automatise la génération des prix pour tous les variants d'abrivélos ByCommute. Il permet de :

1. **Générer automatiquement** tous les fichiers Excel
2. **Calculer automatiquement** tous les prix
3. **Extraire automatiquement** tous les prix et composants
4. **Générer automatiquement** toutes les URLs SharePoint
5. **Préparer les données** pour l'intégration Odoo

**Le processus complet** prend plusieurs heures mais est entièrement automatisé. Une fois configuré, il suffit de lancer le script principal et d'attendre la fin du processus.

**Pour toute modification** (nouveau type d'abri, nouveaux variants, nouveau fichier de base), suivez les cas d'usage correspondants dans cette documentation.

---

## 📞 Support

Pour toute question ou problème :
1. Consultez cette documentation complète
2. Vérifiez les sections de dépannage
3. Vérifiez les cas d'usage correspondants
4. Vérifiez les commentaires dans les scripts Python

---

**Dernière mise à jour :** 2024
