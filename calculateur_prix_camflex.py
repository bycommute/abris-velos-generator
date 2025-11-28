#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Calculateur Prix Camflex - Script Principal
============================================

Ce script guide l'utilisateur à travers tout le processus de génération des prix
des abrivélos à partir du fichier de base Camflex.

Processus :
1. Vérification/Configuration du fichier de base
2. Génération des fichiers Excel pour chaque variant d'abrivélo
3. Calcul des formules Excel (ouverture dans Excel)
4. Extraction des prix et composants depuis les Excel
5. Génération du fichier final resultats_tous.json

Utilisation :
    python calculateur_prix_camflex.py
"""

import os
import sys
import subprocess
import shutil
import json
from datetime import datetime
from pathlib import Path

# Configuration
BASE_DIR = 'fichier de base'
SOURCE_FILE = os.path.join(BASE_DIR, 'nepastoucher.xlsx')
RESULTATS_DIR = 'résultats'
COMPOSANT_DIR = 'composant'
RESULTATS_JSON = 'resultats_tous.json'

# Liste de tous les scripts de génération
GENERATION_SCRIPTS = [
    'generate_carport.py',
    'generate_bosquet_ferme.py',
    'generate_bosquet_ferme_compact.py',
    'generate_bosquet_ouvert.py',
    'generate_domino_ferme.py',
    'generate_domino_ferme_compact.py',
    'generate_domino_ouvert.py',
    'generate_metallique_ferme.py',
    'generate_metallique_ferme_compact.py',
    'generate_metallique_ouvert.py',
    'generate_neve_ouvert.py',
]

def print_header(title):
    """Affiche un en-tête formaté"""
    print("\n" + "=" * 80)
    print(title)
    print("=" * 80)

def print_section(title):
    """Affiche une section formatée"""
    print(f"\n{'─' * 80}")
    print(f"  {title}")
    print(f"{'─' * 80}")

def demander_oui_non(question, defaut=True):
    """Pose une question oui/non à l'utilisateur"""
    reponse_defaut = "O/n" if defaut else "o/N"
    while True:
        reponse = input(f"{question} [{reponse_defaut}] : ").strip().lower()
        if not reponse:
            return defaut
        if reponse in ['o', 'oui', 'y', 'yes']:
            return True
        if reponse in ['n', 'non', 'no']:
            return False
        print("   ⚠️  Réponse invalide. Répondez 'o' pour oui ou 'n' pour non.")

def verifier_fichier_base():
    """Vérifie l'existence du fichier de base et demande confirmation"""
    print_header("ÉTAPE 1 : VÉRIFICATION DU FICHIER DE BASE")
    
    if not os.path.exists(SOURCE_FILE):
        print(f"\n❌ Le fichier de base n'existe pas : {SOURCE_FILE}")
        print("\n📝 Pour continuer, vous devez :")
        print(f"   1. Placer votre fichier Excel Camflex dans le dossier '{BASE_DIR}/'")
        print(f"   2. Le renommer en 'nepastoucher.xlsx'")
        return False
    
    # Afficher les informations du fichier
    file_size = os.path.getsize(SOURCE_FILE)
    file_size_mb = file_size / (1024 * 1024)
    modif_time = datetime.fromtimestamp(os.path.getmtime(SOURCE_FILE))
    
    print(f"\n📄 Fichier de base trouvé : {SOURCE_FILE}")
    print(f"   Taille : {file_size_mb:.2f} Mo")
    print(f"   Dernière modification : {modif_time.strftime('%Y-%m-%d %H:%M:%S')}")
    
    # Demander confirmation
    utiliser_fichier = demander_oui_non(
        "\n✅ Voulez-vous utiliser ce fichier de base pour générer les prix ?",
        defaut=True
    )
    
    if not utiliser_fichier:
        print("\n📝 Pour mettre à jour le fichier de base :")
        print(f"   1. Remplacez le fichier dans '{BASE_DIR}/nepastoucher.xlsx'")
        print(f"   2. Relancez ce script")
        return False
    
    # Demander si on veut mettre à jour le fichier
    mettre_a_jour = demander_oui_non(
        "\n🔄 Voulez-vous remplacer le fichier de base par un nouveau fichier ?",
        defaut=False
    )
    
    if mettre_a_jour:
        nouveau_fichier = input("\n📁 Entrez le chemin complet du nouveau fichier Excel : ").strip()
        
        if not nouveau_fichier:
            print("   ⚠️  Aucun fichier spécifié. Utilisation du fichier existant.")
            return True
        
        if not os.path.exists(nouveau_fichier):
            print(f"   ❌ Le fichier n'existe pas : {nouveau_fichier}")
            return False
        
        if not nouveau_fichier.endswith('.xlsx'):
            print("   ⚠️  Le fichier doit être un fichier Excel (.xlsx)")
            return False
        
        # Créer une sauvegarde de l'ancien fichier
        backup_file = f"{SOURCE_FILE}.backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
        if os.path.exists(SOURCE_FILE):
            shutil.copy2(SOURCE_FILE, backup_file)
            print(f"   💾 Ancien fichier sauvegardé : {backup_file}")
        
        # Copier le nouveau fichier
        shutil.copy2(nouveau_fichier, SOURCE_FILE)
        print(f"   ✅ Fichier de base mis à jour : {SOURCE_FILE}")
        print("\n   ⚠️  ATTENTION : Si vous avez changé le fichier de base,")
        print("      vous devrez régénérer tous les fichiers Excel.")
    
    return True

def verifier_scripts_generation():
    """Vérifie que tous les scripts de génération existent"""
    print_section("Vérification des scripts de génération")
    
    scripts_manquants = []
    for script in GENERATION_SCRIPTS:
        if not os.path.exists(script):
            scripts_manquants.append(script)
        else:
            print(f"   ✅ {script}")
    
    if scripts_manquants:
        print(f"\n   ❌ Scripts manquants : {', '.join(scripts_manquants)}")
        return False
    
    print(f"\n   ✅ Tous les {len(GENERATION_SCRIPTS)} scripts de génération sont présents")
    return True

def generer_tous_excel():
    """Génère tous les fichiers Excel pour chaque variant"""
    print_header("ÉTAPE 2 : GÉNÉRATION DES FICHIERS EXCEL")
    
    # Vérifier les scripts
    if not verifier_scripts_generation():
        return False
    
    # Demander si on veut régénérer
    excel_existants = compter_fichiers_excel()
    if excel_existants > 0:
        print(f"\n📊 {excel_existants} fichiers Excel existent déjà dans '{RESULTATS_DIR}/'")
        regenerer = demander_oui_non(
            "🔄 Voulez-vous régénérer tous les fichiers Excel ?",
            defaut=False
        )
        if not regenerer:
            print("\n   ⏭️  Utilisation des fichiers Excel existants")
            return True
    
    print(f"\n🚀 Génération des fichiers Excel pour {len(GENERATION_SCRIPTS)} types d'abrivélos...")
    print("   (Cela peut prendre plusieurs minutes)\n")
    
    succes = 0
    echecs = 0
    
    for i, script in enumerate(GENERATION_SCRIPTS, 1):
        print(f"[{i}/{len(GENERATION_SCRIPTS)}] 📝 Génération avec {script}...")
        
        try:
            result = subprocess.run(
                [sys.executable, script],
                capture_output=True,
                text=True,
                timeout=300
            )
            
            if result.returncode == 0:
                print(f"   ✅ {script} : Succès")
                succes += 1
            else:
                print(f"   ⚠️  {script} : Avertissements (code {result.returncode})")
                if result.stderr:
                    print(f"      {result.stderr[:200]}")
                succes += 1  # On continue même avec des avertissements
                
        except subprocess.TimeoutExpired:
            print(f"   ⚠️  {script} : Timeout (trop long)")
            echecs += 1
        except Exception as e:
            print(f"   ❌ {script} : Erreur - {e}")
            echecs += 1
    
    print(f"\n📊 Résumé : {succes} succès, {echecs} échecs")
    
    if echecs > 0:
        continuer = demander_oui_non(
            "\n⚠️  Certains scripts ont échoué. Voulez-vous continuer quand même ?",
            defaut=True
        )
        return continuer
    
    return True

def compter_fichiers_excel():
    """Compte le nombre de fichiers Excel dans le dossier résultats"""
    count = 0
    if os.path.exists(RESULTATS_DIR):
        for root, dirs, files in os.walk(RESULTATS_DIR):
            for file in files:
                if file.endswith('.xlsx') and not file.startswith('~'):
                    count += 1
    return count

def extraire_prix_et_composants():
    """Extrait les prix et composants depuis les fichiers Excel"""
    print_header("ÉTAPE 3 : EXTRACTION DES PRIX ET COMPOSANTS")
    
    excel_count = compter_fichiers_excel()
    if excel_count == 0:
        print("\n❌ Aucun fichier Excel trouvé dans le dossier résultats")
        print("   Vous devez d'abord générer les fichiers Excel (Étape 2)")
        return False
    
    print(f"\n📊 {excel_count} fichiers Excel trouvés")
    
    # Vérifier si extract_prices_and_components.py existe
    script_extraction = 'extract_prices_and_components.py'
    if not os.path.exists(script_extraction):
        print(f"\n❌ Script d'extraction introuvable : {script_extraction}")
        return False
    
    # Demander si on veut réextraire
    if os.path.exists(RESULTATS_JSON):
        print(f"\n📄 Fichier de résultats existant : {RESULTATS_JSON}")
        reextraire = demander_oui_non(
            "🔄 Voulez-vous réextraire tous les prix ? (sinon, seuls les nouveaux fichiers seront traités)",
            defaut=False
        )
        if not reextraire:
            print("\n   ⏭️  Extraction uniquement des nouveaux fichiers")
    
    print(f"\n🚀 Extraction des prix et composants...")
    print("   (Cette étape ouvre chaque fichier Excel pour calculer les formules)")
    print("   (Cela peut prendre beaucoup de temps selon le nombre de fichiers)\n")
    
    continuer = demander_oui_non(
        "⚠️  Cette étape va ouvrir Excel et traiter tous les fichiers. Continuer ?",
        defaut=True
    )
    
    if not continuer:
        return False
    
    try:
        result = subprocess.run(
            [sys.executable, script_extraction],
            text=True,
            timeout=3600  # 1 heure max
        )
        
        if result.returncode == 0:
            print("\n✅ Extraction terminée avec succès")
            return True
        else:
            print(f"\n⚠️  Extraction terminée avec des avertissements (code {result.returncode})")
            return True  # On continue même avec des avertissements
            
    except subprocess.TimeoutExpired:
        print("\n❌ L'extraction a pris trop de temps")
        return False
    except Exception as e:
        print(f"\n❌ Erreur lors de l'extraction : {e}")
        return False

def afficher_resultats_finaux():
    """Affiche un résumé des résultats finaux"""
    print_header("RÉSULTATS FINAUX")
    
    if not os.path.exists(RESULTATS_JSON):
        print(f"\n❌ Fichier de résultats introuvable : {RESULTATS_JSON}")
        return
    
    try:
        with open(RESULTATS_JSON, 'r', encoding='utf-8') as f:
            data = json.load(f)
        
        resultats = data.get('resultats', [])
        total = len(resultats)
        
        # Compter les fichiers avec prix complets
        prix_complets = [
            r for r in resultats 
            if r.get('prix_avant_reduction') is not None 
            and r.get('prix_apres_reduction') is not None
        ]
        
        print(f"\n📊 Statistiques :")
        print(f"   Total de fichiers traités : {total}")
        print(f"   Fichiers avec prix complets : {len(prix_complets)}")
        print(f"   Fichiers sans prix : {total - len(prix_complets)}")
        
        if 'date_derniere_maj' in data:
            print(f"\n📅 Dernière mise à jour : {data['date_derniere_maj']}")
        
        print(f"\n💾 Fichiers générés :")
        print(f"   📄 {RESULTATS_JSON} : Tous les prix des abrivélos")
        print(f"   📁 {COMPOSANT_DIR}/ : Composants détaillés par type d'abrivélo")
        
        if len(prix_complets) > 0:
            print(f"\n✅ SUCCÈS : {len(prix_complets)} abrivélos avec prix calculés !")
        else:
            print(f"\n⚠️  ATTENTION : Aucun prix n'a été calculé.")
            print("   Vérifiez que les fichiers Excel ont bien été ouverts dans Excel.")
        
    except Exception as e:
        print(f"\n❌ Erreur lors de la lecture des résultats : {e}")

def main():
    """Fonction principale"""
    print_header("CALCULATEUR PRIX CAMFLEX")
    print("\nCe script vous guide à travers le processus complet de génération")
    print("des prix des abrivélos à partir du fichier de base Camflex.\n")
    
    # Étape 1 : Vérification du fichier de base
    if not verifier_fichier_base():
        print("\n❌ Impossible de continuer sans fichier de base valide")
        return
    
    # Étape 2 : Génération des fichiers Excel
    if not generer_tous_excel():
        print("\n❌ Échec lors de la génération des fichiers Excel")
        return
    
    # Étape 3 : Extraction des prix
    if not extraire_prix_et_composants():
        print("\n❌ Échec lors de l'extraction des prix")
        return
    
    # Résultats finaux
    afficher_resultats_finaux()
    
    print_header("PROCESSUS TERMINÉ")
    print("\n✅ Le calculateur a terminé avec succès !")
    print(f"\n📄 Fichier final : {RESULTATS_JSON}")
    print(f"📁 Composants : {COMPOSANT_DIR}/")
    print("\n💡 Vous pouvez maintenant utiliser ces fichiers pour votre application.")

if __name__ == '__main__':
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n⚠️  Processus interrompu par l'utilisateur")
        sys.exit(1)
    except Exception as e:
        print(f"\n\n❌ Erreur fatale : {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


