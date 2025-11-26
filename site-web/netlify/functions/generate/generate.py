import json
import os
import shutil
import subprocess
import tempfile
import zipfile
from pathlib import Path

def handler(event, context):
    """
    Fonction Netlify pour générer toutes les variantes d'abris vélos
    """
    try:
        # Créer un dossier temporaire
        temp_dir = tempfile.mkdtemp()
        
        # Chemin vers les scripts Python (à adapter selon votre structure)
        scripts_dir = Path(__file__).parent.parent.parent.parent.parent / 'generate_*.py'
        
        # Liste des scripts à exécuter
        scripts = [
            'generate_bosquet_ouvert.py',
            'generate_bosquet_ferme.py',
            'generate_bosquet_ouvert_compact.py',
            'generate_bosquet_ferme_compact.py',
            'generate_domino_ouvert.py',
            'generate_domino_ferme.py',
            'generate_domino_ouvert_compact.py',
            'generate_domino_ferme_compact.py',
            'generate_metallique_ouvert.py',
            'generate_metallique_ferme.py',
            'generate_metallique_ouvert_compact.py',
            'generate_metallique_ferme_compact.py',
            'generate_neve_ouvert.py',
            'generate_neve_ferme.py',
            'generate_neve_ferme_compact.py'
        ]
        
        # Créer la structure de dossiers
        base_dir = Path(temp_dir) / 'fichier de base'
        resultats_dir = Path(temp_dir) / 'résultats'
        base_dir.mkdir(parents=True, exist_ok=True)
        resultats_dir.mkdir(parents=True, exist_ok=True)
        
        # Gérer le fichier uploadé
        if event.get('body'):
            # Le fichier devrait être dans event['body'] (base64 ou binaire)
            # Pour l'instant, on utilise le fichier par défaut
            pass
        
        # Copier le fichier de base (à adapter)
        # source_file = Path('/path/to/nepastoucher.xlsx')
        # shutil.copy(source_file, base_dir / 'nepastoucher.xlsx')
        
        # Exécuter tous les scripts
        for script in scripts:
            try:
                script_path = Path(__file__).parent.parent.parent.parent.parent / script
                if script_path.exists():
                    subprocess.run(
                        ['python3', str(script_path)],
                        cwd=temp_dir,
                        check=True,
                        capture_output=True
                    )
            except subprocess.CalledProcessError as e:
                print(f"Erreur avec {script}: {e}")
                continue
        
        # Créer le fichier ZIP
        zip_path = Path(temp_dir) / 'resultats.zip'
        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for excel_file in resultats_dir.rglob('*.xlsx'):
                zipf.write(excel_file, excel_file.relative_to(resultats_dir))
        
        # Lire le ZIP en base64
        with open(zip_path, 'rb') as f:
            zip_data = f.read()
        
        # Nettoyer
        shutil.rmtree(temp_dir)
        
        return {
            'statusCode': 200,
            'headers': {
                'Content-Type': 'application/zip',
                'Content-Disposition': 'attachment; filename="abris-velos-variantes.zip"'
            },
            'body': zip_data,
            'isBase64Encoded': False
        }
        
    except Exception as e:
        return {
            'statusCode': 500,
            'body': json.dumps({'error': str(e)})
        }

