"""
API alternative pour générer les variantes
Peut être déployée séparément (Heroku, Railway, etc.)
"""
from flask import Flask, request, send_file
from flask_cors import CORS
import os
import shutil
import subprocess
import tempfile
import zipfile
from pathlib import Path

app = Flask(__name__)
CORS(app)

@app.route('/generate', methods=['POST'])
def generate():
    """Génère toutes les variantes et retourne un ZIP"""
    try:
        # Créer un dossier temporaire
        temp_dir = tempfile.mkdtemp()
        
        # Créer la structure
        base_dir = Path(temp_dir) / 'fichier de base'
        resultats_dir = Path(temp_dir) / 'résultats'
        base_dir.mkdir(parents=True, exist_ok=True)
        resultats_dir.mkdir(parents=True, exist_ok=True)
        
        # Gérer le fichier uploadé
        if 'file' in request.files:
            file = request.files['file']
            file.save(base_dir / 'nepastoucher.xlsx')
        else:
            # Utiliser le fichier par défaut
            default_file = Path(__file__).parent.parent / 'fichier de base' / 'nepastoucher.xlsx'
            if default_file.exists():
                shutil.copy(default_file, base_dir / 'nepastoucher.xlsx')
        
        # Scripts à exécuter
        scripts_dir = Path(__file__).parent.parent
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
        
        # Exécuter les scripts
        for script in scripts:
            script_path = scripts_dir / script
            if script_path.exists():
                try:
                    subprocess.run(
                        ['python3', str(script_path)],
                        cwd=temp_dir,
                        check=True,
                        capture_output=True
                    )
                except subprocess.CalledProcessError as e:
                    print(f"Erreur avec {script}: {e}")
                    continue
        
        # Créer le ZIP
        zip_path = Path(temp_dir) / 'resultats.zip'
        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for excel_file in resultats_dir.rglob('*.xlsx'):
                zipf.write(excel_file, excel_file.relative_to(resultats_dir))
        
        # Retourner le ZIP
        return send_file(
            zip_path,
            mimetype='application/zip',
            as_attachment=True,
            download_name='abris-velos-variantes.zip'
        )
        
    except Exception as e:
        return {'error': str(e)}, 500
    finally:
        # Nettoyer
        if os.path.exists(temp_dir):
            shutil.rmtree(temp_dir)

if __name__ == '__main__':
    app.run(debug=True, port=5000)

