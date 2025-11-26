#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Script pour pousser les modifications sur GitHub
Version améliorée avec meilleure gestion d'erreurs
"""
import subprocess
import os
import sys
from pathlib import Path

def run_command(cmd, cwd=None, check=True):
    """Exécute une commande et retourne le résultat"""
    try:
        result = subprocess.run(
            cmd,
            shell=True,
            cwd=cwd,
            capture_output=True,
            text=True,
            timeout=60
        )
        if check and result.returncode != 0:
            print(f"❌ Erreur lors de l'exécution: {cmd}")
            print(f"   Sortie: {result.stdout}")
            print(f"   Erreur: {result.stderr}")
            return False, result.stdout, result.stderr
        return result.returncode == 0, result.stdout, result.stderr
    except subprocess.TimeoutExpired:
        print(f"❌ Timeout lors de l'exécution: {cmd}")
        return False, "", "Timeout"
    except Exception as e:
        print(f"❌ Exception: {e}")
        return False, "", str(e)

def check_git_repo(cwd):
    """Vérifie qu'on est dans un repo git"""
    success, _, _ = run_command("git rev-parse --git-dir", cwd=cwd, check=False)
    return success

def get_current_branch(cwd):
    """Récupère la branche actuelle"""
    success, stdout, _ = run_command("git branch --show-current", cwd=cwd, check=False)
    if success:
        return stdout.strip()
    return None

def main():
    # Changer vers le répertoire du projet
    project_dir = Path(__file__).parent.absolute()
    os.chdir(project_dir)
    
    print("📦 Poussage des modifications sur GitHub...")
    print("=" * 60)
    
    # Vérifier qu'on est dans un repo git
    if not check_git_repo(project_dir):
        print("❌ Erreur: Ce répertoire n'est pas un repository Git")
        sys.exit(1)
    
    # Récupérer la branche actuelle
    branch = get_current_branch(project_dir)
    if not branch:
        print("❌ Erreur: Impossible de déterminer la branche actuelle")
        sys.exit(1)
    
    print(f"📍 Branche actuelle: {branch}")
    
    # Vérifier le statut
    print("\n1. Vérification du statut Git...")
    success, stdout, stderr = run_command("git status --short", cwd=project_dir, check=False)
    if success:
        if stdout.strip():
            print("📝 Fichiers modifiés:")
            print(stdout)
        else:
            print("ℹ️  Aucun fichier modifié")
    else:
        print(f"⚠️  Erreur lors de la vérification du statut: {stderr}")
    
    # Ajouter tous les fichiers modifiés et nouveaux
    print("\n2. Ajout des fichiers...")
    success, stdout, stderr = run_command("git add -A", cwd=project_dir)
    if success:
        print("✅ Fichiers ajoutés")
    else:
        print(f"❌ Erreur lors de l'ajout: {stderr}")
        sys.exit(1)
    
    # Vérifier s'il y a quelque chose à commiter
    success, stdout, stderr = run_command("git diff --cached --quiet", cwd=project_dir, check=False)
    if success:
        print("ℹ️  Aucun changement à commiter")
        # Vérifier s'il y a des commits à pousser
        success, stdout, stderr = run_command(f"git log {branch}..origin/{branch} --oneline", cwd=project_dir, check=False)
        if stdout.strip():
            print("📤 Il y a des commits locaux à pousser")
        else:
            print("✅ Tout est à jour, rien à pousser")
            return
    
    # Commit
    print("\n3. Création du commit...")
    commit_message = "Mise à jour du code - Génération d'abris vélos"
    success, stdout, stderr = run_command(
        f'git commit -m "{commit_message}"',
        cwd=project_dir
    )
    if success:
        print("✅ Commit créé")
        if stdout.strip():
            print(stdout)
    else:
        if "nothing to commit" in stderr.lower() or "rien à valider" in stderr.lower():
            print("ℹ️  Rien à commiter (déjà à jour)")
        else:
            print(f"❌ Erreur lors du commit: {stderr}")
            sys.exit(1)
    
    # Push
    print(f"\n4. Push vers GitHub (branche: {branch})...")
    success, stdout, stderr = run_command(f"git push origin {branch}", cwd=project_dir)
    if success:
        print("✅ Push réussi!")
        if stdout.strip():
            print(stdout)
        print("\n" + "=" * 60)
        print("✅ Modifications poussées sur GitHub")
        print("📦 Repository: https://github.com/bycommute/abris-velos-generator")
    else:
        print(f"❌ Erreur lors du push: {stderr}")
        print("\n💡 Suggestions:")
        print("   - Vérifiez votre connexion internet")
        print("   - Vérifiez vos credentials Git")
        print("   - Essayez: git pull origin " + branch)
        sys.exit(1)

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n⚠️  Opération annulée par l'utilisateur")
        sys.exit(1)
    except Exception as e:
        print(f"\n❌ Erreur inattendue: {e}")
        sys.exit(1)
