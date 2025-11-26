# Générateur d'Abris Vélos - Interface Web

Application web pour générer toutes les variantes d'abris vélos à partir d'un fichier Excel de base.

## 🚀 Développement local

### Installation

```bash
npm install
```

### Lancer en développement

```bash
npm run dev
```

L'application sera accessible sur `http://localhost:5173`

### Build pour production

```bash
npm run build
```

## 📦 Déploiement sur Netlify

1. Connectez votre repository GitHub à Netlify
2. Configurez les paramètres de build :
   - Build command: `npm run build`
   - Publish directory: `dist`
3. Netlify détectera automatiquement le fichier `netlify.toml`

## 🔧 Configuration

### Netlify Functions

Les fonctions serverless sont dans `netlify/functions/`. Elles nécessitent :
- Python 3.x installé sur Netlify
- Les scripts Python du projet parent copiés dans la fonction

### Variables d'environnement

Aucune variable d'environnement requise pour le moment.

## 📝 Structure

```
site-web/
├── src/
│   ├── App.jsx          # Composant principal
│   ├── main.jsx         # Point d'entrée
│   └── index.css        # Styles
├── netlify/
│   └── functions/
│       └── generate.js  # Fonction serverless pour générer les fichiers
├── index.html
├── package.json
├── vite.config.js
└── netlify.toml         # Configuration Netlify
```

## ⚠️ Notes importantes

- Les scripts Python doivent être accessibles depuis la fonction Netlify
- Le fichier Excel de base doit être uploadé ou présent par défaut
- La génération peut prendre plusieurs minutes selon le nombre de variantes

