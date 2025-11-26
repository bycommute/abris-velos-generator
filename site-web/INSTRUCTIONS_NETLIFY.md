# Instructions pour déployer sur Netlify

## Problème actuel
Le site affiche une erreur 404. Cela peut être dû à :
1. Le build n'a pas été fait correctement
2. Le dossier de publication n'est pas correctement configuré
3. Le compte Netlify utilisé n'est pas le bon

## Solution

### 1. Vérifier la configuration Netlify

Dans l'interface Netlify (https://app.netlify.com) :

1. Allez sur votre site (chic-phoenix-9fcdd0)
2. Allez dans **Site settings** → **Build & deploy**
3. Vérifiez les paramètres suivants :

**Build settings:**
- **Base directory:** `site-web` (ou laissez vide si le repo est directement dans site-web)
- **Build command:** `npm install && npm run build`
- **Publish directory:** `dist`

### 2. Rebuild manuel

Dans l'interface Netlify :
1. Allez dans **Deploys**
2. Cliquez sur **Trigger deploy** → **Clear cache and deploy site**

### 3. Vérifier le fichier netlify.toml

Le fichier `netlify.toml` est maintenant corrigé et devrait être :
- Build command: `npm run build`
- Publish directory: `dist`
- Redirects configurés pour SPA

### 4. Si le problème persiste

Vérifiez dans les logs de build sur Netlify s'il y a des erreurs lors du build.

## Connexion Netlify CLI (optionnel)

Si vous voulez utiliser le CLI :

```bash
cd site-web
netlify login
netlify link  # Lier au site existant
netlify deploy --prod --dir=dist
```

