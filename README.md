# Générateur d'Abris Vélos Domino

Système de génération automatique de fichiers Excel pour calculer les prix des abris vélos Domino.

## Structure

```
.
├── fichier de base/
│   └── nepastoucher.xlsx          # Fichier Excel source (jamais modifié)
├── résultats/                      # Tous les fichiers Excel générés
│   ├── abris_ouverts/             # 40 fichiers pour abris ouverts
│   └── abris_fermes/              # 40 fichiers pour abris fermés
├── generate_abris_ouverts.py      # Script pour générer les abris ouverts
├── generate_abris_fermes.py       # Script pour générer les abris fermés
└── read_results.py                # Script pour lire les prix calculés
```

## Utilisation

### 1. Générer les abris ouverts

```bash
python generate_abris_ouverts.py
```

Génère 40 fichiers Excel dans `résultats/abris_ouverts/` :
- 5 largeurs (2.03, 2.53, 4.06, 5.06, 6.09m)
- 2 variantes (normal, bosqué)
- 2 traitements (Galvanized, Powder coat)
- 2 versions (Standard, PLUS)

**Configuration des abris ouverts :**
- Murs : haut, droite, gauche (pas en bas)
- Pas de portes
- Roof edge trim : No
- Chemical anchors : Yes

### 2. Générer les abris fermés

```bash
python generate_abris_fermes.py
```

Génère 40 fichiers Excel dans `résultats/abris_fermes/` :
- 5 largeurs (2.03, 2.53, 4.06, 5.06, 6.09m)
- 2 variantes (normal, bosqué)
- 2 traitements (Galvanized, Powder coat)
- 2 versions (Standard, PLUS)

**Configuration des abris fermés :**
- Murs : partout (haut, droite, bas, gauche)
- Portes : Single swing gate (2m, 1 porte)
- Gate hardware kit : Euro cylinder lock
- Roof edge trim : No
- Chemical anchors : Yes

### 3. Lire les résultats

```bash
python read_results.py
```

Lit tous les prix depuis les fichiers Excel dans `résultats/` et génère :
- `résultats/tous_les_resultats.json` : Tous les résultats en JSON
- Affichage dans le terminal avec un résumé

## Variantes

### Variante "normal"
- Wall material : Thermowood
- Remove cladding : Yes

### Variante "bosqué"
- Wall material : Thermowood
- Remove cladding : No

## Notes importantes

1. **Le fichier source `nepastoucher.xlsx` ne doit JAMAIS être modifié directement**
2. Tous les fichiers générés sont des copies du fichier source
3. Après génération, vous devez :
   - Ouvrir chaque fichier dans Excel
   - Appuyer sur F9 pour recalculer
   - Fermer Excel
   - Utiliser `read_results.py` pour lire les prix

## Dépendances

- Python 3
- openpyxl

```bash
pip install openpyxl
```

