```markdown
# Reconciliation Tool 211101

Outil desktop pour automatiser la réconciliation de transactions entre deux fichiers Excel/CSV. 
Développé en Python avec PySide6 pour l'interface graphique et pandas pour le traitement des données.

---

## Fonctionnalités

- **Générique** : Fonctionne avec n'importe quels fichiers Excel/CSV (banques, stocks, clients, paiements...)
- **Drag & Drop** : Glissez-déposez vos fichiers ou parcourez pour les sélectionner
- **Clés de comparaison multiples** : Combinez plusieurs colonnes pour une réconciliation précise (ex: ID + Date + Montant)
- **Mode Strict 1:1** : Appariement ligne à ligne - les doublons non symétriques vont en missing
- **Comparaison des montants** : Détecte les écarts de montant sur les transactions correspondantes
- **Gestion des valeurs vides** : Les lignes avec clés de comparaison incomplètes sont signalées en missing
- **Export Excel multi-feuilles** : Génère un fichier complet avec tous les résultats
- **Threading** : Interface réactive pendant le traitement (pas de blocage)
- **Offline** : Fonctionne sans connexion internet
- **Léger** : Aucune limite de taille de fichier (dépend de la RAM)

---

Architecture du Projet

```
reconciliation_app/
├── core/
│   └── reconciliator.py      # Logique métier de réconciliation
├── gui/
│   └── main_window.py        # Interface PySide6 (DropArea, tableaux, etc.)
├── utils/
│   └── excel_handler.py      # Gestion des exports Excel
├── assets/                   # Icônes, images, ressources
├── main.py                   # Point d'entrée de l'application
├── requirements.txt          # Dépendances Python
└── README.md                 # Ce fichier
```

---

## Modes de Réconciliation

### Mode Strict 1:1 (par défaut)
Chaque ligne est appariée individuellement avec une seule ligne de l'autre fichier.

| Fichier 1 | | Fichier 2 | | Résultat |
|-----------|---|-----------|---|----------|
| REF001 | 250€ | REF001 | 500€ | 1 ligne matchée |
| REF001 | 250€ | | | 1 ligne en F1_missing |

**Règle** : `min(count_F1, count_F2)` lignes matchées, le surplus va en missing.

## Gestion des doublons et valeurs vides

### Doublons asymétriques
| Cas | Fichier 1 | Fichier 2 | Mode Strict |
|-----|-----------|-----------|-------------|
| 2×REF001 vs 1×REF001 | 2 lignes | 1 ligne | 1 match + 1 missing |
| 3×REF002 vs 1×REF002 | 3 lignes | 1 ligne | 1 match + 2 missing |

### Valeurs vides dans les clés de comparaison
Les lignes avec une ou plusieurs valeurs vides dans les colonnes de comparaison sont :
- Assignées une clé unique temporaire (`__EMPTY_{index}__`)
- **Toujours placées en missing** de leur fichier respectif
- **Jamais matchées** avec l'autre fichier

Cela garantit qu'aucune ligne ne "disparaît" silencieusement.

---

## Structure des fichiers générés

L'export Excel contient 6 feuilles détaillées :

| Feuille | Description | Usage |
|---------|-------------|-------|
| `Fichier1_missing` | Transactions présentes dans Fichier 1 mais absentes de Fichier 2 | À traiter/rechercher |
| `Fichier2_missing` | Transactions présentes dans Fichier 2 mais absentes de Fichier 1 | À traiter/rechercher |
| `Fichier1_matched` | Transactions de Fichier 1 ayant une correspondance 1:1 | Vérification |
| `Fichier2_matched` | Transactions de Fichier 2 ayant une correspondance 1:1 | Vérification |
| `Ecarts_montant` | Détail des transactions avec écarts de montant (> 0.01€) | Investigation |
| `Summary` | Statistiques globales, taux de correspondance, totaux | Reporting |

**Colonnes dans Ecarts_montant :**
- `Reference` : La clé composite concernée
- `Montant_[Fichier1]` vs `Montant_[Fichier2]` : Valeurs comparées
- `Ecart` : Différence absolue
- `Ecart_pct` : Pourcentage d'écart

---

## Installation et utilisation

### Prérequis

- Python 3.9 ou supérieur
- Windows 10/11 (Linux/Mac possible avec adaptation des chemins)

### 1. Cloner le repository

```bash
git clone https://github.com/mapaycard/reconciliation-tool.git
cd reconciliation_app
```

### 2. Créer l'environnement virtuel

```bash
python -m venv env
```

### 3. Activer l'environnement virtuel

**Windows (CMD) :**
```cmd
env\Scripts\activate
```

**Windows (PowerShell) :**
```powershell
.\env\Scripts\activate
```

**Linux/Mac :**
```bash
source env/bin/activate
```

### 4. Installer les dépendances

```bash
pip install -r requirements.txt
```

**requirements.txt :**
```
pandas>=1.5.0
openpyxl>=3.0.0
PySide6>=6.4.0
```

---

## Créer l'exécutable (.exe)

### Méthode recommandée (PyInstaller)

Assurez-vous que l'environnement virtuel est activé, puis :

```powershell
# Nettoyer les anciens builds (PowerShell)
Remove-Item -Recurse -Force build, dist -ErrorAction SilentlyContinue

# Ou avec CMD
rmdir /s /q build dist

# Générer l'exécutable
pyinstaller --onefile --windowed --clean --noconfirm --name "Reconciliation_Tool" main.py
```

L'exécutable sera créé dans `dist/Reconciliation_Tool.exe`

### Options PyInstaller

| Option | Description |
|--------|-------------|
| `--onefile` | Crée un seul fichier .exe autonome |
| `--windowed` | Mode fenêtré (pas de console) |
| `--clean` | Nettoie le cache avant build |
| `--noconfirm` | Écrase les fichiers existants sans confirmation |
| `--name` | Nom de l'exécutable généré |

---

## Guide d'utilisation

### 1. Lancer l'application

Double-cliquez sur `Reconciliation_Tool.exe`

### 2. Charger les fichiers

- **Fichier Source 1** : Glissez-déposez ou cliquez pour sélectionner
- **Fichier Source 2** : Glissez-déposez ou cliquez pour sélectionner

Formats supportés : `.xlsx`, `.xls`, `.csv` (séparateur auto-détecté)

### 3. Configurer les clés de comparaison

Cliquez sur **"➕ Ajouter une colonne de comparaison"** pour définir les paires de colonnes :

```
Étape 1 : Sélectionnez la colonne dans Fichier 1 (ex: "Référence")
Étape 2 : Sélectionnez la colonne correspondante dans Fichier 2 (ex: "Ref_Paiement")
Étape 3 : Répétez pour ajouter d'autres critères (Date, Client, etc.)
```

Vous pouvez supprimer une clé avec le bouton 🗑️ à tout moment.

### 4. Configurer les montants (Optionnel)

| Paramètre | Description |
|-----------|-------------|
| **Fichier 1 - Montant** | Colonne contenant les montants du fichier 1 |
| **Fichier 2 - Montant** | Colonne contenant les montants du fichier 2 |

> Si non défini, la réconciliation se fait uniquement sur les clés (présence/absence).


### 5. Lancer la réconciliation

Cliquez sur **"🚀 Lancer la Réconciliation"**

L'opération s'exécute en arrière-plan. Une barre de statut indique la progression.

### 6. Analyser les résultats

**Panneau de statistiques :**
- Nombre total de transactions par fichier
- Nombre de correspondances (lignes matchées 1:1)
- Taux de correspondance global
- Totaux des montants (matched uniquement)
- Écart total et nombre de lignes avec écart

**Tableaux :**
- Transactions manquantes dans chaque fichier
- Coloration alternée pour faciliter la lecture
- Redimensionnement automatique des colonnes

### 8. Exporter les résultats

Cliquez sur **"💾 Exporter vers Excel"** pour générer le fichier complet avec les 6 feuilles.

### 9. Réinitialiser

Cliquez sur **"🔄 Réinitialisation"** pour tout effacer et recommencer avec de nouveaux fichiers.

---

## Cas d'usage typiques

| Scénario | Configuration recommandée |
|----------|--------------------------|
| **Réconciliation bancaire** | Clé : `Numéro_transaction` + `Date`<br>Montant : `Montant` |
| **Paiements échelonnés** | Clé : `ID_Client` + `Référence_Facture`<br>Montant : `Montant`<br>Mode : Agrégation |
| **Contrôle de caisse** | Clé : `Numéro_bon` + `Date` + `Caisse`<br>Montant : `Total_TTC` |
| **Suivi des remboursements** | Clé : `Numéro_dossier` + `Date_remboursement`<br>Montant : `Montant_remboursé` |

---

## Dépannage 071104

| Problème | Solution |
|----------|----------|
| "Format non valide" | Vérifiez que le fichier est bien .xlsx, .xls ou .csv |
| Taux de correspondance faible | Vérifiez le format des données (espaces, casse) ou ajoutez des clés |
| Lignes avec valeurs vides | Elles apparaissent automatiquement dans la feuille missing |
| Écarts de montant inexpliqués | Vérifiez que les colonnes de montant sont correctement sélectionnées |
| Lenteur sur gros fichiers | Normal pour >100k lignes. L'opération est non-bloquante. |
| Caractères spéciaux illisibles | Sauvegardez vos fichiers en UTF-8 avant import |

---