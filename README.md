```markdown
# Reconciliation Tool

Outil desktop léger pour automatiser la réconciliation de transactions entre deux fichiers Excel. 
Développé en Python avec PySide6 pour l'interface graphique et pandas pour le traitement des données.

![Python](https://img.shields.io/badge/Python-3.9+-blue.svg)
![License](https://img.shields.io/badge/License-MIT-green.svg)
![Platform](https://img.shields.io/badge/Platform-Windows-lightgrey.svg)

---

## Fonctionnalités

- **Générique** : Fonctionne avec n'importe quels fichiers Excel (banques, stocks, clients...)
- **Drag & Drop** : Glissez-déposez vos fichiers ou parcourez pour les sélectionner
- **Détection automatique** des colonnes
- **Comparaison par référence** : Identifie les transactions présentes dans un fichier mais pas l'autre
- **Comparaison des montants** : Détecte les écarts de montant sur les transactions correspondantes
- **Export Excel** : Génère 6 feuilles avec les résultats détaillés
- **Offline** : Fonctionne sans connexion internet
- **Léger** : Aucune limite de taille de fichier (dépend de la RAM)

---

## Gestion des doublons

Le programme gère les références en double de la manière suivante :

| Aspect | Comportement | Exemple |
|--------|--------------|---------|
| **Stats** | Compte les **références uniques** | 3 correspondances (REF001, REF002, REF003) |
| **Feuilles matched** | Contient **toutes les lignes** | 4 lignes si REF001 apparaît 2 fois |
| **Montants** | **Somme de toutes les lignes** | 500€ + 500€ = 1000€ pour REF001 |

### Cas pratique

**Fichier 1 :**
| référence | montant |
|-----------|---------|
| REF001 | 500€ |
| REF001 | 500€ | ← même référence, 2ème paiement |
| REF002 | 300€ |

**Fichier 2 :**
| référence | montant |
|-----------|---------|
| REF001 | 1000€ | ← somme des deux paiements |
| REF002 | 300€ |

**Résultat :**
- Stats : "2 correspondances" (REF001 et REF002)
- Feuille Fichier1_matched : 3 lignes (les 2 REF001 + REF002)
- Montant total Fichier 1 : 1300€ (500 + 500 + 300)
- Montant total Fichier 2 : 1300€ (1000 + 300)
- Écart : 0€

> **Note** : Ce comportement est adapté aux cas de paiements échelonnés où une référence peut apparaître plusieurs fois. Les montants sont toujours sommés correctement.

---

## Structure des fichiers générés

| Feuille | Description |
|---------|-------------|
| `Fichier1_missing` | Transactions présentes dans Fichier 1 mais absentes de Fichier 2 |
| `Fichier2_missing` | Transactions présentes dans Fichier 2 mais absentes de Fichier 1 |
| `Fichier1_matched` | Transactions de Fichier 1 ayant une correspondance dans Fichier 2 |
| `Fichier2_matched` | Transactions de Fichier 2 ayant une correspondance dans Fichier 1 |
| `Ecarts_montant` | Détail des transactions avec écarts de montant |
| `Summary` | Statistiques globales et totaux |

---

## Installation et utilisation

### Prérequis

- Python 3.9 ou supérieur
- Windows 10/11

### 1. Cloner le repository

```bash
git clone https://github.com/NabySidime/reconciliation-tool.git
cd reconciliation-tool
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
env\Scripts\Activate.ps1
```

**Linux/Mac :**
```bash
source env/bin/activate
```

### 4. Installer les dépendances

```bash
pip install -r requirements.txt
```

---

## Créer l'exécutable (.exe)

### Méthode recommandée (PyInstaller)

Assurez-vous que l'environnement virtuel est activé, puis :

```bash
# Nettoyer les anciens builds (Windows PowerShell)
Remove-Item -Recurse -Force build, dist -ErrorAction SilentlyContinue

# Ou avec CMD
rmdir /s /q build dist

# Générer l'exécutable
pyinstaller --onefile --windowed --clean --noconfirm --name "Reconciliation_Tool" main.py
```

L'exécutable sera créé dans le dossier `dist/Reconciliation_Tool.exe`

### Options PyInstaller

| Option | Description |
|--------|-------------|
| `--onefile` | Crée un seul fichier .exe autonome |
| `--windowed` | Mode fenêtré (pas de console) |
| `--clean` | Nettoie le cache avant build |
| `--noconfirm` | Écrase les fichiers existants sans confirmation |
| `--name` | Nom de l'exécutable généré |

---

## Comment utiliser

### 1. Lancer l'application

Double-cliquez sur `Reconciliation_Tool.exe` ou lancez :
```bash
python main.py
```

### 2. Charger les fichiers

- **Fichier Source 1** : Glissez-déposez ou cliquez pour sélectionner votre premier fichier Excel
- **Fichier Source 2** : Glissez-déposez ou cliquez pour sélectionner votre second fichier Excel

### 3. Configurer les colonnes

| Colonne | Description | Obligatoire |
|---------|-------------|-------------|
| **Colonne Référence** | Colonne contenant l'identifiant unique de transaction | Oui |
| **Colonne Montant** | Colonne contenant le montant (pour vérification) | Optionnel |

### 4. Lancer la réconciliation

Cliquez sur **"Lancer la Réconciliation"**

### 5. Analyser les résultats

Les résultats s'affichent dans l'interface :
- Statistiques globales (nombre de références uniques)
- Tableaux des transactions manquantes
- Totaux des montants (somme de toutes les lignes, y compris doublons)

### 6. Exporter

Cliquez sur **"Exporter vers Excel"** pour sauvegarder les résultats détaillés.

## Auteur

Créé par Naby Sidimé (https://github.com/NabySidime)
LinkedIn : https://www.linkedin.com/in/naby-sidimé-7688b0352
