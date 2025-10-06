# Convoi Wizard

Convoi Wizard est une application de bureau Tkinter qui permet de consigner un journal de contacts relationnels et d'importer des analyses générées par un modèle de langage. Ce dépôt propose une structure de projet Python moderne et prête pour l'automatisation des tâches.

## Structure du projet

```
.
├── pyproject.toml       # Métadonnées du projet et entrée console `convoi-wizard`
├── requirements.txt     # Dépendances de développement minimales
├── tasks.py             # Tâches Invoke pour automatiser les actions courantes
└── src/
    └── convoi_wizard/
        ├── __init__.py
        ├── __main__.py  # Permet `python -m convoi_wizard`
        └── app.py       # Interface Tkinter principale
```

## Structure d'un workspace

Depuis la refonte du système de sauvegarde, l’application travaille avec un dossier (workspace) plutôt qu’un unique fichier JSON.
Chaque workspace contient la hiérarchie suivante :

```
<workspace>/
├── metadata.json               # Métadonnées globales + index des personnes/contacts
└── people/
    └── <person_id>/
        ├── person.json         # Informations de la personne (notes, sliders, contacts sans journal brut)
        └── contacts/
            └── <contact_id>/
                ├── journal.txt      # Journal libre (inclut la transcription injectée)
                ├── transcript.txt   # Transcription brute générée par Vibe (si audio)
                └── <fichiers audio> # Éventuels fichiers audio copiés dans le workspace
```

`metadata.json` ne contient que des métadonnées : titre, propriétaire, date de mise à jour et références RELATIVES vers chaque
`person.json` ainsi que vers les transcriptions (et éventuels fichiers audio si vous en ajoutez). Les journaux libres restent
donc stockés dans `journal.txt` au sein de chaque contact, ce qui permet de conserver des fichiers volumineux séparés du fichier
de métadonnées principal.

Lors d’une création de workspace, sélectionnez un dossier vide (ou acceptez d’écraser son contenu). L’application générera la
structure ci-dessus et y écrira automatiquement les données à chaque sauvegarde ou autosauvegarde.

## Prérequis

* Python 3.10 ou supérieur
* Pip et un environnement virtuel (recommandé)

## Installation rapide

```bash
python -m venv .venv
source .venv/bin/activate  # Sur Windows: .venv\\Scripts\\activate
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
pip install -e .
```

> Astuce: `invoke install --dev` installera automatiquement le projet en mode editable avec les dépendances de développement.

## Lancement de l'application

```bash
invoke run
```

Cette commande lance `python -m convoi_wizard` grâce aux tâches définies dans `tasks.py`.

## Configuration via fichier `.env`

L'application charge automatiquement un fichier `.env` (via [python-dotenv](https://github.com/theskumar/python-dotenv)) au démarrage.
Copiez le fichier `.env.example` fourni à la racine du dépôt, renommez-le en `.env` et ajustez les valeurs selon votre environnement :

* `HMT_DEFAULT_WORKSPACE` : chemin vers un workspace Convoi Wizard à ouvrir automatiquement (alias accepté : `CONVOI_DEFAULT_WORKSPACE`).
  Si le dossier n'existe pas ou que le chargement échoue, l'application revient sur le flux interactif habituel pour sélectionner/créer un workspace.
* `VIBE_CLI`, `VIBE_MODEL_PATH`, `VIBE_LANGUAGE`, `VIBE_THREADS`, `VIBE_TEMPERATURE` : paramètres optionnels pour piloter Vibe sans
  exporter les variables d'environnement dans votre shell.

> Astuce : si vous souhaitez utiliser un autre nom de fichier (par exemple `~/secrets/convoi.env`), définissez la variable d'environnement
> `CONVOI_ENV_FILE` avant de lancer l'application ; Convoi Wizard chargera ce fichier au lieu de `.env`.

## Transcription audio avec Vibe (mode local)

L’application peut lancer automatiquement la transcription d’un enregistrement audio en s’appuyant sur [Vibe](https://github.com/thewh1teagle/vibe), qui expose un binaire CLI (`vibe --help`) et un serveur HTTP optionnel (`vibe --server`) pour un usage local.【e52688†L32-L54】【1cc06d†L47-L170】

### Installation de Vibe

1. Téléchargez la dernière version de Vibe depuis sa page GitHub et extrayez le binaire (`vibe`, `vibe.exe`, etc.).
2. Récupérez un modèle Whisper au format `.bin` (voir la documentation « Models » du projet Vibe) et placez-le dans un dossier accessible.
3. Définissez les variables d’environnement suivantes avant de lancer Convoi Wizard :
   * `VIBE_CLI` : chemin absolu vers le binaire Vibe si celui-ci n’est pas dans votre `PATH`.
   * `VIBE_MODEL_PATH` : chemin absolu (ou relatif au workspace) vers le modèle Whisper `.bin` à utiliser.
   * Optionnel : `VIBE_LANGUAGE` (par défaut `french`), `VIBE_THREADS` (nombre de threads), `VIBE_TEMPERATURE` (valeur flottante à transmettre à Vibe).

> Astuce : sur macOS/Linux, vous pouvez exporter ces variables dans votre shell (`export VIBE_MODEL_PATH=/chemin/ggml-medium.bin`). Sous Windows, utilisez la boîte de dialogue « Variables d’environnement » ou un script `setx`.

### Fonctionnement dans Convoi Wizard

* Lorsqu’un fichier audio est attaché à un contact, il est copié dans le workspace puis un worker démarre Vibe en arrière-plan.
* La progression (0 – 100 %) est affichée dans un encart dédié avec un bouton « Annuler » qui interrompt proprement la transcription.
* La transcription générée est écrite dans `transcript.txt` aux côtés du fichier audio, et le texte est injecté dans le champ `journal` (avec conservation du journal libre éventuel).
* Le mode actuel est mono-locuteur : aucune diarisation n’est demandée et le modèle est initialisé avec la langue française par défaut.

Si Vibe n’est pas configuré (binaire introuvable ou modèle manquant), l’application affiche un avertissement et laisse le contact inchangé.

## Lint facultatif

Si [Ruff](https://github.com/astral-sh/ruff) est installé dans votre environnement, vous pouvez exécuter :

```bash
invoke lint
```

La tâche affiche un message si Ruff est absent.

## Tests

Aucun test automatisé n'est fourni pour l'instant. Vous pouvez ajouter vos propres tests sous `tests/` et les intégrer aux tâches Invoke selon vos besoins.
