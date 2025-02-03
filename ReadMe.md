# Analyseur de Documents (PPT & PDF)

Un outil Python pour analyser automatiquement des documents PowerPoint et PDF. L'analyseur extrait les informations clés, génère des résumés et identifie les mots-clés importants.

## Fonctionnalités

- Analyse de fichiers PowerPoint (.ppt, .pptx) et PDF
- Extraction automatique du titre et du texte
- Génération de résumés via Mistral AI
- Identification des mots-clés avec KeyBERT
- Extraction de la table des matières
- Interface unifiée pour tous les formats de documents

## Prérequis

### Python et environnement virtuel
```bash
python -m venv venv
source venv/bin/activate  # Linux/Mac
venv\Scripts\activate     # Windows


Dépendances python : 
pip install python-pptx
pip install keybert
pip install mistralai
pip install pytesseract
pip install pdf2image
pip install Pillow

Dépendances système

Tesseract OCR (pour l'analyse PDF)

Windows : Installer depuis le dépôt GitHub de Tesseract
Linux : sudo apt-get install tesseract-ocr
macOS : brew install tesseract


Poppler (pour la conversion PDF)

Windows : Télécharger et configurer dans le PATH
Linux : sudo apt-get install poppler-utils
macOS : brew install poppler



Configuration

Créer un fichier .env à la racine du projet
Ajouter vos clés API :
MISTRAL_API_KEY=votre_clé_mistral

Structure du projet :
.
├── document_analyzer.py    # Point d'entrée principal
├── pdf_analyzer.py        # Analyseur de PDF
├── ppt_analysis.py        # Analyseur de PowerPoint
└── .env                   # Configuration des clés API

Utilisation

1. Lancer l'analyseur :
python document_analyzer.py

2. Entrer le chemin complet du document à analyser (avec extension) :
Entrez le chemin du fichier à analyser : C:/chemin/vers/votre/document.pptx

3. L'analyseur affichera :


Titre du document
Nombre de mots
Mots-clés identifiés
Résumé généré
Table des matières (si disponible)

Sortie
Les résultats sont fournis sous forme de dictionnaire avec les champs suivants :
{
    "title": str,            # Titre du document
    "text_length": int,      # Nombre de mots
    "keywords": list,        # Liste des mots-clés
    "summary": str,          # Résumé généré
    "table_of_contents": str # Table des matières
}

Notes

Pour les PDF, assurez-vous que Tesseract est correctement installé et configuré
Pour les fichiers avec des caractères spéciaux dans le chemin, utilisez des slashs simples (/) ou des antislashs doubles (\)
Le titre est extrait de la première diapositive pour les PowerPoint

Limitation

La qualité de l'extraction PDF dépend de la qualité du document source
La génération de résumés est limitée par la taille du texte (2000 caractères pour PPT, 5000 pour PDF)