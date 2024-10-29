from transformers import pipeline
import pptx
import os

class PPTAnalyzer:
    def __init__(self, hf_token):
        self.summarizer = pipeline(
            "summarization",
            model="facebook/bart-large-cnn",
            token=hf_token
        )
        
        self.classifier = pipeline(
            "zero-shot-classification",
            model="facebook/bart-large-mnli",
            token=hf_token
        )

    def clean_text(self, text):
        """Nettoie et prépare le texte pour l'analyse"""
        # Supprime les caractères spéciaux et les sauts de ligne multiples
        cleaned = ' '.join(text.split())
        # S'assure que le texte a une longueur minimale
        if len(cleaned) < 50:
            cleaned = cleaned + " " + cleaned
        return cleaned

    def extract_text(self, ppt_path):
        try:
            prs = pptx.Presentation(ppt_path)
            all_text = []
            
            for slide in prs.slides:
                slide_text = []
                for shape in slide.shapes:
                    if hasattr(shape, "text"):
                        text = shape.text.strip()
                        if text:
                            slide_text.append(text)
                
                if slide_text:
                    all_text.append(" ".join(slide_text))
            
            # Joindre tout le texte
            full_text = " ".join(all_text)
            print(f"Texte extrait (premiers 100 caractères): {full_text[:100]}...")
            return full_text
            
        except Exception as e:
            print(f"Erreur dans extract_text: {str(e)}")
            raise Exception(f"Erreur lors de l'extraction du texte : {str(e)}")

    def generate_summary(self, text):
        try:
            # Nettoyer et préparer le texte
            cleaned_text = self.clean_text(text)
            print(f"Longueur du texte nettoyé: {len(cleaned_text)} caractères")

            # Vérifier la longueur minimale
            if len(cleaned_text) < 50:
                return "Texte trop court pour générer un résumé."

            # Limiter la taille du texte pour BART
            max_length = 500
            if len(cleaned_text) > max_length:
                cleaned_text = cleaned_text[:max_length]

            print(f"Génération du résumé pour: {cleaned_text[:100]}...")
            
            # Générer le résumé avec des paramètres plus conservateurs
            summary = self.summarizer(
                cleaned_text,
                max_length=100,
                min_length=30,
                do_sample=False,
                truncation=True
            )
            
            if not summary:
                return "Impossible de générer un résumé."
                
            return summary[0]['summary_text']
            
        except Exception as e:
            print(f"Erreur détaillée dans generate_summary: {str(e)}")
            print(f"Texte problématique: {text[:200]}...")
            raise Exception(f"Erreur lors de la génération du résumé : {str(e)}")

    def extract_keywords(self, text):
        try:
            cleaned_text = self.clean_text(text)
            print(f"Extraction des mots-clés du texte de longueur: {len(cleaned_text)}")

            # Limiter la taille pour la classification
            if len(cleaned_text) > 500:
                cleaned_text = cleaned_text[:500]

            candidate_labels = [
                "technologie", "innovation", "business",
                "recherche", "développement", "stratégie",
                "marketing", "finance", "management"
            ]
            
            results = self.classifier(
                cleaned_text,
                candidate_labels=candidate_labels,
                multi_label=True,
                truncation=True
            )
            
            return [label for label, score in zip(results['labels'], results['scores']) if score > 0.3]
        except Exception as e:
            print(f"Erreur dans extract_keywords: {str(e)}")
            return []

    def analyze_ppt(self, ppt_path):
        try:
            print(f"Analyse du fichier: {ppt_path}")
            if not os.path.exists(ppt_path):
                raise FileNotFoundError(f"Le fichier {ppt_path} n'existe pas")
            
            # Extraire le texte
            text = self.extract_text(ppt_path)
            if not text.strip():
                return {
                    "title": "Sans titre",
                    "summary": "Aucun texte trouvé dans la présentation.",
                    "keywords": [],
                    "text_length": 0
                }
            
            # Récupérer le titre
            prs = pptx.Presentation(ppt_path)
            title = ""
            if prs.slides:
                first_slide = prs.slides[0]
                for shape in first_slide.shapes:
                    if hasattr(shape, "text"):
                        title = shape.text.strip()
                        if title:
                            break

            # Analyser le contenu
            summary = self.generate_summary(text)
            keywords = self.extract_keywords(text)
            
            return {
                "title": title or "Sans titre",
                "summary": summary,
                "keywords": keywords,
                "text_length": len(text.split())
            }
            
        except Exception as e:
            print(f"Erreur complète dans analyze_ppt: {str(e)}")
            raise Exception(f"Erreur lors de l'analyse du PPT : {str(e)}")