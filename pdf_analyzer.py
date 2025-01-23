import pytesseract
from pdf2image import convert_from_path
import os

from pdf2image.exceptions import PDFInfoNotInstalledError
import os
os.environ['PATH'] += os.pathsep + r'C:\Program Files\poppler\poppler-24.08.0\Library\bin'

from keybert import KeyBERT
from mistralai import Mistral

class PDFAnalyzer:
    def __init__(self):
            print("Initializing PDF Analyzer...")
            # Initialisation de KeyBERT
            self.keyword_model = KeyBERT(model="all-MiniLM-L6-v2")
            # Initialisation de Mistral
            self.mistral_client = Mistral(api_key=os.getenv("MISTRAL_API_KEY"))
            # Définir explicitement le chemin vers Tesseract
            # pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
            # Définir explicitement le chemin vers les données
            os.environ['TESSDATA_PREFIX'] = r'C:\Program Files\Tesseract-OCR\tessdata'

    def extract_text(self, pdf_path, language="fra"):
        """Extract text from PDF file using OCR"""
        try:
            # Obtenir et afficher le chemin absolu (comme dans PPTAnalyzer)
            abs_path = os.path.abspath(pdf_path)
            print(f"\nLe document analysé est situé ici : {abs_path}")

            # Convertir le PDF en images
            print("Converting PDF to images...")
            pages = convert_from_path(pdf_path)

            # Extraire le texte de chaque page
            text_content = []
            for i, page in enumerate(pages, start=1):
                print(f"Processing page {i}/{len(pages)}...")
                text = self.process_page(page, language)
                if text:
                    text_content.append(text)
                else:
                    print(f"Warning: No text extracted from page {i}")

            # Combiner le texte de toutes les pages
            full_text = "\n\n".join(text_content)
            
            if not full_text.strip():
                print("No text found in PDF")
                return ""

            print("\nExtraction completed successfully")
            return full_text

        except Exception as e:
            print(f"Error extracting text: {str(e)}")
            raise

    def process_page(self, page, language):
        """Helper function to process page with error handling"""
        custom_config = r'--oem 3 --psm 6 -c preserve_interword_spaces=1'
        try:
            # On demande explicitement à Tesseract de nous donner le texte en UTF-8
            raw_text = pytesseract.image_to_string(
                page,
                lang=language,
                config=custom_config,
                output_type=pytesseract.Output.STRING
            )
            
            if not raw_text:
                return ""
            
            # Au lieu de manipuler l'encodage, on nettoie simplement les caractères non désirés
            # tout en préservant les caractères accentués
            cleaned_text = ''
            for char in raw_text:
                # On garde tous les caractères alphanumériques et la ponctuation de base
                if (char.isalnum() or 
                    char.isspace() or 
                    char in '.,!?-:;\'\"()[]{}' or
                    ord(char) > 127):  # Ceci préserve les caractères accentués
                    cleaned_text += char
            
            return cleaned_text.strip()
                
        except Exception as e:
            print(f"Error processing page (detail): {repr(e)}")  # Utilisation de repr() pour plus de détails
            return ""

    def extract_keywords(self, text):
        """Extract keywords from text using KeyBERT"""
        print("\nExtracting keywords...")
        try:
            keywords = self.keyword_model.extract_keywords(
                text,
                keyphrase_ngram_range=(1, 1),  # Extrait les mots simples et les paires
                top_n=10,
                use_maxsum=True,
                nr_candidates=20
            )
            
            print("\nKeywords found:")
            for keyword, score in keywords:
                print(f"- {keyword}: {score:.3f}")
            
            return [kw for kw, _ in keywords]
            
        except Exception as e:
            print(f"Error extracting keywords: {str(e)}")
            return []

    def generate_summary(self, text, keywords):
        print("\nGenerating summary...")
        try:
            messages = [
                {
                    "role": "user",
                    "content": f"""Voici un texte extrait d'un document PDF.
                    Les mots-clés importants sont : {', '.join(keywords)}
                    
                    Texte : {text[:5000]}
                    
                    Tu dois générer un résumé concis et détaillé (5-10 phrases) qui :
                    1. Capture tous les points essentiels du document de manière approfondie 
                    2. Intègre naturellement les mots-clés identifiés
                    3. Est structuré de manière cohérente avec des transitions logiques
                    4. Conserve la complexité et les nuances du contenu d'origine"""
                }
            ]

            response = self.mistral_client.chat.complete(
                model="mistral-tiny",
                messages=messages,
            )
            
            summary = response.choices[0].message.content
            print(f"\nGenerated summary: {summary}")
            return summary
            
        except Exception as e:
            print(f"Error generating summary: {str(e)}")
            return ""


    def extract_table_of_contents(self, text):
        print("\nExtracting/generating table of contents...")
        try:
            messages = [
                {
                    "role": "user",
                    "content": f"""Analyse ce texte extrait d'un document PDF et fais l'une des trois choses suivantes :

                    1. Si tu trouves un sommaire explicite (avec des numéros de chapitres/sections), extrais-le et retourne-le tel quel.
                    2. Si tu ne trouves pas de sommaire explicite mais que le texte a une structure claire avec des sections distinctes, génère un sommaire qui reflète cette structure.
                    3. Si le texte n'a pas de structure claire ou de sections distinctes, retourne une chaîne vide.

                    Texte : {text[:5000]}

                    Important : Ne génère pas de sommaire artificiel si le texte n'a pas de structure claire."""
                }
            ]

            response = self.mistral_client.chat.complete(
                model="mistral-tiny",  # Utilisation du modèle tiny pour le POC
                messages=messages,
            )
            
            toc = response.choices[0].message.content
            print(f"\nTable of contents result: {toc}")
            return toc if toc and not toc.lower().startswith("le texte ne") else ""
            
        except Exception as e:
            print(f"Error extracting table of contents: {str(e)}")
            return ""

            response = self.mistral_client.chat.complete(
                model="mistral-tiny",
                messages=messages,
                max_tokens=1000,  # Augmente la longueur maximale de la réponse
                temperature=0.7,  # Légère augmentation de la créativité
            )
            
            summary = response.choices[0].message.content
            print(f"\nGenerated summary: {summary}")
            return summary
            
        except Exception as e:
            print(f"Error generating summary: {str(e)}")
            return ""

    def analyze(self, pdf_path):
            """Main analysis function"""
            try:
                # Obtenir et afficher le chemin absolu
                abs_path = os.path.abspath(pdf_path)
                print(f"\nLe document analysé est situé ici : {abs_path}")
                
                text = self.extract_text(pdf_path)
            
                if not text.strip():
                    return {
                        "title": "Sans titre",
                        "text_length": 0,
                        "keywords": [],
                        "summary": "",
                        "table_of_contents": "",
                        "extracted_text": ""
                    }

                # Extraire les mots-clés
                keywords = self.extract_keywords(text)
                # Générer le résumé
                summary = self.generate_summary(text, keywords)
                # Extraire le sommaire
                table_of_contents = self.extract_table_of_contents(text)

                return {
                    "title": title,
                    "text_length": len(text.split()),
                    "keywords": keywords,
                    "summary": summary,
                    "table_of_contents": table_of_contents,
                    # "extracted_text": text[:200] + "..."  # Preview des 200 premiers caractères
                }

            except Exception as e:
                print(f"Analysis failed: {str(e)}")
                raise

# Example usage
if __name__ == "__main__":
    analyzer = PDFAnalyzer()
    result = analyzer.analyze("Trends_IA.pdf")
    print("\nAnalysis results:", result)