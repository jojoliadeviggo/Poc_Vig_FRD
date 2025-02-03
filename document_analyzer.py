from ppt_analysis import PPTAnalyzer
from pdf_analyzer import PDFAnalyzer
import os

class DocumentAnalyzer:
    def __init__(self):
        self.ppt_analyzer = PPTAnalyzer()
        self.pdf_analyzer = PDFAnalyzer()
        
    def analyze_document(self, file_path):
        """Analyze a document based on its extension"""
        # Vérifier si le fichier existe
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"Le fichier {file_path} n'existe pas")
            
        # Obtenir l'extension en minuscules
        extension = os.path.splitext(file_path)[1].lower()
        
        try:
            if extension in ['.ppt', '.pptx']:
                print("\nDétection d'un fichier PowerPoint")
                return self.ppt_analyzer.analyze(file_path)
                
            elif extension == '.pdf':
                print("\nDétection d'un fichier PDF")
                return self.pdf_analyzer.analyze(file_path)
                
            else:
                raise ValueError(f"Format de fichier non supporté: {extension}")
                
        except Exception as e:
            print(f"Erreur lors de l'analyse: {str(e)}")
            raise

# Exemple d'utilisation
if __name__ == "__main__":
    analyzer = DocumentAnalyzer()
    
    # Demander le chemin du fichier à l'utilisateur
    file_path = input("Entrez le chemin du fichier à analyser : ")


    # Exemple avec un fichier
    try:
        results = analyzer.analyze_document(file_path)
        print("\nRésultats de l'analyse:")
        print(f"Titre: {results['title']}")
        print(f"Nombre de mots: {results['text_length']}")
        print(f"Mots-clés: {', '.join(results['keywords'])}")
        print(f"Résumé: {results['summary']}")
        print(f"Table des matières: {results['table_of_contents']}")
    except Exception as e:
        print(f"Erreur: {str(e)}")