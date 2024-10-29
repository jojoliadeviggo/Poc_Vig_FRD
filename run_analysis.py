from lecture import PPTAnalyzer
import os
import traceback

# Token direct
hf_token = os.getenv("HF-TOKEN")

try:
    print("=== Début de l'analyse ===")
    print("Initialisation de l'analyseur...")
    analyzer = PPTAnalyzer(hf_token)

    # Utilisation de os.path pour construire le chemin de manière plus robuste
    base_dir = os.path.dirname(os.path.abspath(__file__))
    ppt_file = "LIAaudeladubuzz.pptx"
    ppt_path = os.path.join(base_dir, ppt_file)
    
    print(f"\nAnalyse du fichier: {ppt_path}")
    
    result = analyzer.analyze_ppt(ppt_path)
    
    print("\n=== Résultats de l'analyse ===")
    print(f"Titre: {result['title']}")
    print(f"\nRésumé: {result['summary']}")
    print(f"\nMots-clés: {', '.join(result['keywords'])}")
    print(f"\nNombre de mots: {result['text_length']}")
    
except Exception as e:
    print("\n=== Erreur détaillée ===")
    traceback.print_exc()
    print(f"\nMessage d'erreur: {str(e)}")
    
print("\n=== Fin du programme ===")