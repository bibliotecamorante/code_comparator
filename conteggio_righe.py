import os

def conta_righe_python():
    # Ottiene automaticamente il percorso della cartella dello script
    percorso_base = os.path.dirname(os.path.abspath(__file__))
    
    # Nome del file corrente da escludere
    nome_script_corrente = os.path.basename(__file__)
    
    totale_righe = 0
    file_processati = 0
    
    # Attraversa ricorsivamente tutte le cartelle e sottocartelle
    for root, dirs, files in os.walk(percorso_base):
        for file in files:
            # Controlla se è un file Python ed esclude lo script corrente dal conteggio
            if file.endswith('.py') and file != nome_script_corrente:
                percorso_completo = os.path.join(root, file)
                try:
                    with open(percorso_completo, 'r', encoding='utf-8') as f:
                        righe = len(f.readlines())
                        totale_righe += righe
                        file_processati += 1
                        print(f"{percorso_completo}: {righe} righe")
                except Exception as e:
                    print(f"Errore leggendo {percorso_completo}: {e}")
    
    print(f"\n{'='*50}")
    print(f"Totale file Python: {file_processati}")
    print(f"Totale righe: {totale_righe}")
    print(f"{'='*50}")
    
    return totale_righe

# Esegui il conteggio
conta_righe_python()