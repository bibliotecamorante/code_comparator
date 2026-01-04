# 📊 Code Comparator

**Applicazione desktop per confrontare liste basate su colonne chiave (codici ISBN, inventari, ID, ecc.)**

![Python Version](https://img.shields.io/badge/python-3.9%2B-blue)
![PyQt6](https://img.shields.io/badge/PyQt6-6.4%2B-green)
![License](https://img.shields.io/badge/license-MIT-blue)
![Status](https://img.shields.io/badge/status-beta-orange)

---

## 📸 Screenshot

![Schermata Principale](https://github.com/user-attachments/assets/f3aa4eb1-7238-4c94-a673-756c9b737e2b)


---

## ✨ Caratteristiche

- ✅ **Confronto rapido** tra due liste Excel/CSV
- 📊 **Visualizzazione immediata** di corrispondenze e mancanti
- 🎨 **Interfaccia moderna** con PyQt6
- 💾 **Export formattato** in Excel professionale
- ⚙️ **Preset configurabili** per diversi tipi di dati
- 🔄 **Gestione duplicati** con avvisi automatici
- ↩️ **Undo/Redo** per eliminazioni accidentali
- 🔍 **Normalizzazione automatica** delle chiavi di confronto
- 📋 **Copy/Paste** da Excel/Google Sheets
- 🎯 **Zero configurazione**: incolla e confronta!

---

## 🚀 Installazione Rapida

### Prerequisiti

- **Python 3.9 o superiore**
- **pip** (gestore pacchetti Python)

### Metodo 1: Installazione da Git (Consigliato)
```bash
# Clona il repository
git clone https://github.com/bibliotecamorante/code-comparator.git
cd code-comparator

# Installa le dipendenze
pip install -r requirements.txt

# Avvia l'applicazione
python code_comparator.py
```

### Metodo 2: Installazione come Pacchetto
```bash
# Clona e installa
git clone https://github.com/bibliotecamorante/code-comparator.git
cd code-comparator
pip install .

# Avvia da qualsiasi cartella
code-comparator
```

### Metodo 3: Solo Dipendenze
```bash
pip install PyQt6 pandas openpyxl
python code_comparator.py
```

---

## 💡 Guida Rapida

### 1️⃣ Preparazione Dati

Apri i tuoi file Excel/Google Sheets e **copia le colonne** che vuoi confrontare (CTRL+C).

**Esempio - Lista 1 (Worklist):**
```
ISBN            TITOLO                  AUTORE
9788806123456   Il nome della rosa      Umberto Eco
9788804567890   1984                    George Orwell
9788817012345   Il Gattopardo           Giuseppe Tomasi
```

**Esempio - Lista 2 (Magazzino):**
```
ISBN            TITOLO                  GIACENZA
9788806123456   Il nome della rosa      5
9788845678901   Cent'anni di solitudine 3
```

### 2️⃣ Incolla nelle Liste

1. Clicca **"📋 Incolla da Excel"** in Lista 1
2. Clicca **"📋 Incolla da Excel"** in Lista 2

### 3️⃣ Confronta

Clicca il pulsante **"🔍 CONFRONTA LISTE"**

### 4️⃣ Risultati

- **Tab "✓ Corrispondenze"**: ISBN presenti in entrambe le liste
- **Tab "✗ Mancanti"**: ISBN presenti solo in Lista 1

### 5️⃣ Esporta

- **"💾 ESPORTA CORRISPONDENZE"**: Salva i match sul Desktop
- **"💾 ESPORTA MANCANTI"**: Salva ciò che manca sul Desktop

---

## ⚙️ Preset Configurabili

### Preset Predefiniti

#### 📚 **ISBN (Predefinito)**
```
Colonna Chiave: ISBN
Colonne: ISBN, TITOLO, AUTORE, EDITORE, ANNO, PREZZO, NOTE
```

#### 📖 **Inventario (Biblioteca)**
```
Colonna Chiave: Inventario
Colonne: Inventario, Sezione, Collocazione, Specificazione, Sequenza, Descrizione ISBD, Legami
```

### Creare Preset Personalizzati

1. Clicca **"⚙️ Impostazioni"**
2. Compila la sezione **"➕ Crea Nuovo Preset"**:
   - **Nome Preset**: Es. "Magazzino"
   - **Nome colonna chiave**: Es. "Codice Prodotto"
   - **Numero colonna chiave**: Da 1 a 7
   - **Colonne aggiuntive**: Aggiungi fino a 7 colonne
3. Clicca **"✓ Crea e Usa Nuovo Preset"**

I preset vengono salvati automaticamente in `presets.json`.

---

## 🎯 Casi d'Uso

### 📚 Biblioteche
- Confrontare catalogo con inventario fisico
- Verificare ISBN mancanti in ordini
- Trovare duplicati tra collezioni

### 🏪 Magazzini
- Confrontare ordini con giacenze
- Verificare codici prodotto tra fornitori
- Controllare discrepanze inventario

### 🏢 Uffici
- Confrontare anagrafiche clienti
- Verificare codici fiscali/P.IVA
- Trovare dati mancanti tra database

### 🎓 Didattica
- Confrontare liste studenti
- Verificare presenze
- Trovare duplicati in anagrafiche

---

## 🛠️ Funzionalità Avanzate

### ↩️ Annulla Eliminazioni (CTRL+Z)
Se elimini righe per errore:
- Premi **CTRL+Z** o
- Clicca **"↩️ Ripristina righe"**

### 📋 Copia Righe (CTRL+C)
1. Seleziona righe nella tabella
2. Premi **CTRL+C**
3. Incolla in Excel

### 🗑️ Elimina Righe Selezionate
1. Seleziona righe
2. Tasto destro → **"✖ Elimina righe selezionate"**
3. Usa CTRL+Z per annullare

### 🔍 Ordina Colonne
Clicca sull'intestazione di qualsiasi colonna per ordinarla.

### ⚠️ Avvisi Duplicati
L'applicazione ti avvisa automaticamente se rileva:
- Chiavi duplicate in Lista 1
- Chiavi duplicate in Lista 2

---

## 📊 Come Funziona il Confronto

### Normalizzazione Chiavi
Prima del confronto, le chiavi vengono normalizzate:
```
Originale:        978-88-06-12345-6
Normalizzato:     9788806123456

Originale:        ISBN 978 88 06 12345 6
Normalizzato:     9788806123456
```

**Cosa viene rimosso:**
- Spazi
- Trattini
- Caratteri speciali
- Maiuscole/minuscole

### Algoritmo
```
Per ogni riga in Lista 1:
  └─ Se chiave normalizzata esiste in Lista 2:
       └─ CORRISPONDENZA ✓
  └─ Altrimenti:
       └─ MANCANTE ✗
```

---

## 📁 Struttura File
```
code-comparator/
├── code_comparator.py      # Applicazione principale
├── presets.json            # Preset salvati (generato automaticamente)
├── requirements.txt        # Dipendenze
├── setup.py               # Configurazione installazione
├── README.md              # Questa documentazione
├── LICENSE                # Licenza MIT
└── .gitignore            # File da ignorare su Git
```

---

## 🐛 Risoluzione Problemi

### ❌ Errore: "ModuleNotFoundError: No module named 'PyQt6'"
```bash
pip install PyQt6
```

### ❌ Errore: "No module named 'pandas'"
```bash
pip install pandas openpyxl
```

### ❌ L'app non si avvia su macOS
```bash
# Installa PyQt6 con supporto macOS
pip install --upgrade PyQt6
```

### ❌ "La colonna chiave non esiste"
Verifica che il nome della colonna nel preset corrisponda **esattamente** al nome della colonna nei tuoi dati (maiuscole/minuscole incluse).

### ❌ Il confronto è lento con file grandi (>10.000 righe)
L'app mostra automaticamente un cursore di attesa. Per file enormi (>50.000 righe), considera di suddividere i dati.

---

## 🔧 Requisiti Tecnici

### Dipendenze Python
```
PyQt6 >= 6.4.0     # Interfaccia grafica
pandas >= 2.0.0    # Elaborazione dati
openpyxl >= 3.1.0  # Export Excel formattato
```

### Requisiti Sistema
- **RAM**: Minimo 4GB (8GB consigliati per file grandi)
- **Spazio disco**: 100MB per installazione + spazio per output
- **Sistema operativo**: Windows 10/11, macOS 10.14+, Linux (Ubuntu 20.04+)

---

## 🤝 Contribuire

I contributi sono benvenuti! Per contribuire:

1. **Fork** il repository
2. **Crea** un branch per la tua feature (`git checkout -b feature/NuovaFunzione`)
3. **Commit** le modifiche (`git commit -m 'Aggiungi NuovaFunzione'`)
4. **Push** sul branch (`git push origin feature/NuovaFunzione`)
5. **Apri** una Pull Request

### Segnalare Bug
Apri una [Issue](https://github.com/bibliotecamorante/code-comparator/issues) descrivendo:
- Sistema operativo e versione Python
- Passi per riprodurre il bug
- Comportamento atteso vs effettivo
- Screenshot (se utili)

---

## 📝 Changelog

### [1.0.0] - 2025-01-04
#### ✨ Aggiunto
- Rilascio iniziale
- Confronto tra due liste basate su colonna chiave
- Preset configurabili (ISBN, Inventario)
- Export Excel formattato
- Undo/Redo per eliminazioni
- Gestione duplicati con avvisi
- Normalizzazione automatica chiavi
- Menu contestuale con azioni rapide

#### 🔧 Ottimizzazioni
- Cursore di attesa per dataset grandi (>1000 righe)
- Rendering tabelle ottimizzato con `itertuples()`
- Validazione robusta indici

---

## 📄 Licenza

Questo progetto è rilasciato sotto licenza **MIT License**.  
Vedi il file [LICENSE](LICENSE) per i dettagli completi.
```
Copyright (c) 2025 Biblioteca Morante

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
```

---

## 👤 Autore

**Biblioteca Morante**  
📧 Email: bibliotecamorante@gmail.com  
🌐 GitHub: [github.com/bibliotecamorante](https://github.com/bibliotecamorante)

---

## 🙏 Ringraziamenti

- **PyQt6** - Framework GUI
- **Pandas** - Elaborazione dati
- **OpenPyXL** - Manipolazione Excel
- Comunità open source per supporto e feedback

---

## 🔗 Link Utili

- 📖 [Documentazione PyQt6](https://www.riverbankcomputing.com/static/Docs/PyQt6/)
- 📊 [Pandas Documentation](https://pandas.pydata.org/docs/)
- 📝 [OpenPyXL Documentation](https://openpyxl.readthedocs.io/)

---

**⭐ Se ti piace questo progetto, lascia una stella su GitHub!**
```
Made with ❤️ by Biblioteca Morante
```
