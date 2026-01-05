# 📊 Code Comparator

**Applicazione desktop per confrontare liste basate su colonne chiave (codici ISBN, inventari, ID, ecc.)**

![Python Version](https://img.shields.io/badge/python-3.9%2B-blue)
![PyQt6](https://img.shields.io/badge/PyQt6-6.4%2B-green)
![License](https://img.shields.io/badge/license-MIT-blue)

---

## ✨ Caratteristiche

- ✅ **Confronto rapido** tra due liste Excel/CSV
- 📊 **Visualizzazione immediata** di corrispondenze e mancanti
- 🎨 **Interfaccia moderna** con PyQt6
- 💾 **Export formattato** in Excel professionale
- ⚙️ **Preset configurabili** salvati automaticamente
- 🔄 **Gestione duplicati** con finestra dettagliata
- ↩️ **Ripristino righe** nella posizione originale (CTRL+Z)
- 📋 **Copy/Paste** da/verso Excel (CTRL+C)
- 🔍 **Normalizzazione automatica** delle chiavi
- 🔢 **Ordinamento colonne** con indicatori visibili (▲▼)
- 📝 **Rilevamento intestazioni** automatico con conferma

---

## 🚀 Installazione

### Prerequisiti
- **Python 3.9+**
- **pip**

### Installazione Rapida
```bash
# Clona il repository
git clone https://github.com/bibliotecamorante/code-comparator.git
cd code-comparator

# Installa dipendenze
pip install -r requirements.txt

# Avvia l'applicazione
python confronto_isbn.py
```

---

## 💡 Guida Rapida

1. **Copia** i dati da Excel/Google Sheets (CTRL+C)
2. **Incolla** in Lista 1 e Lista 2 (📋 pulsante o CTRL+V)
3. **Confronta** con il pulsante centrale 🔍
4. **Esporta** i risultati (💾 Corrispondenze o Mancanti)

---

## ⚙️ Preset e Configurazione

### Preset Predefiniti

**📚 ISBN (Predefinito)**
```
Colonna Chiave: ISBN
Colonne: ISBN, TITOLO, AUTORE, EDITORE, ANNO, PREZZO, NOTE
```

**📖 Stampa Registri (Biblioteca)**
```
Colonna Chiave: Inventario
Colonne: Inventario, Sezione, Collocazione, Specificazione, 
         Sequenza, Descrizione ISBD, Legami
```

### Creare Preset Personalizzati

1. Clicca **⚙️ Impostazioni**
2. Compila sezione **"➕ Crea Nuovo Preset"**:
   - Nome Preset
   - Nome colonna chiave
   - Numero colonna chiave (1-7)
   - Fino a 7 colonne totali
3. Clicca **"✓ Crea e Usa Nuovo Preset"**

I preset vengono salvati automaticamente in `presets.json` con backup automatico.

---

## 🛠️ Funzionalità Principali

### ↩️ Ripristino Righe Intelligente
- **CTRL+Z**: Ripristina righe eliminate nella posizione originale
- **Pulsante "↩️ Ripristina"**: Alternativa con interfaccia
- **Evidenziazione**: Righe ripristinate evidenziate in giallo

### 📋 Copia/Incolla Avanzato
- **CTRL+C**: Copia righe selezionate in formato Excel
- **Menu contestuale**: Tasto destro per azioni rapide
- **Incolla multipli**: Aggiungi dati senza sovrascrivere

### 🔍 Gestione Duplicati
- **Label cliccabile**: Mostra finestra con righe duplicate
- **Indicatore posizioni**: Vedi dove sono i duplicati nella tabella
- **Esclusi automaticamente**: Righe con chiave vuota
- **Conferma confronto**: Avviso prima di confrontare con duplicati

### 🔢 Ordinamento Tabelle
- **Click su header**: Ordina colonna con frecce ▲▼ visibili
- **Toggle asc/desc**: Click ripetuto inverte ordine
- **Reset automatico**: Si resetta incollando nuovi dati

### 📝 Rilevamento Intestazioni
- **Automatico**: Rileva se prima riga è intestazione
- **Conferma utente**: Dialog prima di applicare intestazione
- **Aggiornamento colonne**: Preset aggiornato con nomi reali

### ✖️ Eliminazione Righe
- **Selezione multipla**: CTRL+Click o Shift+Click
- **Validazione robusta**: Cerca indice in tutte le colonne
- **Conferma eliminazione**: Dialog di sicurezza

---

## 📊 Come Funziona

### Normalizzazione Chiavi
```
Originale:        978-88-06-12345-6
Normalizzato:     9788806123456

Rimuove: spazi, trattini, caratteri speciali
Converte: tutto in MAIUSCOLO
```

### Algoritmo Confronto
```
Per ogni riga in Lista 1:
  └─ Se chiave normalizzata esiste in Lista 2:
       └─ CORRISPONDENZA ✓
  └─ Altrimenti:
       └─ MANCANTE ✗
```

---

## 🎯 Casi d'Uso

- **📚 Biblioteche**: Confronta catalogo/inventario, verifica ISBN
- **🏪 Magazzini**: Ordini vs giacenze, codici prodotto
- **🏢 Uffici**: Anagrafiche clienti, codici fiscali
- **🎓 Didattica**: Liste studenti, presenze

---

## 📁 File Generati

```
code-comparator/
├── confronto_isbn.py       # Applicazione principale
├── presets.json            # Preset salvati automaticamente
├── presets.json.bak        # Backup automatico
└── Desktop/
    ├── ISBN_Corrispondenze.xlsx    # Export corrispondenze
    └── ISBN_Mancanti.xlsx          # Export mancanti
```

---

## 🐛 Risoluzione Problemi

### ❌ Modulo non trovato
```bash
pip install PyQt6 pandas openpyxl
```

### ❌ Colonna chiave non esiste
Verifica che il nome nel preset corrisponda esattamente (maiuscole/minuscole).

### ❌ Ripristino non funziona
Seleziona la tabella prima di premere CTRL+Z.

---

## 🔧 Requisiti

```
PyQt6 >= 6.4.0      # Interfaccia grafica
pandas >= 2.0.0     # Elaborazione dati
openpyxl >= 3.1.0   # Export Excel
```

**Sistema**: Windows 10+, macOS 10.14+, Linux (Ubuntu 20.04+)

---

## 📝 Changelog

### [1.1.0] - 2025-01-05
#### ✨ Nuovo
- Ripristino righe nella posizione originale
- Finestra dettagliata duplicati con indicatore posizioni
- Ordinamento colonne con frecce Unicode visibili
- Rilevamento intelligente intestazioni
- Validazione indici robusta (cerca in tutte le colonne)
- Reset automatico ordinamento su incolla
- Esclusione automatica righe vuote da conteggio duplicati

#### 🔧 Migliorato
- Gestione CTRL+Z più affidabile
- Performance con dataset grandi (>1000 righe)
- Backup automatico `presets.json`

### [1.0.0] - 2025-01-04
- Rilascio iniziale

---

## 📄 Licenza

**MIT License** - Vedi [LICENSE](LICENSE) per dettagli.

---

## 👤 Autore

**Biblioteca Morante**  
📧 bibliotecamorante@gmail.com  
🌐 [github.com/bibliotecamorante](https://github.com/bibliotecamorante)

---

**⭐ Se ti piace questo progetto, lascia una stella su GitHub!**

```
Made with ❤️ by Biblioteca Morante
```
