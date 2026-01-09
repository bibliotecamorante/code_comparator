"""
Modelli di business logic e gestione dati
"""
import os
import json
import shutil
import logging
import pandas as pd

from config import CONFIG_FILE, DEFAULT_PRESETS, PROTECTED_PRESETS, MAX_COLUMNS


class ComparisonEngine:
    """Motore di confronto logico separato dalla UI"""
    
    @staticmethod
    def count_duplicates(df: pd.DataFrame, key_column: str = 'KEY_NORMALIZED') -> tuple[int, int]:
        """
        Conta duplicati in un DataFrame.
        
        Args:
            df: DataFrame con colonna chiave
            key_column: Nome colonna chiave (default: 'KEY_NORMALIZED')
        
        Returns:
            tuple: (unique_keys_duplicated, total_rows_duplicated)
        """
        if df.empty or key_column not in df.columns:
            return 0, 0
        
        valid_mask = df[key_column].notna() & (df[key_column].astype(str).str.strip() != "")
        df_valid = df[valid_mask]
        
        if df_valid.empty:
            return 0, 0
        
        duplicates_mask = df_valid[key_column].duplicated(keep=False)
        unique_keys = df_valid[duplicates_mask][key_column].nunique()
        total_rows = duplicates_mask.sum()
        
        return unique_keys, total_rows
    
    @staticmethod
    def compare(df1, df2, key_column='KEY_NORMALIZED'):
        """
        Confronta due DataFrame basandosi sulla colonna chiave normalizzata.
        
        Args:
            df1: DataFrame principale (worklist)
            df2: DataFrame di confronto
            key_column: Nome della colonna normalizzata (default: 'KEY_NORMALIZED')
        
        Returns:
            tuple: (matches_df, mismatches_df, stats_dict)
        """
        keys_df2 = set(df2[key_column].dropna())
        
        matches_mask = df1[key_column].isin(keys_df2)
        matches = df1[matches_mask].copy()
        mismatches = df1[~matches_mask].copy()
        
        duplicates_df1, _ = ComparisonEngine.count_duplicates(df1, key_column)
        duplicates_df2, _ = ComparisonEngine.count_duplicates(df2, key_column)
        
        stats = {
            'total': len(df1),
            'matches': len(matches),
            'mismatches': len(mismatches),
            'duplicates_df1': duplicates_df1,
            'duplicates_df2': duplicates_df2
        }
        
        return matches, mismatches, stats


class PresetManager:
    """Gestisce caricamento, salvataggio e validazione dei preset"""
    
    def __init__(self):
        self.config_file = CONFIG_FILE
        self.presets = DEFAULT_PRESETS.copy()
        self.protected_presets = PROTECTED_PRESETS.copy()
        self.default_preset = None
        self.preset_order = []
        
        self.load_presets()
    
    def load_presets(self):
        """Carica i preset e il preset predefinito da presets.json se esiste"""
        if not os.path.exists(self.config_file):
            logging.info("Nessun file presets.json trovato. Uso preset predefiniti.")
            return
        
        try:
            with open(self.config_file, "r", encoding="utf-8") as f:
                loaded_data = json.load(f)
            
            if isinstance(loaded_data, dict):
                if "presets" in loaded_data:
                    loaded_presets = loaded_data.get("presets", {})
                    self.default_preset = loaded_data.get("default_preset", None)
                    self.preset_order = loaded_data.get("preset_order", [])
                else:
                    loaded_presets = loaded_data
                    self.default_preset = None
                    self.preset_order = []
            else:
                loaded_presets = {}
                self.default_preset = None
                self.preset_order = []
            
            valid_presets = {}
            for preset_name, preset_data in loaded_presets.items():
                if self._validate_preset(preset_name, preset_data):
                    valid_presets[preset_name] = preset_data
                else:
                    logging.warning(f"Preset '{preset_name}' non valido, ignorato.")
            
            self.presets.update(valid_presets)
            
            if self.default_preset and self.default_preset not in self.presets:
                logging.warning(f"Preset predefinito '{self.default_preset}' non trovato. Resettato.")
                self.default_preset = None
            
            logging.info(f"Caricati {len(valid_presets)} preset da: {self.config_file}")
            
        except json.JSONDecodeError as e:
            logging.exception(f"Errore parsing JSON: {e}")
            self._restore_backup()
        except Exception as e:
            logging.exception(f"Errore nel caricamento dei preset: {e}")
    
    def save_presets(self):
        """Salva i preset attuali e il preset predefinito in presets.json"""
        try:
            if os.path.exists(self.config_file):
                backup_file = self.config_file + ".bak"
                try:
                    shutil.copy2(self.config_file, backup_file)
                except Exception as e:
                    logging.warning(f"Impossibile creare backup: {e}")
            
            data_to_save = {
                "default_preset": self.default_preset,
                "preset_order": self.preset_order,
                "presets": self.presets
            }
            
            with open(self.config_file, "w", encoding="utf-8") as f:
                json.dump(data_to_save, f, indent=4, ensure_ascii=False)
            
            logging.info(f"Preset salvati in: {self.config_file}")
            
        except PermissionError as e:
            logging.exception(f"PermissionError salvando preset: {e}")
            raise
        except Exception as e:
            logging.exception(f"Errore nel salvataggio dei preset: {e}")
            raise
    
    def _validate_preset(self, preset_name, preset_data):
        """Valida la struttura di un preset"""
        if not isinstance(preset_data, dict):
            return False
        
        if "key_column" not in preset_data or "columns" not in preset_data:
            return False
        
        key_column = preset_data["key_column"]
        columns = preset_data["columns"]
        
        if not preset_name or not key_column:
            return False
        if len(columns) < 1 or len(columns) > MAX_COLUMNS:
            return False
        if key_column not in columns:
            return False
        
        return True
    
    def _restore_backup(self):
        """Tenta di ripristinare il backup in caso di file corrotto"""
        backup_file = self.config_file + ".bak"
        
        if os.path.exists(backup_file):
            try:
                shutil.copy2(backup_file, self.config_file)
                logging.info(f"Backup ripristinato da: {backup_file}")
                self.load_presets()
            except Exception as e:
                logging.exception(f"Impossibile ripristinare il backup: {e}")
    
    @staticmethod
    def validate_preset_data(preset_name, key_column, columns):
        """Validazione unificata dei dati preset"""
        if not preset_name: 
            return False, "Il nome del preset è obbligatorio"
        if not key_column: 
            return False, "Seleziona una colonna come chiave (clicca sul radio button)"
        if len(columns) < 1: 
            return False, "Definisci almeno una colonna"
        if len(columns) > MAX_COLUMNS: 
            return False, f"Massimo {MAX_COLUMNS} colonne consentite"
        if key_column not in columns: 
            return False, f"La colonna chiave '{key_column}' deve essere presente nell'elenco"
        return True, ""