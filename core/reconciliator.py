# Logique métier
import pandas as pd
from typing import Dict, List, Tuple, Optional, Union
from dataclasses import dataclass
from pathlib import Path

@dataclass
class ReconciliationResult:
    file1_name: str
    file2_name: str
    file1_total: int
    file2_total: int
    matched_count: int
    file1_missing: pd.DataFrame
    file2_missing: pd.DataFrame
    file1_matched: pd.DataFrame
    file2_matched: pd.DataFrame
    # Totaux sur les matched uniquement
    file1_matched_amount_total: Optional[float]
    file2_matched_amount_total: Optional[float]
    amount_difference: Optional[float]
    # Détail des écarts
    amount_discrepancies: pd.DataFrame  # lignes avec écarts de montant
    discrepancy_count: int

class Reconciliator:
    def __init__(self):
        self.file1_data: Optional[pd.DataFrame] = None
        self.file2_data: Optional[pd.DataFrame] = None
        self.file1_name: str = ""
        self.file2_name: str = ""
        self.ref_col1: Optional[str] = None
        self.ref_col2: Optional[str] = None
        self.amount_col1: Optional[str] = None
        self.amount_col2: Optional[str] = None
    
    def load_file1(self, file_path: str) -> Tuple[bool, str]:
        try:
            self.file1_data = pd.read_excel(file_path)
            self.file1_name = Path(file_path).stem
            return True, f"✓ Fichier 1 chargé : {len(self.file1_data)} transactions"
        except Exception as e:
            return False, f"✗ Erreur chargement fichier 1 : {str(e)}"
    
    def load_file2(self, file_path: str) -> Tuple[bool, str]:
        try:
            self.file2_data = pd.read_excel(file_path)
            self.file2_name = Path(file_path).stem
            return True, f"✓ Fichier 2 chargé : {len(self.file2_data)} transactions"
        except Exception as e:
            return False, f"✗ Erreur chargement fichier 2 : {str(e)}"
    
    def get_file1_columns(self) -> List[str]:
        if self.file1_data is None:
            return []
        return list(self.file1_data.columns)
    
    def get_file2_columns(self) -> List[str]:
        if self.file2_data is None:
            return []
        return list(self.file2_data.columns)
    
    def set_columns(self, ref_col1: str, ref_col2: str, 
                   amount_col1: Optional[str] = None, 
                   amount_col2: Optional[str] = None):
        self.ref_col1 = ref_col1
        self.ref_col2 = ref_col2
        self.amount_col1 = amount_col1
        self.amount_col2 = amount_col2
    
    def reconcile(self) -> Tuple[bool, Union[ReconciliationResult, str]]:
        if self.file1_data is None or self.file2_data is None:
            return False, "Les deux fichiers doivent être chargés"
        
        if not self.ref_col1 or not self.ref_col2:
            return False, "Les colonnes de référence doivent être sélectionnées"
        
        try:
            # Nettoyer les références
            refs1 = self.file1_data[self.ref_col1].astype(str).str.strip()
            refs2 = self.file2_data[self.ref_col2].astype(str).str.strip()
            
            # Ensembles pour comparaison
            set1 = set(refs1)
            set2 = set(refs2)
            
            # Différences et intersection
            only_in_1 = set1 - set2
            only_in_2 = set2 - set1
            matched_refs = set1 & set2
            
            # Filtrer DataFrames
            file1_missing = self.file1_data[refs1.isin(only_in_1)].copy()
            file2_missing = self.file2_data[refs2.isin(only_in_2)].copy()
            file1_matched = self.file1_data[refs1.isin(matched_refs)].copy()
            file2_matched = self.file2_data[refs2.isin(matched_refs)].copy()
            
            # Calculs des montants sur les MATCHED uniquement
            amount1_total = None
            amount2_total = None
            amount_diff = None
            discrepancies_df = pd.DataFrame()
            discrepancy_count = 0
            
            if (self.amount_col1 and self.amount_col1 in self.file1_data.columns and
                self.amount_col2 and self.amount_col2 in self.file2_data.columns):
                
                # Convertir en numérique
                file1_matched[self.amount_col1] = pd.to_numeric(
                    file1_matched[self.amount_col1], errors='coerce'
                )
                file2_matched[self.amount_col2] = pd.to_numeric(
                    file2_matched[self.amount_col2], errors='coerce'
                )
                
                # Totaux sur les matched
                amount1_total = file1_matched[self.amount_col1].sum()
                amount2_total = file2_matched[self.amount_col2].sum()
                amount_diff = amount1_total - amount2_total
                
                # Créer un index sur les références pour merger
                file1_matched_indexed = file1_matched.set_index(
                    file1_matched[self.ref_col1].astype(str).str.strip()
                )
                file2_matched_indexed = file2_matched.set_index(
                    file2_matched[self.ref_col2].astype(str).str.strip()
                )
                
                # Identifier les écarts ligne par ligne
                discrepancies = []
                for ref in matched_refs:
                    row1 = file1_matched_indexed.loc[ref]
                    row2 = file2_matched_indexed.loc[ref]
                    
                    # Gérer les doublons (si plusieurs lignes avec même référence)
                    if isinstance(row1, pd.DataFrame):
                        row1 = row1.iloc[0]
                    if isinstance(row2, pd.DataFrame):
                        row2 = row2.iloc[0]
                    
                    amt1 = row1[self.amount_col1]
                    amt2 = row2[self.amount_col2]
                    
                    # Comparer avec une petite marge pour les floats
                    if pd.notna(amt1) and pd.notna(amt2) and abs(amt1 - amt2) > 0.01:
                        discrepancies.append({
                            'Reference': ref,
                            f'Montant_{self.file1_name}': amt1,
                            f'Montant_{self.file2_name}': amt2,
                            'Ecart': amt1 - amt2,
                            'Ecart_pct': ((amt1 - amt2) / amt2 * 100) if amt2 != 0 else None
                        })
                
                if discrepancies:
                    discrepancies_df = pd.DataFrame(discrepancies)
                    discrepancy_count = len(discrepancies)
            
            result = ReconciliationResult(
                file1_name=self.file1_name,
                file2_name=self.file2_name,
                file1_total=len(self.file1_data),
                file2_total=len(self.file2_data),
                matched_count=len(matched_refs),
                file1_missing=file1_missing,
                file2_missing=file2_missing,
                file1_matched=file1_matched,
                file2_matched=file2_matched,
                file1_matched_amount_total=amount1_total,
                file2_matched_amount_total=amount2_total,
                amount_difference=amount_diff,
                amount_discrepancies=discrepancies_df,
                discrepancy_count=discrepancy_count
            )
            
            return True, result
            
        except Exception as e:
            return False, f"Erreur lors de la réconciliation : {str(e)}"