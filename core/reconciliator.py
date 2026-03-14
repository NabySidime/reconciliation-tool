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
    file1_matched_amount_total: Optional[float]
    file2_matched_amount_total: Optional[float]
    amount_difference: Optional[float]
    amount_discrepancies: pd.DataFrame
    discrepancy_count: int
    aggregation_mode: bool
    comparison_keys: List[Tuple[str, str]]  # Liste des paires de colonnes comparées

class Reconciliator:
    def __init__(self):
        self.file1_data: Optional[pd.DataFrame] = None
        self.file2_data: Optional[pd.DataFrame] = None
        self.file1_name: str = ""
        self.file2_name: str = ""
        self.comparison_keys: List[Tuple[str, str]] = []  # [(col1_fichier1, col1_fichier2), ...]
        self.amount_col1: Optional[str] = None
        self.amount_col2: Optional[str] = None
        self.aggregation_mode: bool = False
    
    def load_file1(self, file_path: str) -> Tuple[bool, str]:
        try:
            file_path_lower = file_path.lower()
            if file_path_lower.endswith('.csv'):
                self.file1_data = pd.read_csv(file_path, sep=None, engine='python')
            else:
                self.file1_data = pd.read_excel(file_path)
            
            self.file1_name = Path(file_path).stem
            return True, f"✓ Fichier 1 chargé : {len(self.file1_data)} transactions"
        except Exception as e:
            return False, f"✗ Erreur chargement fichier 1 : {str(e)}"
    
    def load_file2(self, file_path: str) -> Tuple[bool, str]:
        try:
            file_path_lower = file_path.lower()
            if file_path_lower.endswith('.csv'):
                self.file2_data = pd.read_csv(file_path, sep=None, engine='python')
            else:
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
    
    def set_comparison_keys(self, keys: List[Tuple[str, str]], 
                           amount_col1: Optional[str] = None, 
                           amount_col2: Optional[str] = None,
                           aggregation_mode: bool = False):
        """Définit les clés de comparaison (plusieurs paires de colonnes)"""
        self.comparison_keys = keys
        self.amount_col1 = amount_col1
        self.amount_col2 = amount_col2
        self.aggregation_mode = aggregation_mode
    
    def _clean_reference(self, series: pd.Series) -> pd.Series:
        """Nettoie les références : enlève espaces, convertit en string sans décimales"""
        result = series.astype(str)
        result = result.str.strip()
        # Enlever .0 à la fin UNIQUEMENT pour les nombres entiers
        result = result.str.replace(r'^(\d+)\.0$', r'\1', regex=True)
        return result
    
    def _create_composite_key(self, df: pd.DataFrame, columns: List[str]) -> pd.Series:
        """Crée une clé composite à partir de plusieurs colonnes"""
        # Nettoyer chaque colonne et les concaténer avec un séparateur
        parts = [self._clean_reference(df[col]) for col in columns]
        # Utiliser || comme séparateur (peu probable d'apparaître dans les données)
        composite = parts[0]
        for part in parts[1:]:
            composite = composite + '||' + part
        return composite
    
    def reconcile(self) -> Tuple[bool, Union[ReconciliationResult, str]]:
        if self.file1_data is None or self.file2_data is None:
            return False, "Les deux fichiers doivent être chargés"
        
        if not self.comparison_keys:
            return False, "Au moins une paire de colonnes de comparaison doit être sélectionnée"
        
        try:
            # Extraire les colonnes de chaque fichier
            file1_cols = [pair[0] for pair in self.comparison_keys]
            file2_cols = [pair[1] for pair in self.comparison_keys]
            
            # Créer les clés composites
            self.file1_data['_composite_key'] = self._create_composite_key(self.file1_data, file1_cols)
            self.file2_data['_composite_key'] = self._create_composite_key(self.file2_data, file2_cols)
            
            # Utiliser les clés composites pour la comparaison
            refs1 = self.file1_data['_composite_key']
            refs2 = self.file2_data['_composite_key']
            
            if self.aggregation_mode and self.amount_col1 and self.amount_col2:
                # Mode agrégé : sommer les montants par clé composite
                file1_agg = self.file1_data.groupby('_composite_key').agg({
                    self.amount_col1: 'sum',
                    **{col: 'first' for col in self.file1_data.columns 
                       if col not in [self.amount_col1, '_composite_key']}
                }).reset_index()
                
                file2_agg = self.file2_data.groupby('_composite_key').agg({
                    self.amount_col2: 'sum',
                    **{col: 'first' for col in self.file2_data.columns 
                       if col not in [self.amount_col2, '_composite_key']}
                }).reset_index()
                
                refs1 = file1_agg['_composite_key']
                refs2 = file2_agg['_composite_key']
                working_df1 = file1_agg
                working_df2 = file2_agg
            else:
                working_df1 = self.file1_data
                working_df2 = self.file2_data
            
            # Ensembles pour comparaison
            set1 = set(refs1)
            set2 = set(refs2)
            
            # Différences et intersection
            only_in_1 = set1 - set2
            only_in_2 = set2 - set1
            matched_refs = set1 & set2
            
            # Filtrer DataFrames
            if self.aggregation_mode:
                file1_missing = working_df1[working_df1['_composite_key'].isin(only_in_1)].copy()
                file2_missing = working_df2[working_df2['_composite_key'].isin(only_in_2)].copy()
                file1_matched = working_df1[working_df1['_composite_key'].isin(matched_refs)].copy()
                file2_matched = working_df2[working_df2['_composite_key'].isin(matched_refs)].copy()
            else:
                file1_missing = self.file1_data[self.file1_data['_composite_key'].isin(only_in_1)].copy()
                file2_missing = self.file2_data[self.file2_data['_composite_key'].isin(only_in_2)].copy()
                file1_matched = self.file1_data[self.file1_data['_composite_key'].isin(matched_refs)].copy()
                file2_matched = self.file2_data[self.file2_data['_composite_key'].isin(matched_refs)].copy()
            
            # Nettoyer les colonnes temporaires
            for df in [file1_missing, file2_missing, file1_matched, file2_matched]:
                if '_composite_key' in df.columns:
                    df.drop(columns=['_composite_key'], inplace=True)
            
            # Calculs des montants sur les MATCHED
            amount1_total = None
            amount2_total = None
            amount_diff = None
            discrepancies_df = pd.DataFrame()
            discrepancy_count = 0
            
            if (self.amount_col1 and self.amount_col1 in working_df1.columns and
                self.amount_col2 and self.amount_col2 in working_df2.columns):
                
                file1_matched[self.amount_col1] = pd.to_numeric(
                    file1_matched[self.amount_col1], errors='coerce'
                )
                file2_matched[self.amount_col2] = pd.to_numeric(
                    file2_matched[self.amount_col2], errors='coerce'
                )
                
                amount1_total = file1_matched[self.amount_col1].sum()
                amount2_total = file2_matched[self.amount_col2].sum()
                amount_diff = amount1_total - amount2_total
                
                # Identifier les écarts ligne par ligne
                discrepancies = []
                
                # Recréer les clés composites pour comparaison
                file1_matched['_ref'] = self._create_composite_key(
                    file1_matched, [pair[0] for pair in self.comparison_keys]
                ) if not self.aggregation_mode else file1_matched.index.astype(str)
                
                file2_matched['_ref'] = self._create_composite_key(
                    file2_matched, [pair[1] for pair in self.comparison_keys]
                ) if not self.aggregation_mode else file2_matched.index.astype(str)
                
                for ref in matched_refs:
                    rows1 = file1_matched[file1_matched['_ref'] == ref]
                    rows2 = file2_matched[file2_matched['_ref'] == ref]
                    
                    if len(rows1) > 0 and len(rows2) > 0:
                        amt1 = rows1[self.amount_col1].iloc[0]
                        amt2 = rows2[self.amount_col2].iloc[0]
                        
                        if pd.notna(amt1) and pd.notna(amt2) and abs(amt1 - amt2) > 0.01:
                            discrepancies.append({
                                'Reference': ref.replace('||', ' + '),
                                f'Montant_{self.file1_name}': amt1,
                                f'Montant_{self.file2_name}': amt2,
                                'Ecart': amt1 - amt2,
                                'Ecart_pct': ((amt1 - amt2) / amt2 * 100) if amt2 != 0 else None
                            })
                
                if discrepancies:
                    discrepancies_df = pd.DataFrame(discrepancies)
                    discrepancy_count = len(discrepancies)
                
                for df in [file1_matched, file2_matched]:
                    if '_ref' in df.columns:
                        df.drop(columns=['_ref'], inplace=True)
            
            # Nettoyer les données originales
            if '_composite_key' in self.file1_data.columns:
                self.file1_data.drop(columns=['_composite_key'], inplace=True)
            if '_composite_key' in self.file2_data.columns:
                self.file2_data.drop(columns=['_composite_key'], inplace=True)
            
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
                discrepancy_count=discrepancy_count,
                aggregation_mode=self.aggregation_mode,
                comparison_keys=self.comparison_keys
            )
            
            return True, result
            
        except Exception as e:
            return False, f"Erreur lors de la réconciliation : {str(e)}"