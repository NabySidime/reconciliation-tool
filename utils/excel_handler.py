import pandas as pd
from typing import Tuple
from core.reconciliator import ReconciliationResult

class ExcelHandler:
    @staticmethod
    def export_results(result: ReconciliationResult, output_path: str) -> Tuple[bool, str]:
        try:
            # Nettoyer les noms pour Excel (max 31 caractères) 07112004
            name1 = result.file1_name[:25].replace('/', '-').replace('\\', '-').replace(':', '-')
            name2 = result.file2_name[:25].replace('/', '-').replace('\\', '-').replace(':', '-')
            
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                
                # Feuille 1: Missing dans fichier 1
                sheet1 = f"{name1}_missing"[:31]
                if len(result.file1_missing) > 0:
                    result.file1_missing.to_excel(writer, sheet_name=sheet1, index=False)
                else:
                    pd.DataFrame({'Message': ['Aucune transaction manquante']}).to_excel(
                        writer, sheet_name=sheet1, index=False
                    )
                
                # Feuille 2: Missing dans fichier 2
                sheet2 = f"{name2}_missing"[:31]
                if len(result.file2_missing) > 0:
                    result.file2_missing.to_excel(writer, sheet_name=sheet2, index=False)
                else:
                    pd.DataFrame({'Message': ['Aucune transaction manquante']}).to_excel(
                        writer, sheet_name=sheet2, index=False
                    )
                
                # Feuille 3: Matched fichier 1
                sheet3 = f"{name1}_matched"[:31]
                if len(result.file1_matched) > 0:
                    result.file1_matched.to_excel(writer, sheet_name=sheet3, index=False)
                else:
                    pd.DataFrame({'Message': ['Aucune transaction correspondante']}).to_excel(
                        writer, sheet_name=sheet3, index=False
                    )
                
                # Feuille 4: Matched fichier 2
                sheet4 = f"{name2}_matched"[:31]
                if len(result.file2_matched) > 0:
                    result.file2_matched.to_excel(writer, sheet_name=sheet4, index=False)
                else:
                    pd.DataFrame({'Message': ['Aucune transaction correspondante']}).to_excel(
                        writer, sheet_name=sheet4, index=False
                    )
                
                # Feuille 5: Écarts de montant (NOUVEAU)
                sheet5 = "Ecarts_montant"
                if result.discrepancy_count > 0:
                    result.amount_discrepancies.to_excel(
                        writer, sheet_name=sheet5, index=False
                    )
                else:
                    pd.DataFrame({
                        'Message': ['Aucun écart de montant détecté entre les transactions correspondantes']
                    }).to_excel(writer, sheet_name=sheet5, index=False)
                
                # Feuille 6: Summary
                summary_data = {
                    'Métrique': [
                        f'Total transactions {result.file1_name}',
                        f'Total transactions {result.file2_name}',
                        'Transactions correspondantes',
                        f'Transactions uniquement dans {result.file1_name}',
                        f'Transactions uniquement dans {result.file2_name}',
                        'Taux de correspondance (%)'
                    ],
                    'Valeur': [
                        result.file1_total,
                        result.file2_total,
                        result.matched_count,
                        len(result.file1_missing),
                        len(result.file2_missing),
                        round((result.matched_count / max(result.file1_total, result.file2_total)) * 100, 2)
                    ]
                }
                
                # Ajouter les montants sur les matched
                if result.file1_matched_amount_total is not None:
                    summary_data['Métrique'].append(f'Total montant matched {result.file1_name}')
                    summary_data['Valeur'].append(round(result.file1_matched_amount_total, 2))
                
                if result.file2_matched_amount_total is not None:
                    summary_data['Métrique'].append(f'Total montant matched {result.file2_name}')
                    summary_data['Valeur'].append(round(result.file2_matched_amount_total, 2))
                
                if result.amount_difference is not None:
                    summary_data['Métrique'].append('Écart total de montant')
                    summary_data['Valeur'].append(round(result.amount_difference, 2))
                
                if result.discrepancy_count is not None:
                    summary_data['Métrique'].append('Nombre de lignes avec écart')
                    summary_data['Valeur'].append(result.discrepancy_count)
                
                pd.DataFrame(summary_data).to_excel(writer, sheet_name='Summary', index=False)
            
            return True, f"✓ Export réussi : {output_path}"
        except Exception as e:
            return False, f"✗ Erreur export : {str(e)}"