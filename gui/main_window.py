# Interface principale
from PySide6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
    QPushButton, QLabel, QComboBox, QFileDialog, 
    QTableWidget, QTableWidgetItem, QGroupBox, 
    QMessageBox, QSplitter, QFrame,
    QHeaderView, QStatusBar, QGridLayout,
    QDialog
)
from PySide6.QtCore import Qt, Signal, QObject, QThread
from PySide6.QtGui import QFont, QCursor
from pathlib import Path

from core.reconciliator import Reconciliator, ReconciliationResult
from utils.excel_handler import ExcelHandler

class DropArea(QFrame):
    """Zone de dépôt de fichiers 211101"""
    file_loaded = Signal(str)
    
    def __init__(self, title, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.setMinimumHeight(100)
        self.setFrameStyle(QFrame.StyledPanel | QFrame.Sunken)
        self.setStyleSheet("""
            DropArea {
                background-color: #f0f4f8;
                border: 2px dashed #cbd5e0;
                border-radius: 8px;
            }
            DropArea[active="true"] {
                background-color: #e6fffa;
                border-color: #38b2ac;
            }
        """)
        
        layout = QVBoxLayout(self)
        self.label = QLabel(f"📁 {title}\n\nGlissez-déposez un fichier\nExcel ou CSV\nou cliquez pour parcourir")
        self.label.setAlignment(Qt.AlignCenter)
        self.label.setStyleSheet("color: #4a5568; font-size: 12px;")
        layout.addWidget(self.label)
        
        self.file_path = None
        
    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            self.setProperty("active", "true")
            self.style().unpolish(self)
            self.style().polish(self)
            event.acceptProposedAction()
    
    def dragLeaveEvent(self, event):
        self.setProperty("active", "false")
        self.style().unpolish(self)
        self.style().polish(self)
    
    def dropEvent(self, event):
        self.setProperty("active", "false")
        self.style().unpolish(self)
        self.style().polish(self)
        
        urls = event.mimeData().urls()
        if urls:
            file_path = urls[0].toLocalFile()
            if self.is_valid_file(file_path):
                self.file_path = file_path
                self.label.setText(f"✓ {Path(file_path).name}")
                self.label.setStyleSheet("color: #38a169; font-weight: bold;")
                self.file_loaded.emit(file_path)
            else:
                self.label.setText("✗ Format non valide\n(Excel ou CSV requis)")
                self.label.setStyleSheet("color: #e53e3e;")
    
    def is_valid_file(self, file_path):
        """Vérifie si le fichier est au format accepté"""
        valid_extensions = ('.xlsx', '.xls', '.csv')
        return file_path.lower().endswith(valid_extensions)

class ReconciliationWorker(QObject):
    """Worker thread pour la réconciliation sans bloquer l'UI"""
    finished = Signal(object)
    error = Signal(str)
    
    def __init__(self, reconciliator):
        super().__init__()
        self.reconciliator = reconciliator
    
    def run(self):
        try:
            success, result = self.reconciliator.reconcile()
            
            if success:
                self.finished.emit(result)
            else:
                self.error.emit(str(result))
        except Exception as e:
            self.error.emit(str(e))

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.reconciliator = Reconciliator()
        self.current_result = None
        self.worker = None
        self.thread = None
        self.comparison_keys = []  # Liste des paires (col_fichier1, col_fichier2)
        
        self.setWindowTitle("Réconciliation de Fichiers Excel/CSV")
        self.setMinimumSize(1400, 900)
        
        self.setup_ui()
        self.apply_styles()
    
    def setup_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        main_layout = QVBoxLayout(central_widget)
        main_layout.setSpacing(20)
        main_layout.setContentsMargins(20, 20, 20, 20)
        
        # En-tête
        header = QLabel("🔍 Outil de Réconciliation de Transactions")
        header.setFont(QFont("Segoe UI", 20, QFont.Bold))
        header.setStyleSheet("color: #2d3748; margin-bottom: 10px;")
        header.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(header)
        
        # Section fichiers
        files_group = QGroupBox("📂 Chargement des Fichiers (Excel ou CSV)")
        files_layout = QHBoxLayout(files_group)
        
        self.file1_drop = DropArea("Fichier Source 1")
        self.file1_drop.file_loaded.connect(self.load_file1_from_drop)
        self.file1_drop.mousePressEvent = lambda e: self.select_file(1)
        files_layout.addWidget(self.file1_drop)
        
        self.file2_drop = DropArea("Fichier Source 2")
        self.file2_drop.file_loaded.connect(self.load_file2_from_drop)
        self.file2_drop.mousePressEvent = lambda e: self.select_file(2)
        files_layout.addWidget(self.file2_drop)
        
        main_layout.addWidget(files_group)
        
        # Section configuration des clés de comparaison
        keys_group = QGroupBox("🔑 Clés de Comparaison (Réconciliation)")
        keys_layout = QVBoxLayout(keys_group)
        
        # Liste des paires de colonnes
        self.keys_list_widget = QWidget()
        self.keys_list_layout = QVBoxLayout(self.keys_list_widget)
        self.keys_list_layout.setSpacing(5)
        self.keys_list_layout.setContentsMargins(0, 0, 0, 0)
        
        # Message initial
        self.keys_empty_label = QLabel("Aucune clé de comparaison définie. Cliquez sur 'Ajouter' pour commencer.")
        self.keys_empty_label.setStyleSheet("color: #718096; font-style: italic;")
        self.keys_list_layout.addWidget(self.keys_empty_label)
        
        keys_layout.addWidget(self.keys_list_widget)
        
        # Boutons d'action
        keys_buttons_layout = QHBoxLayout()
        
        self.add_key_btn = QPushButton("➕ Ajouter une colonne de comparaison")
        self.add_key_btn.setEnabled(False)
        self.add_key_btn.clicked.connect(self.add_comparison_key)
        keys_buttons_layout.addWidget(self.add_key_btn)
        
        keys_layout.addLayout(keys_buttons_layout)
        
        # Info
        keys_info = QLabel("💡 Astuce: Ajoutez plusieurs colonnes pour une réconciliation plus précise (ex: ID + Date + Client)")
        keys_info.setStyleSheet("color: #4a5568; font-size: 11px;")
        keys_layout.addWidget(keys_info)
        
        main_layout.addWidget(keys_group)
        
        # Section montants (optionnel)
        amounts_group = QGroupBox("💰 Colonnes de Montant (Optionnel)")
        amounts_layout = QGridLayout(amounts_group)
        
        amounts_layout.addWidget(QLabel("Fichier 1 - Montant:"), 0, 0)
        self.amount1_combo = QComboBox()
        self.amount1_combo.setEnabled(False)
        self.amount1_combo.addItem("(Aucune)")
        amounts_layout.addWidget(self.amount1_combo, 0, 1)
        
        amounts_layout.addWidget(QLabel("Fichier 2 - Montant:"), 0, 2)
        self.amount2_combo = QComboBox()
        self.amount2_combo.setEnabled(False)
        self.amount2_combo.addItem("(Aucune)")
        amounts_layout.addWidget(self.amount2_combo, 0, 3)
        
        main_layout.addWidget(amounts_group)
        
        # Section boutons (Réconciliation + Réinitialisation)
        buttons_layout = QHBoxLayout()
        
        # Bouton de réconciliation
        self.reconcile_btn = QPushButton("🚀 Lancer la Réconciliation")
        self.reconcile_btn.setEnabled(False)
        self.reconcile_btn.setMinimumHeight(50)
        self.reconcile_btn.setFont(QFont("Segoe UI", 12, QFont.Bold))
        self.reconcile_btn.setCursor(QCursor(Qt.PointingHandCursor))
        self.reconcile_btn.clicked.connect(self.run_reconciliation)
        buttons_layout.addWidget(self.reconcile_btn, stretch=4)
        
        # Bouton Réinitialisation
        self.reset_btn = QPushButton("🔄 Réinitialisation")
        self.reset_btn.setObjectName("reset_btn")
        self.reset_btn.setToolTip("Effacer tous les fichiers et recommencer")
        self.reset_btn.setMinimumHeight(50)
        self.reset_btn.setFont(QFont("Segoe UI", 12, QFont.Bold))
        self.reset_btn.setCursor(QCursor(Qt.PointingHandCursor))
        self.reset_btn.clicked.connect(self.reset_all)
        buttons_layout.addWidget(self.reset_btn, stretch=1)
        
        main_layout.addLayout(buttons_layout)
        
        # Label de statut réconciliation (caché par défaut)
        self.status_reconciliation = QLabel("⏳ Réconciliation en cours...")
        self.status_reconciliation.setAlignment(Qt.AlignCenter)
        self.status_reconciliation.setStyleSheet("""
            QLabel {
                background-color: #ebf8ff;
                color: #2b6cb0;
                padding: 10px;
                border-radius: 6px;
                font-weight: bold;
                font-size: 14px;
            }
        """)
        self.status_reconciliation.setVisible(False)
        main_layout.addWidget(self.status_reconciliation)
        
        # Section résultats
        results_group = QGroupBox("📊 Résultats")
        results_layout = QVBoxLayout(results_group)
        
        self.stats_label = QLabel("Aucune réconciliation effectuée")
        self.stats_label.setAlignment(Qt.AlignCenter)
        self.stats_label.setStyleSheet("""
            QLabel {
                background-color: #edf2f7;
                padding: 15px;
                border-radius: 8px;
                font-size: 14px;
                color: #4a5568;
            }
        """)
        results_layout.addWidget(self.stats_label)
        
        splitter = QSplitter(Qt.Vertical)
        
        w1 = QWidget()
        l1 = QVBoxLayout(w1)
        l1.setContentsMargins(0, 0, 0, 0)
        self.label_missing1 = QLabel("❌ Transactions uniquement dans Fichier 1")
        self.label_missing1.setFont(QFont("Segoe UI", 10, QFont.Bold))
        l1.addWidget(self.label_missing1)
        self.table_missing1 = QTableWidget()
        self.table_missing1.setAlternatingRowColors(True)
        l1.addWidget(self.table_missing1)
        splitter.addWidget(w1)
        
        w2 = QWidget()
        l2 = QVBoxLayout(w2)
        l2.setContentsMargins(0, 0, 0, 0)
        self.label_missing2 = QLabel("❌ Transactions uniquement dans Fichier 2")
        self.label_missing2.setFont(QFont("Segoe UI", 10, QFont.Bold))
        l2.addWidget(self.label_missing2)
        self.table_missing2 = QTableWidget()
        self.table_missing2.setAlternatingRowColors(True)
        l2.addWidget(self.table_missing2)
        splitter.addWidget(w2)
        
        results_layout.addWidget(splitter)
        
        self.export_btn = QPushButton("💾 Exporter vers Excel")
        self.export_btn.setEnabled(False)
        self.export_btn.setMinimumHeight(40)
        self.export_btn.setCursor(QCursor(Qt.PointingHandCursor))
        self.export_btn.clicked.connect(self.export_results)
        results_layout.addWidget(self.export_btn)
        
        main_layout.addWidget(results_group, stretch=1)
        
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage("Prêt")
    
    def apply_styles(self):
        self.setStyleSheet("""
            QMainWindow { 
                background-color: #f7fafc; 
            }
            QGroupBox {
                font-weight: bold;
                border: 2px solid #e2e8f0;
                border-radius: 8px;
                margin-top: 10px;
                padding-top: 10px;
                color: #2d3748;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px;
            }
            QPushButton {
                background-color: #3182ce;
                color: white;
                border: none;
                border-radius: 6px;
                padding: 10px 20px;
                font-weight: bold;
            }
            QPushButton:hover { 
                background-color: #2c5282; 
            }
            QPushButton:disabled {
                background-color: #cbd5e0;
                color: #718096;
            }
            QPushButton#reset_btn {
                background-color: #4299e1;
            }
            QPushButton#reset_btn:hover {
                background-color: #3182ce;
            }
            QComboBox {
                padding: 8px;
                border: 2px solid #4a5568;
                border-radius: 6px;
                background-color: white;
                color: #1a202c;
                min-width: 200px;
                font-weight: normal;
            }
            QComboBox:disabled { 
                background-color: #e2e8f0; 
                color: #a0aec0;
            }
            QComboBox::drop-down {
                border: none;
                padding-right: 10px;
            }
            QComboBox QAbstractItemView {
                background-color: white;
                color: #1a202c;
                selection-background-color: #4299e1;
                selection-color: white;
                border: 1px solid #e2e8f0;
            }
            QComboBox::selected {
                background-color: #ebf8ff;
                color: #2b6cb0;
                font-weight: bold;
            }
            QLabel {
                color: #2d3748;
            }
            QTableWidget {
                border: 2px solid #4a5568;
                border-radius: 6px;
                background-color: white;
                color: #1a202c;
                gridline-color: #cbd5e0;
                alternate-background-color: #f7fafc;
            }
            QTableWidget::item {
                color: #1a202c;
                padding: 5px;
            }
            QTableWidget::item:selected {
                background-color: #4299e1;
                color: white;
            }
            QTableWidget::item:alternate {
                background-color: #edf2f7;
                color: #1a202c;
            }
            QHeaderView::section {
                background-color: #4a5568;
                color: white;
                padding: 10px;
                border: none;
                font-weight: bold;
            }
            QStatusBar {
                background-color: #edf2f7;
                color: #2d3748;
            }
        """)
    
    def load_file1_from_drop(self, file_path):
        self.file1_drop.file_path = file_path
        success, message = self.reconciliator.load_file1(file_path)
        if success:
            self.update_combos(1)
        self.status_bar.showMessage(message)
    
    def load_file2_from_drop(self, file_path):
        self.file2_drop.file_path = file_path
        success, message = self.reconciliator.load_file2(file_path)
        if success:
            self.update_combos(2)
        self.status_bar.showMessage(message)
    
    def select_file(self, file_num):
        file_path, _ = QFileDialog.getOpenFileName(
            self, 
            f"Sélectionner le fichier {file_num}", 
            "", 
            "Fichiers supportés (*.xlsx *.xls *.csv);;Excel (*.xlsx *.xls);;CSV (*.csv)"
        )
        
        if not file_path:
            return
            
        if file_num == 1:
            self.file1_drop.file_path = file_path
            self.file1_drop.label.setText(f"✓ {Path(file_path).name}")
            self.file1_drop.label.setStyleSheet("color: #38a169; font-weight: bold;")
            success, message = self.reconciliator.load_file1(file_path)
            if success:
                self.update_combos(1)
        else:
            self.file2_drop.file_path = file_path
            self.file2_drop.label.setText(f"✓ {Path(file_path).name}")
            self.file2_drop.label.setStyleSheet("color: #38a169; font-weight: bold;")
            success, message = self.reconciliator.load_file2(file_path)
            if success:
                self.update_combos(2)
        
        self.status_bar.showMessage(message)
    
    def update_combos(self, file_num):
        if file_num == 1:
            cols = self.reconciliator.get_file1_columns()
            self.amount1_combo.clear()
            self.amount1_combo.addItem("(Aucune)")
            self.amount1_combo.addItems(cols)
            self.amount1_combo.setEnabled(True)
        else:
            cols = self.reconciliator.get_file2_columns()
            self.amount2_combo.clear()
            self.amount2_combo.addItem("(Aucune)")
            self.amount2_combo.addItems(cols)
            self.amount2_combo.setEnabled(True)
        
        # Activer le bouton d'ajout si les deux fichiers sont chargés
        if (self.reconciliator.file1_data is not None and 
            self.reconciliator.file2_data is not None):
            self.add_key_btn.setEnabled(True)
    
    def check_ready(self):
        ready = len(self.comparison_keys) > 0
        self.reconcile_btn.setEnabled(ready)
    
    def run_reconciliation(self):
        amount1 = self.amount1_combo.currentText()
        amount2 = self.amount2_combo.currentText()
        
        if amount1 == "(Aucune)":
            amount1 = None
        if amount2 == "(Aucune)":
            amount2 = None
        
        self.reconciliator.set_comparison_keys(
            self.comparison_keys,
            amount1,
            amount2
        )
        
        self.reconcile_btn.setEnabled(False)
        self.reconcile_btn.setText("⏳ Patientez...")
        self.status_reconciliation.setVisible(True)
        self.status_bar.showMessage("Réconciliation en cours...")
        
        self.thread = QThread()
        self.worker = ReconciliationWorker(self.reconciliator)
        self.worker.moveToThread(self.thread)
        
        self.thread.started.connect(self.worker.run)
        self.worker.finished.connect(self.handle_reconciliation_success)
        self.worker.error.connect(self.handle_reconciliation_error)
        self.worker.finished.connect(self.thread.quit)
        self.worker.error.connect(self.thread.quit)
        self.worker.finished.connect(self.worker.deleteLater)
        self.worker.error.connect(self.worker.deleteLater)
        self.thread.finished.connect(self.thread.deleteLater)
        
        self.thread.start()
    
    def handle_reconciliation_success(self, result):
        self.current_result = result
        self.display_results(result)
        
        self.reconcile_btn.setEnabled(True)
        self.reconcile_btn.setText("🚀 Lancer la Réconciliation")
        self.status_reconciliation.setVisible(False)
        self.export_btn.setEnabled(True)
        self.status_bar.showMessage("Réconciliation terminée avec succès")
    
    def handle_reconciliation_error(self, error_msg):
        QMessageBox.critical(self, "Erreur", error_msg)
        
        self.reconcile_btn.setEnabled(True)
        self.reconcile_btn.setText("🚀 Lancer la Réconciliation")
        self.status_reconciliation.setVisible(False)
        self.status_bar.showMessage("Erreur de réconciliation")
    
    def display_results(self, result: ReconciliationResult):
        self.label_missing1.setText(f"❌ Transactions uniquement dans {result.file1_name} ({len(result.file1_missing)})")
        self.label_missing2.setText(f"❌ Transactions uniquement dans {result.file2_name} ({len(result.file2_missing)})")
        
        # Calcul du taux basé sur les lignes réelles (strict 1:1)
        file1_match_rate = (result.file1_matched_lines / result.file1_total * 100) if result.file1_total > 0 else 0
        file2_match_rate = (result.file2_matched_lines / result.file2_total * 100) if result.file2_total > 0 else 0
        match_rate = min(file1_match_rate, file2_match_rate)  # Taux le plus conservateur
        
        stats_html = f"""
        <table style='margin: 0 auto; font-size: 14px;'>
            <tr>
                <td style='padding: 5px 20px;'><b>📁 {result.file1_name}:</b></td>
                <td style='padding: 5px 20px; color: #4299e1;'>{result.file1_total} transactions</td>
                <td style='padding: 5px 20px;'><b>📁 {result.file2_name}:</b></td>
                <td style='padding: 5px 20px; color: #4299e1;'>{result.file2_total} transactions</td>
            </tr>
            <tr>
                <td style='padding: 5px 20px;'><b>✅ Correspondances:</b></td>
                <td style='padding: 5px 20px; color: #48bb78;'>{result.file1_matched_lines}</td>
                <td style='padding: 5px 20px;'><b>📊 Taux:</b></td>
                <td style='padding: 5px 20px; color: {"#48bb78" if match_rate > 90 else "#ed8936" if match_rate > 70 else "#e53e3e"}; font-weight: bold;'>{match_rate:.1f}%</td>
            </tr>
        """
        
        if result.file1_matched_amount_total is not None:
            stats_html += f"""
            <tr style='background-color: #ebf8ff;'>
                <td style='padding: 5px 20px;'><b>💰 Total matched {result.file1_name}:</b></td>
                <td style='padding: 5px 20px; color: #2b6cb0;'>{result.file1_matched_amount_total:,.2f}</td>
            """
            
            if result.file2_matched_amount_total is not None:
                stats_html += f"""
                <td style='padding: 5px 20px;'><b>💰 Total matched {result.file2_name}:</b></td>
                <td style='padding: 5px 20px; color: #2b6cb0;'>{result.file2_matched_amount_total:,.2f}</td>
            </tr>
            <tr style='background-color: {"#fed7d7" if result.amount_difference != 0 else "#c6f6d5"};'>
                <td style='padding: 5px 20px;'><b>📉 Écart total:</b></td>
                <td style='padding: 5px 20px; color: {"#c53030" if result.amount_difference != 0 else "#276749"}; font-weight: bold;'>{result.amount_difference:,.2f}</td>
                <td style='padding: 5px 20px;'><b>⚠️ Lignes avec écart:</b></td>
                <td style='padding: 5px 20px; color: {"#c53030" if result.discrepancy_count > 0 else "#276749"}; font-weight: bold;'>{result.discrepancy_count}</td>
            </tr>
                """
            else:
                stats_html += "</tr>"
        
        stats_html += "</table>"
        self.stats_label.setText(stats_html)
        
        # CORRECTION : Utiliser self.table_missing2 et self.table_missing1 (pas self.table.table_missing2)
        self.populate_table(self.table_missing1, result.file1_missing)
        self.populate_table(self.table_missing2, result.file2_missing)
    
    def populate_table(self, table: QTableWidget, df):
        if len(df) == 0:
            table.setColumnCount(1)
            table.setRowCount(1)
            table.setItem(0, 0, QTableWidgetItem("Aucune transaction"))
            return
        
        table.setColumnCount(len(df.columns))
        table.setRowCount(len(df))
        table.setHorizontalHeaderLabels(list(df.columns))
        
        for i, (_, row) in enumerate(df.iterrows()):
            for j, value in enumerate(row):
                item = QTableWidgetItem(str(value))
                item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                table.setItem(i, j, item)
        
        table.resizeColumnsToContents()
        table.horizontalHeader().setStretchLastSection(True)
    
    def export_results(self):
        if not self.current_result:
            return
        
        default_name = f"reconciliation_{self.current_result.file1_name}_{self.current_result.file2_name}.xlsx"
        
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "Exporter les résultats",
            default_name,
            "Excel Files (*.xlsx)"
        )
        
        if file_path:
            success, message = ExcelHandler.export_results(self.current_result, file_path)
            if success:
                QMessageBox.information(self, "Succès", message)
                self.status_bar.showMessage(message)
            else:
                QMessageBox.critical(self, "Erreur", message)
                self.status_bar.showMessage(message)

    def reset_all(self):
        self.reconciliator = Reconciliator()
        self.current_result = None
        
        self.file1_drop.file_path = None
        self.file1_drop.label.setText("📁 Fichier Source 1\n\nGlissez-déposez un fichier\nExcel ou CSV\nou cliquez pour parcourir")
        self.file1_drop.label.setStyleSheet("color: #4a5568; font-size: 12px;")
        
        self.file2_drop.file_path = None
        self.file2_drop.label.setText("📁 Fichier Source 2\n\nGlissez-déposez un fichier\nExcel ou CSV\nou cliquez pour parcourir")
        self.file2_drop.label.setStyleSheet("color: #4a5568; font-size: 12px;")
        
        self.amount1_combo.clear()
        self.amount1_combo.addItem("(Aucune)")
        self.amount1_combo.setEnabled(False)
        
        self.amount2_combo.clear()
        self.amount2_combo.addItem("(Aucune)")
        self.amount2_combo.setEnabled(False)
        
        self.reconcile_btn.setEnabled(False)
        self.export_btn.setEnabled(False)
        
        self.comparison_keys = []
        self.refresh_keys_list()
        self.add_key_btn.setEnabled(False)
        
        self.stats_label.setText("Aucune réconciliation effectuée")
        self.table_missing1.clear()
        self.table_missing1.setRowCount(0)
        self.table_missing1.setColumnCount(0)
        self.table_missing2.clear()
        self.table_missing2.setRowCount(0)
        self.table_missing2.setColumnCount(0)
        self.label_missing1.setText("❌ Transactions uniquement dans Fichier 1")
        self.label_missing2.setText("❌ Transactions uniquement dans Fichier 2")
        
        self.status_reconciliation.setVisible(False)
        
        self.status_bar.showMessage("Prêt - Tout a été réinitialisé")

    def add_comparison_key(self):
        """Ajoute une nouvelle paire de colonnes de comparaison"""
        dialog = QDialog(self)
        dialog.setWindowTitle("Ajouter une clé de comparaison")
        dialog.setMinimumWidth(400)
        
        layout = QVBoxLayout(dialog)
        
        layout.addWidget(QLabel("Colonne dans Fichier 1:"))
        combo1 = QComboBox()
        combo1.addItems(self.reconciliator.get_file1_columns())
        layout.addWidget(combo1)
        
        layout.addWidget(QLabel("Colonne dans Fichier 2:"))
        combo2 = QComboBox()
        combo2.addItems(self.reconciliator.get_file2_columns())
        layout.addWidget(combo2)
        
        buttons = QHBoxLayout()
        btn_ok = QPushButton("Ajouter")
        btn_ok.clicked.connect(dialog.accept)
        btn_cancel = QPushButton("Annuler")
        btn_cancel.clicked.connect(dialog.reject)
        buttons.addWidget(btn_ok)
        buttons.addWidget(btn_cancel)
        layout.addLayout(buttons)
        
        if dialog.exec() == QDialog.Accepted:
            col1 = combo1.currentText()
            col2 = combo2.currentText()
            
            if col1 and col2:
                self.comparison_keys.append((col1, col2))
                self.refresh_keys_list()
                self.check_ready()
    
    def refresh_keys_list(self):
        """Rafraîchit l'affichage des clés de comparaison"""
        while self.keys_list_layout.count():
            item = self.keys_list_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
        
        if not self.comparison_keys:
            self.keys_empty_label = QLabel("Aucune clé de comparaison définie. Cliquez sur 'Ajouter' pour commencer.")
            self.keys_empty_label.setStyleSheet("color: #718096; font-style: italic;")
            self.keys_list_layout.addWidget(self.keys_empty_label)
        else:
            for i, (col1, col2) in enumerate(self.comparison_keys):
                key_widget = QWidget()
                key_layout = QHBoxLayout(key_widget)
                key_layout.setContentsMargins(5, 5, 5, 5)
                key_layout.setSpacing(10)
                
                bg_color = "#f7fafc" if i % 2 == 0 else "#edf2f7"
                key_widget.setStyleSheet(f"""
                    QWidget {{
                        background-color: {bg_color};
                        border-radius: 4px;
                        border: 1px solid #e2e8f0;
                    }}
                """)
                
                label = QLabel(f"<b>{i+1}.</b> {col1} ↔ {col2}")
                label.setStyleSheet("color: #2d3748; border: none; background: transparent;")
                key_layout.addWidget(label)
                
                key_layout.addStretch()
                
                btn_delete = QPushButton("🗑️")
                btn_delete.setToolTip("Supprimer cette clé")
                btn_delete.setMaximumWidth(40)
                btn_delete.setStyleSheet("""
                    QPushButton {
                        background-color: #fc8181;
                        color: white;
                        border: none;
                        border-radius: 4px;
                        padding: 2px;
                    }
                    QPushButton:hover {
                        background-color: #f56565;
                    }
                """)
                btn_delete.setCursor(QCursor(Qt.PointingHandCursor))
                btn_delete.clicked.connect(lambda checked, idx=i: self.remove_key(idx))
                key_layout.addWidget(btn_delete)
                
                self.keys_list_layout.addWidget(key_widget)
        
        self.keys_list_layout.addStretch()
    
    def remove_key(self, index):
        """Supprime une clé de comparaison"""
        if 0 <= index < len(self.comparison_keys):
            del self.comparison_keys[index]
            self.refresh_keys_list()
            self.check_ready()