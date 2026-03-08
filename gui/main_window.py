# Interface principale
from PySide6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
    QPushButton, QLabel, QComboBox, QFileDialog, 
    QTableWidget, QTableWidgetItem, QGroupBox, 
    QMessageBox, QSplitter, QFrame,
    QHeaderView, QStatusBar, QGridLayout
)
from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QFont
from pathlib import Path

from core.reconciliator import Reconciliator, ReconciliationResult
from utils.excel_handler import ExcelHandler

class DropArea(QFrame):
    """Zone de dépôt de fichiers"""
    file_loaded = Signal(str)  # Signal émis quand un fichier est chargé
    
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
        self.label = QLabel(f"📁 {title}\n\nGlissez-déposez un fichier Excel\nou cliquez pour parcourir")
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
            if file_path.endswith(('.xlsx', '.xls')):
                self.file_path = file_path
                self.label.setText(f"✓ {Path(file_path).name}")
                self.label.setStyleSheet("color: #38a169; font-weight: bold;")
                # Émettre le signal pour charger le fichier
                self.file_loaded.emit(file_path)
            else:
                self.label.setText("✗ Format non valide (Excel requis)")
                self.label.setStyleSheet("color: #e53e3e;")

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.reconciliator = Reconciliator()
        self.current_result = None
        
        self.setWindowTitle("Réconciliation de Fichiers Excel")
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
        files_group = QGroupBox("📂 Chargement des Fichiers")
        files_layout = QHBoxLayout(files_group)
        
        self.file1_drop = DropArea("Fichier Source 1")
        self.file1_drop.file_loaded.connect(self.load_file1_from_drop)  # Connecter signal
        self.file1_drop.mousePressEvent = lambda e: self.select_file(1)
        files_layout.addWidget(self.file1_drop)
        
        self.file2_drop = DropArea("Fichier Source 2")
        self.file2_drop.file_loaded.connect(self.load_file2_from_drop)  # Connecter signal
        self.file2_drop.mousePressEvent = lambda e: self.select_file(2)
        files_layout.addWidget(self.file2_drop)
        
        main_layout.addWidget(files_group)
        
        # Section configuration
        config_group = QGroupBox("⚙️ Configuration des Colonnes")
        config_layout = QGridLayout(config_group)
        
        # Fichier 1
        config_layout.addWidget(QLabel("Fichier 1 - Colonne Référence:"), 0, 0)
        self.ref1_combo = QComboBox()
        self.ref1_combo.setEnabled(False)
        config_layout.addWidget(self.ref1_combo, 0, 1)
        
        config_layout.addWidget(QLabel("Fichier 1 - Colonne Montant (optionnel):"), 1, 0)
        self.amount1_combo = QComboBox()
        self.amount1_combo.setEnabled(False)
        self.amount1_combo.addItem("(Aucune)")
        config_layout.addWidget(self.amount1_combo, 1, 1)
        
        # Fichier 2
        config_layout.addWidget(QLabel("Fichier 2 - Colonne Référence:"), 0, 2)
        self.ref2_combo = QComboBox()
        self.ref2_combo.setEnabled(False)
        config_layout.addWidget(self.ref2_combo, 0, 3)
        
        config_layout.addWidget(QLabel("Fichier 2 - Colonne Montant (optionnel):"), 1, 2)
        self.amount2_combo = QComboBox()
        self.amount2_combo.setEnabled(False)
        self.amount2_combo.addItem("(Aucune)")
        config_layout.addWidget(self.amount2_combo, 1, 3)
        
        # Connecter les changements
        self.ref1_combo.currentTextChanged.connect(self.check_ready)
        self.ref2_combo.currentTextChanged.connect(self.check_ready)
        
        main_layout.addWidget(config_group)
        
        # Bouton de réconciliation
        self.reconcile_btn = QPushButton("🚀 Lancer la Réconciliation")
        self.reconcile_btn.setEnabled(False)
        self.reconcile_btn.setMinimumHeight(50)
        self.reconcile_btn.setFont(QFont("Segoe UI", 12, QFont.Bold))
        self.reconcile_btn.clicked.connect(self.run_reconciliation)
        main_layout.addWidget(self.reconcile_btn)
        
        # Section résultats
        results_group = QGroupBox("📊 Résultats")
        results_layout = QVBoxLayout(results_group)
        
        # Statistiques
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
        
        # Splitter pour les tableaux
        splitter = QSplitter(Qt.Vertical)
        
        # Tableau 1: Missing dans fichier 1
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
        
        # Tableau 2: Missing dans fichier 2
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
        
        # Bouton export
        self.export_btn = QPushButton("💾 Exporter vers Excel")
        self.export_btn.setEnabled(False)
        self.export_btn.setMinimumHeight(40)
        self.export_btn.clicked.connect(self.export_results)
        results_layout.addWidget(self.export_btn)
        
        main_layout.addWidget(results_group, stretch=1)
        
        # Barre de statut
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
        """Charger le fichier 1 depuis le drop"""
        self.file1_drop.file_path = file_path
        success, message = self.reconciliator.load_file1(file_path)
        if success:
            self.update_combos(1)
        self.status_bar.showMessage(message)
    
    def load_file2_from_drop(self, file_path):
        """Charger le fichier 2 depuis le drop"""
        self.file2_drop.file_path = file_path
        success, message = self.reconciliator.load_file2(file_path)
        if success:
            self.update_combos(2)
        self.status_bar.showMessage(message)
    
    def select_file(self, file_num):
        """Ouvre le dialogue de sélection de fichier"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, f"Sélectionner le fichier {file_num}", "", "Excel Files (*.xlsx *.xls)"
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
            self.ref1_combo.clear()
            self.ref1_combo.addItems(cols)
            self.ref1_combo.setEnabled(True)
            
            self.amount1_combo.clear()
            self.amount1_combo.addItem("(Aucune)")
            self.amount1_combo.addItems(cols)
            self.amount1_combo.setEnabled(True)
        else:
            cols = self.reconciliator.get_file2_columns()
            self.ref2_combo.clear()
            self.ref2_combo.addItems(cols)
            self.ref2_combo.setEnabled(True)
            
            self.amount2_combo.clear()
            self.amount2_combo.addItem("(Aucune)")
            self.amount2_combo.addItems(cols)
            self.amount2_combo.setEnabled(True)
    
    def check_ready(self):
        ready = (self.ref1_combo.currentText() != "" and 
                self.ref2_combo.currentText() != "")
        self.reconcile_btn.setEnabled(ready)
    
    def run_reconciliation(self):
        # Récupérer les colonnes sélectionnées
        amount1 = self.amount1_combo.currentText()
        amount2 = self.amount2_combo.currentText()
        
        if amount1 == "(Aucune)":
            amount1 = None
        if amount2 == "(Aucune)":
            amount2 = None
        
        self.reconciliator.set_columns(
            self.ref1_combo.currentText(),
            self.ref2_combo.currentText(),
            amount1,
            amount2
        )
        
        self.status_bar.showMessage("Réconciliation en cours...")
        self.reconcile_btn.setEnabled(False)
        
        success, result = self.reconciliator.reconcile()
        
        if not success:
            QMessageBox.critical(self, "Erreur", str(result))
            self.status_bar.showMessage("Erreur de réconciliation")
            self.reconcile_btn.setEnabled(True)
            return
        
        self.current_result = result
        self.display_results(result)
        
        self.status_bar.showMessage("Réconciliation terminée")
        self.reconcile_btn.setEnabled(True)
        self.export_btn.setEnabled(True)
    
    def display_results(self, result: ReconciliationResult):
        # Mettre à jour les labels avec les vrais noms de fichiers
        self.label_missing1.setText(f"❌ Transactions uniquement dans {result.file1_name} ({len(result.file1_missing)})")
        self.label_missing2.setText(f"❌ Transactions uniquement dans {result.file2_name} ({len(result.file2_missing)})")
        
        # Statistiques
        match_rate = (result.matched_count / max(result.file1_total, result.file2_total)) * 100
        
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
                <td style='padding: 5px 20px; color: #48bb78;'>{result.matched_count}</td>
                <td style='padding: 5px 20px;'><b>📊 Taux:</b></td>
                <td style='padding: 5px 20px; color: {"#48bb78" if match_rate > 90 else "#ed8936" if match_rate > 70 else "#e53e3e"};'>{match_rate:.1f}%</td>
            </tr>
        """
        
        # Ajouter les montants si disponibles (sur les matched uniquement)
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
        
        # Tableaux
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