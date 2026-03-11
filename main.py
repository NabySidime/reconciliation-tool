# Point d'entrée 211101
import sys
from PySide6.QtWidgets import QApplication
from PySide6.QtCore import Qt
from gui.main_window import MainWindow

def main():
    # Activer le support DPI haute résolution
    QApplication.setHighDpiScaleFactorRoundingPolicy(
        Qt.HighDpiScaleFactorRoundingPolicy.PassThrough
    )
    
    app = QApplication(sys.argv)
    app.setApplicationName("Réconciliation OM/PayCard")
    app.setApplicationVersion("1.0.0")
    
    window = MainWindow()
    window.show()
    
    sys.exit(app.exec())

if __name__ == "__main__":
    main()