# Point d'entrée 211101
import sys
from gui.main_window import MainWindow
from PySide6.QtWidgets import (
    QProgressBar,
    QApplication,
    QHBoxLayout,
)
from PySide6.QtCore import Qt, Signal, QThread, QObject

def main():
    # Activer le support DPI haute résolution
    QApplication.setHighDpiScaleFactorRoundingPolicy(
        Qt.HighDpiScaleFactorRoundingPolicy.PassThrough
    )
    
    app = QApplication(sys.argv)
    app.setApplicationName("Réconciliation")
    app.setApplicationVersion("1.0.0")
    
    window = MainWindow()
    window.show()
    
    sys.exit(app.exec())

if __name__ == "__main__":
    main()