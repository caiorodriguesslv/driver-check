import sys
from PyQt6.QtWidgets import (QApplication, QMainWindow, QPushButton, QVBoxLayout, QWidget,
                             QTextEdit, QLabel, QProgressBar, QHBoxLayout)
from PyQt6.QtCore import Qt, QTimer
from scripts.driver_list import list_intel_drivers
import json

class DriverCheckApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Driver Check - Intel")
        self.setGeometry(300, 300, 700, 500)
        self.initUI()

    def initUI(self):
        # Cria o layout principal
        main_layout = QVBoxLayout()

        # Título
        self.title_label = QLabel("Verificação de Drivers Intel")
        self.title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.title_label.setStyleSheet("font-size: 20px; font-weight: bold;")
        main_layout.addWidget(self.title_label)

        # Barra de progresso
        self.progress_bar = QProgressBar()
        self.progress_bar.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.progress_bar.setValue(0)
        self.progress_bar.setVisible(False)
        main_layout.addWidget(self.progress_bar)

        # Botão para iniciar a verificação
        self.start_button = QPushButton("Verificar Drivers Intel")
        self.start_button.setStyleSheet("font-size: 16px; padding: 10px;")
        self.start_button.clicked.connect(self.run_driver_check)
        main_layout.addWidget(self.start_button)

        # Área de texto para exibir resultados
        self.result_area = QTextEdit()
        self.result_area.setReadOnly(True)
        self.result_area.setStyleSheet("font-family: Courier; font-size: 14px;")
        main_layout.addWidget(self.result_area)

        # Status de execução
        self.status_label = QLabel("Pronto para iniciar.")
        self.status_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.status_label.setStyleSheet("font-size: 14px; color: green;")
        main_layout.addWidget(self.status_label)

        # Container para o layout
        container = QWidget()
        container.setLayout(main_layout)
        self.setCentralWidget(container)

    def run_driver_check(self):
        # Atualiza o status e exibe a barra de progresso
        self.status_label.setText("Verificando drivers... Por favor, aguarde.")
        self.status_label.setStyleSheet("font-size: 14px; color: orange;")
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(30)

        # Simula atraso de execução
        QTimer.singleShot(2000, self.execute_driver_check)

    def execute_driver_check(self):
        # Executa a verificação
        drivers_found, quantity_drivers = list_intel_drivers()
        parsed_json = json.loads(drivers_found)

        # Formata a saída para exibir na área de texto
        display_text = json.dumps(parsed_json, indent=4, ensure_ascii=False)
        self.result_area.setText(display_text)

        # Atualiza o status e barra de progresso
        self.progress_bar.setValue(100)
        self.progress_bar.setVisible(False)
        self.status_label.setText(f"Verificação concluída! {quantity_drivers} drivers encontrados.")
        self.status_label.setStyleSheet("font-size: 14px; color: green;")
        print("Drivers encontrados:", quantity_drivers)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = DriverCheckApp()
    window.show()
    sys.exit(app.exec())
