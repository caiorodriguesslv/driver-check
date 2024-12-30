import sys
from PyQt6.QtWidgets import QApplication, QMainWindow, QPushButton, QVBoxLayout, QWidget, QTextEdit
from scripts.driver_list import list_intel_drivers  # Importa a função que já existe
import json

class DriverCheckApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Driver Check - Intel")
        self.setGeometry(300, 300, 600, 400)
        self.initUI()

    def initUI(self):
        # Cria o layout principal
        layout = QVBoxLayout()

        # Botão para iniciar a verificação
        self.start_button = QPushButton("Verificar Drivers Intel")
        self.start_button.clicked.connect(self.run_driver_check)
        layout.addWidget(self.start_button)

        # Área de texto para exibir resultados
        self.result_area = QTextEdit()
        self.result_area.setReadOnly(True)
        layout.addWidget(self.result_area)

        # Container para o layout
        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

    def run_driver_check(self):
        # Chama a função de verificação
        drivers_found, quantity_drivers = list_intel_drivers()
        parsed_json = json.loads(drivers_found)

        # Formata a saída para exibir na área de texto
        display_text = json.dumps(parsed_json, indent=4, ensure_ascii=False)
        self.result_area.setText(display_text)

        # Mensagem no console
        print("Drivers encontrados:", quantity_drivers)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = DriverCheckApp()
    window.show()
    sys.exit(app.exec())
