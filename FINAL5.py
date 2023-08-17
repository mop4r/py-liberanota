import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QVBoxLayout, QWidget, QLabel, QLineEdit, QPushButton, QMessageBox, QComboBox
from PyQt5.QtCore import Qt
import pyodbc

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        # lista de ODBC, sorted alphabetically
        odbc_list = sorted(pyodbc.dataSources())

        # Criação dos widgets
        self.odbc_label = QLabel('Selecione o ODBC')
        self.odbc_combobox = QComboBox()
        self.odbc_combobox.addItems(odbc_list)
        self.connect_button = QPushButton('Conectar')
        self.connect_button.clicked.connect(self.connect_to_database)
        self.num_nota_label = QLabel('Número da nota fiscal:')
        self.num_nota_input = QLineEdit()
        self.num_nota_input.setDisabled(True)
        self.num_nota_input.returnPressed.connect(self.focus_data_input)
        self.data_nota_label = QLabel('Data da nota fiscal (DDMMAAAA):')
        self.data_nota_input = QLineEdit()
        self.data_nota_input.setDisabled(True)
        self.data_nota_input.returnPressed.connect(self.submit_query)
        self.submit_button = QPushButton('Submit')
        self.submit_button.setDisabled(True)
        self.submit_button.clicked.connect(self.submit_query)
        self.cancel_button = QPushButton('Cancel')
        self.cancel_button.clicked.connect(self.close)

        # Layout da janela principal
        layout = QVBoxLayout()
        layout.addWidget(self.odbc_label)
        layout.addWidget(self.odbc_combobox)
        layout.addWidget(self.connect_button)
        layout.addWidget(self.num_nota_label)
        layout.addWidget(self.num_nota_input)
        layout.addWidget(self.data_nota_label)
        layout.addWidget(self.data_nota_input)
        layout.addWidget(self.submit_button)
        layout.addWidget(self.cancel_button)

        # Widget principal
        central_widget = QWidget()
        central_widget.setLayout(layout)
        self.setCentralWidget(central_widget)

    def connect_to_database(self):
        selected_odbc = self.odbc_combobox.currentText()
        try:
            conn = pyodbc.connect('DSN=' + selected_odbc)
            QMessageBox.information(self, 'Conexão', 'Conexão estabelecida com sucesso!')
            self.connect_button.setDisabled(True)
            self.odbc_combobox.setDisabled(True)
        except Exception as e:
            QMessageBox.critical(self, 'Erro de conexão', f'Erro ao conectar ao banco de dados: {str(e)}')
            conn = None
            self.connect_button.setDisabled(False)

        if conn is not None:
            self.num_nota_input.setDisabled(False)
            self.data_nota_input.setDisabled(False)
            self.submit_button.setDisabled(False)

    def focus_data_input(self):
        self.data_nota_input.setFocus()

    def submit_query(self):
        selected_odbc = self.odbc_combobox.currentText()
        conn = pyodbc.connect('DSN=' + selected_odbc)
        try:
            cursor = conn.cursor()
            cursor.execute(
                "select getdatanormal(nsu004) as Data_nota, nsu003 as Série,nsu005 as Número, replace(replace(nsu006,2,'Saída'),1,'Entrada') as Tipo, nsu004, nsu006, empresa, nsu002, nsu011 from ges_359 where nsu004=(select getdatanumero(?)) and nsu005=?",
                (self.data_nota_input.text(), self.num_nota_input.text())
            )
            data = cursor.fetchone()
            if data is None:
                QMessageBox.critical(self, 'Nota fiscal não encontrada', 'Nota fiscal não encontrada!')
                return

            message = f"Data: {data[0]}\nSérie: {data[1]}\nNúmero: {data[2]}\nTipo: {data[3]}\nEmpresa: {data[6]}\nLocal: {data[7]}"

            confirm_dialog = QMessageBox.question(self, 'Confirmação', f"{message}\n\nDeseja corrigir a nota fiscal?",
                                                   QMessageBox.Yes | QMessageBox.No)
            if confirm_dialog == QMessageBox.No:
                return

            if data[8] == 'Autorizado o uso da NF-e':
                confirm_dialog = QMessageBox.question(self, 'Confirmação', "Essa nota já consta como 'Autorizado o uso da NF-e'. Deseja corrigir mesmo assim?",
                                                       QMessageBox.Yes | QMessageBox.No)
                if confirm_dialog == QMessageBox.No:
                    return

            cursor.execute(
                "update ges_359 set nsu007='', nsu008=88, nsu009='',nsu010='',nsu011='',nsu015='' where nsu004=? and nsu005=?",
                (data[4], data[2])
            )
            conn.commit()
            QMessageBox.information(self, 'Sucesso', 'Nota corrigida com sucesso! Valide pelo Gestor.')

        except Exception as e:
            QMessageBox.critical(self, 'Erro', f'Erro ao consultar banco de dados: {str(e)}')
        finally:
            conn.close()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.setWindowTitle('Consulta de nota fiscal')
    window.show()
    sys.exit(app.exec_())
