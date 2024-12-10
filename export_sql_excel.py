import sys
import mysql.connector
from PyQt5.QtWidgets import QApplication, QMainWindow, QLabel, QLineEdit, QPushButton, QMessageBox
import pandas as pd

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        # Definindo as dimensões da janela principal
        self.setGeometry(100, 100, 500, 400)
        self.setWindowTitle("Exportar para xlsx")

        # Label e campos para o banco de dados
        label_host = QLabel("Host:", self)
        label_host.move(10, 10)
        self.host_input = QLineEdit(self)
        self.host_input.move(140, 10)
        self.host_input.resize(200, 25)

        label_banco = QLabel("Nome do banco:", self)
        label_banco.move(10, 50)
        self.banco_input = QLineEdit(self)
        self.banco_input.move(140, 50)
        self.banco_input.resize(200, 25)

        label_login = QLabel("Login:", self)
        label_login.move(10, 90)
        self.login_input = QLineEdit(self)
        self.login_input.move(140, 90)
        self.login_input.resize(200, 25)

        label_senha = QLabel("Senha:", self)
        label_senha.move(10, 130)
        self.senha_input = QLineEdit(self)
        self.senha_input.move(140, 130)
        self.senha_input.resize(200, 25)
        self.senha_input.setEchoMode(QLineEdit.Password)

        # Campo para a instrução SQL
        label_sql = QLabel("Instrução SQL:", self)
        label_sql.move(10, 170)
        self.sql_input = QLineEdit(self)
        self.sql_input.move(10, 200)
        self.sql_input.resize(480, 25)

        # Botão para exportar
        exportar_btn = QPushButton("Exportar", self)
        exportar_btn.move(10, 240)
        exportar_btn.clicked.connect(self.exportar_sql)

    def exportar_sql(self):
        # Pegando os valores dos campos
        host = self.host_input.text()
        banco = self.banco_input.text()
        login = self.login_input.text()
        senha = self.senha_input.text()
        sql = self.sql_input.text()

        # Criando conexão com o banco de dados
        try:
            cnx = mysql.connector.connect(user=login, password=senha, host=host, database=banco, auth_plugin='mysql_native_password')
            cursor = cnx.cursor()
            cursor.execute(sql)
            resultado = cursor.fetchall()
            cursor.close()
            cnx.close()

            # Exportando para o Excel (xlsx)
            df = pd.DataFrame(resultado)            
            df.to_excel('{}.xlsx'.format(banco))            

            QMessageBox.information(self, "Sucesso", "Dados exportados com sucesso para '{}.xlsx'".format(banco))

        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Ocorreu um erro ao executar a instrução SQL:\n\n{str(e)}")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
