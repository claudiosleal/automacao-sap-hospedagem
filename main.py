# -*- coding: utf-8 -*-
"""
main.py

Unifica as funcionalidades de:
- Pedido (btn_pc)  -> mm._pedido()
- Requisição (btn_rc) -> mm._requisicao()
- Registro de Serviço (btn_frs) -> mm._frs()
- Gestão de Documentos (btn_gd) -> mm._gd()
- Gerenciador de Senhas (btn_senha) -> keyring (cadastro/consulta)

Requisitos:
  pip install PySide2 keyring
  (Além dos seus módulos locais: ui_main.py e mm.servicos)

A UI é a mesma Ui_MainWindow fornecida (com btn_abrir, btn_pc, btn_rc, btn_frs, btn_gd, btn_senha,
plainTextEdit e txt_path).
"""

import sys
import os
from datetime import datetime

try:
    import keyring  # usado apenas no diálogo de senhas
except Exception:
    keyring = None  # permite abrir a UI mesmo sem keyring instalado

from PySide2.QtWidgets import (
    QApplication, QMainWindow, QMessageBox, QFileDialog, QDialog,
    QVBoxLayout, QFormLayout, QLineEdit, QPushButton
)

from ui_main import Ui_MainWindow
from mm.servicos import mm


# ------------------------------------------------------------------
# Redirecionador de stdout para o QPlainTextEdit
# ------------------------------------------------------------------
# EmissorDeLog: Redireciona tudo que é impresso (stdout) 
# para um widget QPlainTextEdit, permitindo visualizar logs e mensagens
# diretamente na interface gráfica.


class EmissorDeLog:
    def __init__(self, widget):
        self.widget = widget

    def write(self, mensagem):
        try:
            texto = (mensagem or "").rstrip("\n")
            if texto:
                self.widget.appendPlainText(texto)
        except Exception:
            # Em último caso, evita travar se o widget ainda não estiver pronto
            pass

    def flush(self):
        pass


# ------------------------------------------------------------------
# Diálogo de gerenciamento de senhas
# ------------------------------------------------------------------
# PasswordDialog: Diálogo para cadastrar/consultar senhas usando o keyring
# local (usuário/sistema/senha)

class PasswordDialog(QDialog):
    def __init__(self, parent=None):
        super(PasswordDialog, self).__init__(parent)
        self.setWindowTitle("Gerenciador de Senhas")
        self.setMinimumSize(360, 220)

        layout = QVBoxLayout(self)
        form_layout = QFormLayout()

        self.user_input = QLineEdit()
        self.system_input = QLineEdit()
        self.password_input = QLineEdit()
        self.password_input.setEchoMode(QLineEdit.Password)

        form_layout.addRow("Usuário:", self.user_input)
        form_layout.addRow("Sistema:", self.system_input)
        form_layout.addRow("Senha:", self.password_input)

        layout.addLayout(form_layout)

        self.save_button = QPushButton("Cadastrar Senha")
        self.retrieve_button = QPushButton("Consultar Senha")
        layout.addWidget(self.save_button)
        layout.addWidget(self.retrieve_button)

        self.save_button.clicked.connect(self.save_password)
        self.retrieve_button.clicked.connect(self.retrieve_password)

    def _require_keyring(self) -> bool:
        if keyring is None:
            QMessageBox.warning(
                self,
                "Dependência ausente",
                "O módulo 'keyring' não está instalado. Execute:\n  pip install keyring"
            )
            return False
        return True

    def get_user(self):
        return self.user_input.text().strip()

    def save_password(self):
        if not self._require_keyring():
            return
        user = self.user_input.text().strip()
        system = self.system_input.text().strip()
        password = self.password_input.text()
        if not (user and system and password):
            QMessageBox.warning(self, "Erro", "Todos os campos devem ser preenchidos.")
            return
        try:
            keyring.set_password(system, user, password)
            QMessageBox.information(
                self, "Sucesso",
                f"Senha para o sistema '{system}' e usuário '{user}' cadastrada com sucesso!"
            )
            self.accept()  # <<< CORRIGIDO
        except Exception as e:
            self.accept()
            QMessageBox.critical(self, "Erro", f"Não foi possível cadastrar a senha:\n{e}")

    def retrieve_password(self):
        if not self._require_keyring():
            return
        user = self.user_input.text().strip()
        system = self.system_input.text().strip()
        if not (user and system):
            QMessageBox.warning(self, "Erro", "Campos 'Usuário' e 'Sistema' devem ser preenchidos.")
            return
        try:
            password = keyring.get_password(system, user)
            if password:
                QMessageBox.information(
                    self, "Sucesso",
                    f"A senha para o usuário '{user}' e sistema '{system}' é: {password}"
                )
                self.accept()  # <<< CORRIGIDO
            else:
                QMessageBox.warning(
                    self, "Aviso",
                    f"Nenhuma senha encontrada para o usuário '{user}' e sistema '{system}'."
                )
        except Exception as e:
            self.accept()
            QMessageBox.critical(self, "Erro", f"Não foi possível consultar a senha:\n{e}")


# ------------------------------------------------------------------
# Janela principal unificada
# ------------------------------------------------------------------
# MainWindow:  inicializa UI, conecta botões
# aos fluxos e gerencia parâmetros SAP.


class MainWindow(QMainWindow, Ui_MainWindow):

    def _salvar_usuario_sap(self, user):
        caminho_excel = self.txt_path.text().strip()

        # Se o caminho não for um arquivo válido, avise o usuário e pare a execução
        if not caminho_excel or not os.path.isfile(caminho_excel):
            print("AVISO: Usuário SAP não foi salvo. Selecione um arquivo Excel primeiro.")
            QMessageBox.warning(
                self,
                "Usuário não salvo em arquivo",
                f"O usuário SAP '{user}' foi definido para esta sessão, mas não pôde ser salvo em 'sap_user.txt'.\n\n"
                "Por favor, selecione um arquivo Excel válido para que a configuração seja salva na mesma pasta."
            )
            return

        pasta_excel = os.path.dirname(caminho_excel)
        caminho_arquivo = os.path.join(pasta_excel, "sap_user.txt")
        try:
            with open(caminho_arquivo, "w", encoding="utf-8") as f:
                f.write(user.strip())
            # Adiciona um feedback de sucesso no log
            print(f"Usuário SAP '{user}' salvo com sucesso em: {caminho_arquivo}")
        except Exception as e:
            print(f"Erro ao salvar usuário SAP: {e}")
            QMessageBox.critical(self, "Erro ao Salvar", f"Não foi possível criar o arquivo 'sap_user.txt'.\n\nErro: {e}")


    def _carregar_usuario_sap(self):
        caminho_excel = self.txt_path.text().strip()
        if caminho_excel and os.path.isfile(caminho_excel):
            pasta_excel = os.path.dirname(caminho_excel)
            caminho_arquivo = os.path.join(pasta_excel, "sap_user.txt")
            if os.path.exists(caminho_arquivo):
                try:
                    with open(caminho_arquivo, "r", encoding="utf-8") as f:
                        return f.read().strip()
                except Exception as e:
                    print(f"Erro ao carregar usuário SAP: {e}")
        return "TTTT" # Retorna um padrão se nada for carregado

    def __init__(self):
            super(MainWindow, self).__init__()
            self.setupUi(self)
            self.setWindowTitle("Sistema Gestor de Hospedagem")
            # Redireciona prints para o QPlainTextEdit
            self._stdout_original = sys.stdout
            sys.stdout = EmissorDeLog(self.plainTextEdit)

            # Ligações dos botões (mantendo a mesma UI)
            self.btn_abrir.clicked.connect(self.open_file)
            self.btn_rc.clicked.connect(self.process_requisicao)
            self.btn_pc.clicked.connect(self.process_pedido)
            self.btn_frs.clicked.connect(self.process_frs)
            self.btn_gd.clicked.connect(self.process_gd)
            self.btn_senha.clicked.connect(self.open_password_dialog)

            # Texto padrão (a UI já tem placeholder, mas deixo um valor inicial visível)
            if not self.txt_path.text().strip():
                self.txt_path.setText("Planilha com dados ---- >")

            # Configurações padrão do SAP (ajuste para o seu ambiente)
            self._sap_user = self._carregar_usuario_sap()
            self._sap_environment = "F04 - SAP Scripting Produção"
            self._sap_logon_path = r"C:\\Program Files (x86)\\SAP\\FrontEnd\\SAPgui\\saplogon.exe"

    # ----------------------- UI -----------------------
    def closeEvent(self, event):
        try:
            sys.stdout = self._stdout_original
        except Exception:
            pass
        super().closeEvent(event)

    def open_password_dialog(self):
        # Abre a caixa de diálogo de gerenciamento de senhas
        dlg = PasswordDialog(self)
        if dlg.exec_() == QDialog.Accepted:
            user = dlg.get_user()
            if user:
                self._sap_user = user
                self._salvar_usuario_sap(user)
            print(f"Usuário SAP atualizado para: {self._sap_user}")

    def open_file(self):
        # Abre uma caixa de diálogo para selecionar a planilha Excel com os dados
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Carregar a planilha com os dados",
            "",
            "Planilhas Excel (*.xlsx *.xls);;Todos os arquivos (*.*)"
        )
        if file_path:
            self.txt_path.setText(file_path)
            # Ao selecionar um novo arquivo, tenta carregar o usuário SAP daquela pasta
            self._sap_user = self._carregar_usuario_sap()
            print(f"Usuário SAP definido como '{self._sap_user}' (carregado da pasta do Excel ou padrão).")
        else:
            QMessageBox.information(self, "Aviso", "Nenhum arquivo selecionado.")

    def _validar_caminho_excel(self) -> str:
        # Valida se o caminho informado aponta para um arquivo Excel existente.
        caminho = self.txt_path.text().strip()
        if not caminho or not os.path.isfile(caminho):
            QMessageBox.warning(self, "Erro", "Por favor, selecione um arquivo Excel válido.")
            return ""
        return caminho

    # ----------------------- Conexão ao SAP ----------------------
    def _conectar_sap(self):
        # Cria a instância da automação (mm) e estabelece a sessão no SAP
        # Se a conexão falhar, exibe um aviso e retorna None
        print("Iniciando conexão com o SAP...")
        automacao_sap = mm(
            sap_user=self._sap_user,
            sap_environment=self._sap_environment,
            sap_logon_path=self._sap_logon_path,
        )
        sessao_ativa = automacao_sap._conecta()
        if sessao_ativa:
            print("Conexão com o SAP estabelecida com sucesso.")
            return automacao_sap
        else:
            QMessageBox.warning(
                self,
                "Erro",
                "Não foi possível conectar ao SAP. Verifique as configurações ou se o SAP GUI está instalado."
            )
            return None

    # ----------------------- Fluxos ---------------------------
    def process_requisicao(self):
        # Fluxo da Requisição: lê dados e chama mm._requisicao()
        # O caminho do Excel é validado e o SAP é conectado antes de iniciar a requisição
        # Se o caminho não for válido, exibe um aviso e não executa a requisição
        # A requisição é executada com os dados lidos do Excel e o caminho do Excel
        # O tempo total de execução é impresso no console
        # Se ocorrer algum erro, exibe uma mensagem de erro
        # e não executa a requisição
        # A requisição é finalizada com uma mensagem de sucesso
        # e o tempo total de execução é impresso no console
        caminho_excel = self._validar_caminho_excel()
        if not caminho_excel:
            return
        self.plainTextEdit.clear()
        inicio = datetime.now()

        automacao_sap = self._conectar_sap()
        if not automacao_sap:
            return
        try:
            print("Preparando dados para a requisição...")
            lista = automacao_sap._relatorio(caminho_excel)
            print("Iniciando a automação da requisição...")
            automacao_sap._requisicao(lista, caminho_excel)
            print("Processo de requisição finalizado.")
        except Exception as e:
            QMessageBox.critical(self, "Erro na requisição", f"Falha no processamento:\n{e}")
        finally:
            print(f"Tempo total de execução (requisição): {datetime.now() - inicio}")

    def process_pedido(self):
        # Fluxo do Pedido: lê dados e chama mm._pedido()
        # O caminho do Excel é validado e o SAP é conectado antes de iniciar o pedido
        # Se o caminho não for válido, exibe um aviso e não executa o pedido
        # O pedido é executado com os dados lidos do Excel e o caminho do Excel
        # O tempo total de execução é impresso no console
        # Se ocorrer algum erro, exibe uma mensagem de erro
        # e não executa o pedido
        # O pedido é finalizado com uma mensagem de sucesso
        # e o tempo total de execução é impresso no console
        caminho_excel = self._validar_caminho_excel()
        if not caminho_excel:
            return
        self.plainTextEdit.clear()
        inicio = datetime.now()

        automacao_sap = self._conectar_sap()
        if not automacao_sap:
            return
        try:
            print("Preparando dados para o pedido...")
            lista = automacao_sap._relatorio(caminho_excel)
            print("Iniciando a automação do pedido...")
            automacao_sap._pedido(lista, caminho_excel)
            print("Processo de pedido finalizado.")
        except Exception as e:
            QMessageBox.critical(self, "Erro no pedido", f"Falha no processamento:\n{e}")
        finally:
            print(f"Tempo total de execução (pedido): {datetime.now() - inicio}")

    def process_frs(self):
        # Fluxo do Registro de Serviço (FRS): lê dados e chama mm._frs()
        # O caminho do Excel é validado e o SAP é conectado antes de iniciar o FRS
        # Se o caminho não for válido, exibe um aviso e não executa o FRS
        # O FRS é executado com os dados lidos do Excel e o caminho do Excel
        # O tempo total de execução é impresso no console
        # Se ocorrer algum erro, exibe uma mensagem de erro
        # e não executa o FRS
        # O FRS é finalizado com uma mensagem de sucesso
        # e o tempo total de execução é impresso no console
        caminho_excel = self._validar_caminho_excel()
        if not caminho_excel:
            return
        self.plainTextEdit.clear()
        inicio = datetime.now()

        automacao_sap = self._conectar_sap()
        if not automacao_sap:
            return
        try:
            print("Preparando dados para Registro de Serviço (FRS)...")
            lista = automacao_sap._relatorio(caminho_excel)
            print("Iniciando a automação do FRS...")
            automacao_sap._frs(lista, caminho_excel)
            print("Processo de FRS finalizado.")
        except Exception as e:
            QMessageBox.critical(self, "Erro no FRS", f"Falha no processamento:\n{e}")
        finally:
            print(f"Tempo total de execução (FRS): {datetime.now() - inicio}")

    def process_gd(self):
        # Fluxo da Gestão de Documentos (GD): lê dados e chama mm._gd()
        # O caminho do Excel é validado e o SAP é conectado antes de iniciar o GD
        # Se o caminho não for válido, exibe um aviso e não executa o GD
        # O GD é executado com os dados lidos do Excel e o caminho do Excel
        # O tempo total de execução é impresso no console
        # Se ocorrer algum erro, exibe uma mensagem de erro
        # e não executa o GD
        # O GD é finalizado com uma mensagem de sucesso
        # e o tempo total de execução é impresso no console

        caminho_excel = self._validar_caminho_excel()
        if not caminho_excel:
            return
        self.plainTextEdit.clear()
        inicio = datetime.now()

        automacao_sap = self._conectar_sap()
        if not automacao_sap:
            return
        try:
            print("Preparando dados para Gestão de Documentos (GD)...")
            lista = automacao_sap._relatorio(caminho_excel)
            print("Iniciando a automação do GD...")
            automacao_sap._gd(lista, caminho_excel)
            print("Processo de GD finalizado.")
        except Exception as e:
            QMessageBox.critical(self, "Erro no GD", f"Falha no processamento:\n{e}")
        finally:
            print(f"Tempo total de execução (GD): {datetime.now() - inicio}")


# ------------------------------------------------------------------
# Inicialização da aplicação
# ------------------------------------------------------------------
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())