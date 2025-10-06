# -*- coding: utf-8 -*-

################################################################################
## Form generated from reading UI file 'hospeda.ui'
##
## Created by: Qt User Interface Compiler version 5.15.2
##
## WARNING! All changes made in this file will be lost when recompiling UI file!
################################################################################

# `QtCore`, `QtGui`, `QtWidgets` (PySide2): classes base para construir a interface.


from PySide2.QtCore import *
from PySide2.QtGui import *
from PySide2.QtWidgets import *

# Ui_MainWindow: Classe gerada pelo Qt Designer (UIC). 
# Esta classe define a interface do usuário para o aplicativo de hospedagem.
# Ela inclui a configuração de layout, widgets e estilos.
class Ui_MainWindow(object):
    # Monta a interface (widgets, layouts, estilos, conexões).
    def setupUi(self, MainWindow):
        if not MainWindow.objectName():
            MainWindow.setObjectName(u"MainWindow")
        MainWindow.resize(667, 394)
        self.centralwidget = QWidget(MainWindow)
        self.centralwidget.setObjectName(u"centralwidget")
        self.centralwidget.setStyleSheet(u"background-color: rgb(255, 255, 255)")
        self.verticalLayout_2 = QVBoxLayout(self.centralwidget)
        self.verticalLayout_2.setObjectName(u"verticalLayout_2")
        self.frame_3 = QFrame(self.centralwidget)
        self.frame_3.setObjectName(u"frame_3")
        self.frame_3.setFrameShape(QFrame.StyledPanel)
        self.frame_3.setFrameShadow(QFrame.Raised)
        self.verticalLayout = QVBoxLayout(self.frame_3)
        self.verticalLayout.setObjectName(u"verticalLayout")
        self.label = QLabel(self.frame_3)
        self.label.setObjectName(u"label")
        font = QFont()
        font.setFamily(u"Bahnschrift Light")
        font.setPointSize(11)
        font.setBold(True)
        font.setUnderline(False)
        font.setWeight(75)
        font.setStrikeOut(False)
        self.label.setFont(font)

        self.verticalLayout.addWidget(self.label)


        self.verticalLayout_2.addWidget(self.frame_3)

        self.frame_4 = QFrame(self.centralwidget)
        self.frame_4.setObjectName(u"frame_4")
        self.frame_4.setStyleSheet(u"QPushButton{\n"
"	background-color: #550000;\n"
"	color: rgb(255, 255, 255);\n"
"	border-color: rgb(0, 0, 0);\n"
"	border-radius: 15px;\n"
"	border: 1px solid black;\n"
"\n"
"}\n"
"\n"
"QPushButton:hover{\n"
"	color: rgb(0, 0, 0);\n"
"	background-color: #ffff7f;\n"
"}\n"
"\n"
"QLineEdit{\n"
"	color: rgb(0, 0, 0);\n"
"	border-color: rgb(0, 0, 0);\n"
"	border-radius: 15px;\n"
"	border: 1px solid black;\n"
"\n"
"}\n"
"\n"
"\n"
"")
        self.frame_4.setFrameShape(QFrame.StyledPanel)
        self.frame_4.setFrameShadow(QFrame.Raised)
        self.horizontalLayout_2 = QHBoxLayout(self.frame_4)
        self.horizontalLayout_2.setObjectName(u"horizontalLayout_2")
        self.txt_path = QLineEdit(self.frame_4)
        self.txt_path.setObjectName(u"txt_path")
        self.txt_path.setMinimumSize(QSize(120, 80))
        font1 = QFont()
        font1.setPointSize(11)
        self.txt_path.setFont(font1)
        self.txt_path.setStyleSheet(u"")
        self.txt_path.setAlignment(Qt.AlignCenter)

        self.horizontalLayout_2.addWidget(self.txt_path)

        self.btn_abrir = QPushButton(self.frame_4)
        self.btn_abrir.setObjectName(u"btn_abrir")
        self.btn_abrir.setMinimumSize(QSize(120, 80))
        self.btn_abrir.setFont(font1)
        self.btn_abrir.setCursor(QCursor(Qt.PointingHandCursor))

        self.horizontalLayout_2.addWidget(self.btn_abrir)


        self.verticalLayout_2.addWidget(self.frame_4)

        self.frame = QFrame(self.centralwidget)
        self.frame.setObjectName(u"frame")
        self.frame.setFrameShape(QFrame.StyledPanel)
        self.frame.setFrameShadow(QFrame.Raised)
        self.horizontalLayout_3 = QHBoxLayout(self.frame)
        self.horizontalLayout_3.setObjectName(u"horizontalLayout_3")
        self.plainTextEdit = QPlainTextEdit(self.frame)
        self.plainTextEdit.setObjectName(u"plainTextEdit")
        self.plainTextEdit.setReadOnly(True)

        self.horizontalLayout_3.addWidget(self.plainTextEdit)


        self.verticalLayout_2.addWidget(self.frame)

        self.frame_5 = QFrame(self.centralwidget)
        self.frame_5.setObjectName(u"frame_5")
        self.frame_5.setFrameShape(QFrame.StyledPanel)
        self.frame_5.setFrameShadow(QFrame.Raised)
        self.horizontalLayout = QHBoxLayout(self.frame_5)
        self.horizontalLayout.setObjectName(u"horizontalLayout")
        self.btn_rc = QPushButton(self.frame_5)
        self.btn_rc.setObjectName(u"btn_rc")
        self.btn_rc.setMinimumSize(QSize(275, 150))
        font2 = QFont()
        font2.setPointSize(10)
        font2.setBold(True)
        font2.setWeight(75)
        self.btn_rc.setFont(font2)
        self.btn_rc.setCursor(QCursor(Qt.PointingHandCursor))
        self.btn_rc.setStyleSheet(u"QPushButton{\n"
"	background-color: #3498db;\n"
"	color: rgb(255, 255, 255);\n"
"	border-color: rgb(0, 0, 0);\n"
"	border-radius: 15px;\n"
"	border: 2px solid black;\n"
"\n"
"}\n"
"\n"
"QPushButton:hover{\n"
"	color: rgb(0, 0, 0);\n"
"	background-color: #ffff7f;\n"
"}")

        self.horizontalLayout.addWidget(self.btn_rc)

        self.btn_pc = QPushButton(self.frame_5)
        self.btn_pc.setObjectName(u"btn_pc")
        self.btn_pc.setMinimumSize(QSize(275, 150))
        self.btn_pc.setFont(font2)
        self.btn_pc.setCursor(QCursor(Qt.PointingHandCursor))
        self.btn_pc.setStyleSheet(u"QPushButton{\n"
"	background-color: #2ecc71;\n"
"	color: rgb(255, 255, 255);\n"
"	border-color: rgb(0, 0, 0);\n"
"	border-radius: 15px;\n"
"	border: 2px solid black;\n"
"\n"
"}\n"
"\n"
"QPushButton:hover{\n"
"	color: rgb(0, 0, 0);\n"
"	background-color: #ffff7f;\n"
"}")

        self.horizontalLayout.addWidget(self.btn_pc)

        self.btn_frs = QPushButton(self.frame_5)
        self.btn_frs.setObjectName(u"btn_frs")
        self.btn_frs.setMinimumSize(QSize(275, 150))
        self.btn_frs.setFont(font2)
        self.btn_frs.setCursor(QCursor(Qt.PointingHandCursor))
        self.btn_frs.setStyleSheet(u"QPushButton{\n"
"	background-color: #f39c12;\n"
"	color: rgb(255, 255, 255);\n"
"	border-color: rgb(0, 0, 0);\n"
"	border-radius: 15px;\n"
"	border: 2px solid black;\n"
"\n"
"}\n"
"\n"
"QPushButton:hover{\n"
"	color: rgb(0, 0, 0);\n"
"	background-color: #ffff7f;\n"
"}")

        self.horizontalLayout.addWidget(self.btn_frs)

        self.btn_gdf = QPushButton(self.frame_5)
        self.btn_gdf.setObjectName(u"btn_gdf")
        self.btn_gdf.setMinimumSize(QSize(275, 150))
        self.btn_gdf.setFont(font2)
        self.btn_gdf.setCursor(QCursor(Qt.PointingHandCursor))
        self.btn_gdf.setStyleSheet(u"QPushButton{\n"
"	background-color: #9b59b6;\n"
"	color: rgb(255, 255, 255);\n"
"	border-color: rgb(0, 0, 0);\n"
"	border-radius: 15px;\n"
"	border: 2px solid black;\n"
"\n"
"}\n"
"\n"
"QPushButton:hover{\n"
"	color: rgb(0, 0, 0);\n"
"	background-color: #ffff7f;\n"
"}")

        self.horizontalLayout.addWidget(self.btn_gdf)

        self.btn_senha = QPushButton(self.frame_5)
        self.btn_senha.setObjectName(u"btn_senha")
        self.btn_senha.setMinimumSize(QSize(275, 150))
        self.btn_senha.setFont(font2)
        self.btn_senha.setCursor(QCursor(Qt.PointingHandCursor))
        self.btn_senha.setStyleSheet(u"QPushButton{\n"
"	background-color: #e74c3c;\n"
"	color: rgb(255, 255, 255);\n"
"	border-color: rgb(0, 0, 0);\n"
"	border-radius: 15px;\n"
"	border: 2px solid black;\n"
"\n"
"}\n"
"\n"
"QPushButton:hover{\n"
"	color: rgb(0, 0, 0);\n"
"	background-color: #ffff7f;\n"
"}")

        self.horizontalLayout.addWidget(self.btn_senha)


        self.verticalLayout_2.addWidget(self.frame_5, 0, Qt.AlignHCenter)

        MainWindow.setCentralWidget(self.centralwidget)

        self.retranslateUi(MainWindow)

        QMetaObject.connectSlotsByName(MainWindow)
    
    # Define/atualiza os textos traduzíveis da interface (títulos e rótulos).
    # Centraliza **textos e traduções** dos componentes (título, rótulos dos botões, *placeholder* no `txt_path`).
    # Facilita a **localização** sem reescrever `setupUi`.

    def retranslateUi(self, MainWindow):
        MainWindow.setWindowTitle(QCoreApplication.translate("MainWindow", u"MainWindow", None))
        self.label.setText(QCoreApplication.translate("MainWindow", u"<html><head/><body><p align=\"center\"><span style=\" font-size:30pt; font-weight:600;\">Pagamento de despesas de hospedagem</span></p></body></html>", None))
        self.txt_path.setPlaceholderText(QCoreApplication.translate("MainWindow", u"Planilha com dados ---- >", None))
        self.btn_abrir.setText(QCoreApplication.translate("MainWindow", u"Selecione", None))
        self.btn_rc.setText(QCoreApplication.translate("MainWindow", u"Requisi\u00e7\u00e3o", None))
        self.btn_pc.setText(QCoreApplication.translate("MainWindow", u"Pedido", None))
        self.btn_frs.setText(QCoreApplication.translate("MainWindow", u"Registro de Servi\u00e7o", None))
        self.btn_gdf.setText(QCoreApplication.translate("MainWindow", u"Gest\u00e3o de Documentos", None))
        self.btn_senha.setText(QCoreApplication.translate("MainWindow", u"Senha", None))
    # retranslateUi

