import sys
from PyQt6.QtWidgets import *
from PyQt6.QtGui import *
from PyQt6.QtCharts import *
from PyQt6.QtCore import *
import xlsxwriter
import random
import os
import datetime
import serial.tools.list_ports






# Crée une application et une fenêtre avec le titre "Projet". La fenêtre est maximisée et a un fond noir avec du texte blanc.

# La première ligne de code crée une application. Cette application est nécessaire pour exécuter le code.
app = QApplication(sys.argv)

# La deuxième ligne de code crée une fenêtre. Cette fenêtre est la fenêtre principale de l'application.
window = QWidget()

# La troisième ligne de code définit le titre de la fenêtre.
window.setWindowTitle("Projet")

# La quatrième ligne de code maximise la fenêtre.
window.showMaximized()

# La cinquième ligne de code définit le style de la fenêtre. Le style de la fenêtre est défini en utilisant la palette de couleurs `ChartThemeDark`. Cette palette de couleurs a un fond noir et du texte blanc.
window.setStyleSheet("background-color: ChartThemeDark; color: white;")




# Crée une palette de couleurs et définit les couleurs pour différentes parties de l'interface utilisateur.

# La première ligne de code crée une palette de couleurs.
palette = QPalette()

# Les lignes suivantes de code définissent les couleurs pour différentes parties de l'interface utilisateur.
palette.setColor(QPalette.ColorRole.Window, QColor(53, 53, 53))  # Fond de la fenêtre
palette.setColor(QPalette.ColorRole.WindowText, QColor(250, 250, 250))  # Texte de la fenêtre
palette.setColor(QPalette.ColorRole.Base, QColor(25, 25, 25))  # Fond des widgets
palette.setColor(QPalette.ColorRole.AlternateBase, QColor(53, 53, 53))  # Fond des widgets alternatifs
palette.setColor(QPalette.ColorRole.ToolTipBase, QColor(250, 250, 250))  # Fond des info-bulles
palette.setColor(QPalette.ColorRole.ToolTipText, QColor(250, 250, 250))  # Texte des info-bulles
palette.setColor(QPalette.ColorRole.Text, QColor(250, 250, 250))  # Texte des widgets
palette.setColor(QPalette.ColorRole.Button, QColor(53, 53, 53))  # Fond des boutons
palette.setColor(QPalette.ColorRole.ButtonText, QColor(250, 250, 250))  # Texte des boutons
palette.setColor(QPalette.ColorRole.Highlight, QColor(42, 110, 177))  # Couleur de surbrillance
palette.setColor(QPalette.ColorRole.HighlightedText, QColor(250, 250, 250))  # Texte en surbrillance



# Crée un graphique et lui donne le thème `ChartThemeDark`.

# La première ligne de code crée un graphique.
chart = QChart()

# La deuxième ligne de code définit le thème du graphique en utilisant la valeur `ChartThemeDark`. Ce thème donne au graphique un fond noir et du texte blanc.
chart.setTheme(QChart.ChartTheme.ChartThemeDark)

# Crée 3 séries pour le graphique, une pour chaque axe. Les séries sont nommées "Axe X", "Axe Y" et "Axe Z".
series1 = QLineSeries()
series1.setName("Axe X")
series2 = QLineSeries()
series2.setName("Axe Y")
series3 = QLineSeries()
series3.setName("Axe Z")


# Crée un affichage pour le graphique et lui dit d'utiliser l'anticrénelage pour améliorer l'apparence.
chartview = QChartView(chart)
chartview.setRenderHint(QPainter.RenderHint.Antialiasing)





# Crée trois zones de texte pour afficher les informations des axes. Les zones de texte sont mises en lecture seule et ont un fond noir avec du texte blanc.

# La première ligne de code crée une zone de texte pour afficher les informations de l'axe X.
axis_x_box = QTextEdit()

# La deuxième ligne de code définit la zone de texte comme étant en lecture seule.
axis_x_box.setReadOnly(True)

# La troisième ligne de code définit le style de la zone de texte. Le style de la zone de texte est défini en utilisant la palette de couleurs `ChartThemeDark`. Ce thème donne à la zone de texte un fond noir et du texte blanc.
axis_x_box.setStyleSheet("background-color: ChartThemeDark; color: white; font-size: 14px;")

# La quatrième ligne de code définit la largeur fixe de la zone de texte.
axis_x_box.setFixedWidth(150)

# Les lignes suivantes de code sont identiques pour les zones de texte des axes Y et Z.
axis_y_box = QTextEdit()
axis_y_box.setReadOnly(True)
axis_y_box.setStyleSheet("background-color: ChartThemeDark; color: white; font-size: 14px;")
axis_y_box.setFixedWidth(150)

axis_z_box = QTextEdit()
axis_z_box.setReadOnly(True)
axis_z_box.setStyleSheet("background-color: ChartThemeDark; color: white; font-size: 14px;")
axis_z_box.setFixedWidth(150)


# Ajoute les trois zones de texte dans le layout horizontal.
h_layout.addLayout(axis_x, 1)
h_layout.addLayout(axis_y, 1)
h_layout.addLayout(axis_z, 1)

# Ajoute le layout horizontal au layout vertical.
layout.addLayout(h_layout)

# Crée une zone de liste déroulante pour sélectionner le port COM.
# Création de la zone de liste déroulante
OM1 = QComboBox(window)
valOM1 = OM1.currentText()
OM1.addItem("Selectionnez le port")

# Récupération des ports COM disponibles
ports = serial.tools.list_ports.comports()

# Création d'une liste des ports COM disponibles
OM1_list = []
for p,d,h in sorted(ports):
    OM1_list.append(p)

# Affichage des ports COM disponibles dans la zone de liste déroulante
for port in OM1_list:
    OM1.addItem(port)

# Mise en style de la zone de liste déroulante
OM1.setStyleSheet("background-color: ChartThemeDark; color: white font-size: 18px;")

# Ajout de la zone de liste déroulante au layout vertical
layout.addWidget(OM1)



# Creation du bouton commencer
button = QPushButton("Start", window)
button.setStyleSheet("background-color: green; color: white;")
button.setPalette(palette)
layout.addWidget(button)

# Creation du bouton stop
button_stop = QPushButton("Stop", window)
button_stop.setStyleSheet("background-color: green; color: white;")
button_stop.setPalette(palette)
# layout.addWidget(button_stop) # je lai pas ajouter au layer encore mais il est la 

# Creation du bouton effacer 
clear_button = QPushButton("Effacer", window)
clear_button.setStyleSheet("background-color: ChartThemeDark; color: white;")
clear_button.setPalette(palette)
layout.addWidget(clear_button)

# Creation de la zone de texte pour la sauvegarde
save_text = QLineEdit(window)
save_text.setPlaceholderText("Nom de fichier...")
save_text.setStyleSheet("background-color: ChartThemeDark; color: white;")

# Creation du bouton sauvegarder
save_button = QPushButton("Enregistrer", window)
save_button.setStyleSheet("background-color: green; color: white;")
save_button.setPalette(palette)



# Ajout du bouton et de la zone de texte
save_layout = QHBoxLayout()
save_layout.addWidget(save_text)
save_layout.addWidget(save_button)
layout.addLayout(save_layout)


