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


# Creation de l'application et de la fenetre
app = QApplication(sys.argv)
window = QWidget()
window.setWindowTitle("Projet")
window.showMaximized() 
window.setStyleSheet("background-color: ChartThemeDark; color: white;")



# Creation des couleurs 
palette = QPalette() 
palette.setColor(QPalette.ColorRole.Window, QColor(53, 53, 53)) 
palette.setColor(QPalette.ColorRole.WindowText, QColor(250, 250, 250)) 
palette.setColor(QPalette.ColorRole.Base, QColor(25, 25, 25)) 
palette.setColor(QPalette.ColorRole.AlternateBase, QColor(53, 53, 53)) 
palette.setColor(QPalette.ColorRole.ToolTipBase, QColor(250, 250, 250)) 
palette.setColor(QPalette.ColorRole.ToolTipText, QColor(250, 250, 250)) 
palette.setColor(QPalette.ColorRole.Text, QColor(250, 250, 250)) 
palette.setColor(QPalette.ColorRole.Button, QColor(53, 53, 53)) 
palette.setColor(QPalette.ColorRole.ButtonText, QColor(250, 250, 250)) 
palette.setColor(QPalette.ColorRole.Highlight, QColor(42, 110, 177)) 
palette.setColor(QPalette.ColorRole.HighlightedText, QColor(250, 250, 250))



# Creation du graphique
chart = QChart()
chart.setTheme(QChart.ChartTheme.ChartThemeDark)



# Creation de 3 series pour le graphique
series1 = QLineSeries()
series1.setName("Axe X")
series2 = QLineSeries()
series2.setName("Axe Y")
series3 = QLineSeries()
series3.setName("Axe Z")


# Affichage du graphique
chartview = QChartView(chart)
chartview.setRenderHint(QPainter.RenderHint.Antialiasing)



# Creation des infos autour du graphique
axis_x_box = QTextEdit()
axis_x_box.setReadOnly(True)
axis_x_box.setStyleSheet("background-color: ChartThemeDark; color: white; font-size: 14px;")
axis_x_box.setFixedWidth(150)

axis_y_box = QTextEdit()
axis_y_box.setReadOnly(True)
axis_y_box.setStyleSheet("background-color: ChartThemeDark; color: white; font-size: 14px;")
axis_y_box.setFixedWidth(150)

axis_z_box = QTextEdit()
axis_z_box.setReadOnly(True)
axis_z_box.setStyleSheet("background-color: ChartThemeDark; color: white; font-size: 14px;")
axis_z_box.setFixedWidth(150)



# Creation du layout vertical et horizontal 
layout = QVBoxLayout()
h_layout = QHBoxLayout()
h_layout.addWidget(chartview, 4)
spacer = QSpacerItem(40, 40, QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Minimum)
h_layout.addItem(spacer)



# Creation de 3 colonnes x, y et z 
axis_x = QVBoxLayout()
axis_x_title = QLabel("Axe X")
axis_x_title.setStyleSheet("background-color: ChartThemeDark; color: white font-size: 18px;")
axis_x.addWidget(axis_x_title)
axis_x.addWidget(axis_x_box)

axis_y = QVBoxLayout()
axis_y_title = QLabel("Axe Y")
axis_y_title.setStyleSheet("background-color: ChartThemeDark; color: white font-size: 18px;")
axis_y.addWidget(axis_y_title)
axis_y.addWidget(axis_y_box)

axis_z = QVBoxLayout()
axis_z_title = QLabel("Axe Z")
axis_z_title.setStyleSheet("background-color: ChartThemeDark; color: white font-size: 18px;")
axis_z.addWidget(axis_z_title)
axis_z.addWidget(axis_z_box)



# Ajout des 3 colonnes dans le layout horizontal
h_layout.addLayout(axis_x, 1)
h_layout.addLayout(axis_y, 1)
h_layout.addLayout(axis_z, 1)

# Ajout du layout horizontal dans le layout vertical 
layout.addLayout(h_layout)



# Creation du bouton com
OM1 = QComboBox(window)
valOM1 = OM1.currentText()
OM1.addItem("Selectionnez le port")
ports = serial.tools.list_ports.comports()
OM1_list = []
for p,d,h in sorted(ports):
    OM1_list.append(p)
OM1.setStyleSheet("background-color: ChartThemeDark; color: white font-size: 18px;")
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



# legende du graph
legend = chart.legend()
legend.setAlignment(Qt.AlignmentFlag.AlignRight)



# Axe X
axis_x = QValueAxis()
axis_x.setTitleText("Fréquence [Hz]")
axis_x.setLabelFormat("%.1f")
axis_x.setRange(0, 50)
axis_x.setTickCount(21)
axis_x.setGridLineVisible(True)
axis_x.setGridLineColor(Qt.GlobalColor.gray)
chart.addAxis(axis_x, Qt.AlignmentFlag.AlignBottom)

# Axe Y
axis_y = QValueAxis()
axis_y.setTitleText("Accéléromètre [g]")
axis_y.setLabelFormat("%.1f")
axis_y.setRange(-5, 5)
axis_y.setMinorTickCount(4)
axis_y.setGridLineVisible(True)
axis_y.setGridLineColor(Qt.GlobalColor.gray)
chart.addAxis(axis_y, Qt.AlignmentFlag.AlignLeft)



# Ajoute les series 1, 2 et 3 
chart.addSeries(series1)
chart.addSeries(series2)
chart.addSeries(series3)



# Création du timer
timer = QTimer()
timer.setInterval(1000)  # Timer qui s'exécute toutes les secondes
timer_label = QLabel("00:00:00")
timer_label.setAlignment(Qt.AlignmentFlag.AlignRight)

# Mise à jour du temps écoulé
elapsed_time = QTime(0, 0, 0)
def update_timer():
    global elapsed_time
    elapsed_time = elapsed_time.addSecs(1)
    timer_label.setText(elapsed_time.toString("hh:mm:ss"))

# Connexion du timer à la fonction d'actualisation
timer.timeout.connect(update_timer)

# Ajout du timer à l'interface utilisateur
timer_layout = QHBoxLayout()
timer_layout.addStretch(1)
timer_layout.addWidget(timer_label)
layout.addLayout(timer_layout)


# Fonction qui réinitialise le timer
def reset_timer():
    timer.stop()
    timer_label.setText("00:00:00")
    global elapsed_time
    elapsed_time = QTime(0, 0, 0) # Mettre à jour la variable elapsed_time en la remettant à 0




# Fonction qui affiche les points sur le graphique
def show_curve():
    series1.clear()
    series2.clear()
    series3.clear()

    for i in range(0, 71, 5):
        x = round(random.uniform(-5, 5), 3)
        y = round(random.uniform(-5, 5), 3)
        z = round(random.uniform(-5, 5), 3)
        series1.append(i, x)
        series2.append(i, y)
        series3.append(i, z)
    series1.attachAxis(axis_x)
    series1.attachAxis(axis_y)
    series2.attachAxis(axis_x)
    series2.attachAxis(axis_y)
    series3.attachAxis(axis_x)
    series3.attachAxis(axis_y)
    button.setEnabled(False)
    button.setStyleSheet("background-color: gray; color: white;")
    if not timer.isActive():  # Démarre le timer si ce n'est pas déjà fait
        elapsed_time = QTime(0, 0, 0)
        timer.start()


# Fonction qui enregistre l'image avec le nom entré
def save_image():
    # nom du dossier
    Nom_du_dossier = save_text.text()
    # Nom du chemin
    chemin = ''+Nom_du_dossier
    
    
    filename = save_text.text() + "_" + datetime.datetime.now().strftime("%d-%m-%Y") + ".jpg"
    chemin_vers_dossier = chemin + '/' + filename
    chartview.grab().save((chemin_vers_dossier), b"JPG")



def clear_clicked():
    # Afficher la fenêtre de confirmation
    reply = QMessageBox.question(window, 'Confirmation', 'Êtes-vous sûr de vouloir effacer les données du graphique ?', QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
    
    # Vérifier la réponse de l'utilisateur
    if reply == QMessageBox.StandardButton.Yes:
        # Effacer les données du graphique
        series1.clear()
        series2.clear()
        series3.clear()
        button.setEnabled(True)
        button.setStyleSheet("background-color: ChartThemeDark; color: white;")
        button.setStyleSheet("background-color: green; color: white;")
        reset_timer()
        update_info_x()
        update_info_y()
        update_info_z()


# Fonction qui permet d'afficher les informations lier au graphique dans les boxs x,y et z 
def update_info_x():
    axis_x_box.clear()
    axis_x_box.insertPlainText("")
    for i in range(series1.count()):
        axis_x_box.insertPlainText("Point {}: ({}, {})\n".format(i, series1.at(i).x(), series1.at(i).y()))
def update_info_y():
    axis_y_box.clear()
    axis_y_box.insertPlainText("")
    for i in range(series2.count()):
        axis_y_box.insertPlainText("Point {}: ({}, {})\n".format(i, series2.at(i).x(), series2.at(i).y()))
def update_info_z():
    axis_z_box.clear()
    axis_z_box.insertPlainText("")
    for i in range(series3.count()):
        axis_z_box.insertPlainText("Point {}: ({}, {})\n".format(i, series3.at(i).x(), series3.at(i).y()))



def save_excel():
    # Disable the save button and change its color to gray
    save_button.setEnabled(False)
    save_button.setStyleSheet("background-color: gray;")

    # Use QTimer.singleShot() to enable the save button after 1 second
    QTimer.singleShot(1000, lambda: save_button.setEnabled(True))

    # Change the save button color back to green
    QTimer.singleShot(1000, lambda: save_button.setStyleSheet("background-color: green;"))

    # Récupération du nom de fichier saisi
    file_name = save_text.text()

    # Obtenir la date actuelle
    current_date = datetime.datetime.now().strftime("%d-%m-%Y")

    # Ajouter la date au nom du fichier
    file_name_with_date = f"{file_name}_{current_date}"

    # BEtter name
    Nom_du_dossier = save_text.text()
    Nomdufichier = f"{file_name_with_date}.xlsx"

    # Nom du chemin
    chemin = ''+Nom_du_dossier
    chemin_vers_dossier = chemin + '/' + Nomdufichier

    # Création du fichier Excel avec le nom saisi
    workbook = xlsxwriter.Workbook(chemin_vers_dossier)

    # Ajout d'une feuille de calcul pour chaque axe
    sheet1 = workbook.add_worksheet("Axe X")
    sheet2 = workbook.add_worksheet("Axe Y")
    sheet3 = workbook.add_worksheet("Axe Z")

    # Écriture des en-têtes des colonnes
    sheet1.write(0, 0, "Fréquence [Hz]")
    sheet1.write(0, 1, "Accéléromètre [g]")
    sheet2.write(0, 0, "Fréquence [Hz]")
    sheet2.write(0, 1, "Accéléromètre [g]")
    sheet3.write(0, 0, "Fréquence [Hz]")
    sheet3.write(0, 1, "Accéléromètre [g]")

    # Récupération des données de chaque axe
    x_data = [(i, series1.at(i).y()) for i in range(series1.count())]
    y_data = [(i, series2.at(i).y()) for i in range(series2.count())]
    z_data = [(i, series3.at(i).y()) for i in range(series3.count())]

    # Écriture des données dans chaque feuille de calcul
    for i, data in enumerate([x_data, y_data, z_data]):
        for row, (x, y) in enumerate(data):
            sheet = [sheet1, sheet2, sheet3][i]
            sheet.write(row + 1, 0, x)
            sheet.write(row + 1, 1, y)

    # Ajuster la taille des cellules en fonction du contenu
    for sheet in [sheet1, sheet2, sheet3]:
        sheet.set_column(0, 0, 15)  # Ajuster la largeur de la colonne 0 (A)
        sheet.set_column(1, 1, 15)  # Ajuster la largeur de la colonne 1 (B)
  
    # créez le dossier s'il n'existe pas déjà
    if not os.path.isdir(Nom_du_dossier):
        os.mkdir(Nom_du_dossier)
    else:
        pass


    # Fermeture du fichier Excel
    workbook.close()
    
   
# Lien entre le bouton "Commencer" et la fonction d'affichage des points et de l'updade des info du graphique 
button.clicked.connect(show_curve)        
button.clicked.connect(update_info_x)
button.clicked.connect(update_info_y)
button.clicked.connect(update_info_z)

# Lien entre le bouton "Effacer" et la fonction pour nettoyer le graphique et de l'updade des info du graphique 
clear_button.clicked.connect(clear_clicked)
clear_button.clicked.connect(update_info_x)
clear_button.clicked.connect(update_info_y)
clear_button.clicked.connect(update_info_z)
clear_button.clicked.connect(reset_timer)

# Lien entre le bouton "sauvegarder" et les fonctions pour sauvegarder en fichier exel et image
save_button.clicked.connect(save_excel)
save_button.clicked.connect(save_image)



# Met a jour le graphique 
chart.scaleChanged.connect(update_info_x)
chart.scaleChanged.connect(update_info_y) 
chart.scaleChanged.connect(update_info_z) 



# Ajoute le layout a la fenetre et affiche la fenetre
window.setLayout(layout)
window.show()
sys.exit(app.exec())

