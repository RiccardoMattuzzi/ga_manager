from PyQt5.QtWidgets import (QApplication, QWidget, QVBoxLayout, QTabWidget,
                             QTableWidget, QTableWidgetItem, QComboBox, QPushButton,
                             QHBoxLayout, QFileDialog, QMessageBox)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QColor, QBrush
import sys
import sqlite3
import os
from openpyxl import Workbook, load_workbook

DB_PATH = os.path.expanduser("~/Documents/App candele/materie_prime.db")

def parse_float(value):
    """Converte stringa a float, accetta sia virgola che punto"""
    try:
        return float(str(value).replace(",", "."))
    except:
        return 0.0

def converti_unita(quantita, unita):
    """Converte quantità a unità standard (kg, litri, pz) e ritorna (quantita_convertita, unita_standard)"""
    quantita = parse_float(quantita)

    if unita == "grammi":
        return quantita / 1000, "kg"
    elif unita == "chili":
        return quantita, "kg"
    elif unita == "litri":
        return quantita, "litro"
    elif unita == "ml":
        return quantita / 1000, "litro"
    elif unita == "pezzi":
        return quantita, "pz"
    else:
        return quantita, unita

def calcola_costo_unitario(prezzo, quantita, unita):
    """Calcola costo unitario con unità corretta"""
    prezzo = parse_float(prezzo)

    if quantita == "" or quantita == "0":
        return "—", ""

    quantita_convertita, unita_std = converti_unita(quantita, unita)

    if quantita_convertita == 0:
        return "—", ""

    costo_unit = prezzo / quantita_convertita
    return f"{costo_unit:.3f}", f"€/{unita_std}"

def init_db():
    """Crea il database se non esiste"""
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS materie_prime
                 (id INTEGER PRIMARY KEY,
                  nome TEXT,
                  prezzo REAL,
                  quantita REAL,
                  unita TEXT)''')
    conn.commit()
    conn.close()

def load_data():
    """Legge i dati dal database"""
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("SELECT id, nome, prezzo, quantita, unita FROM materie_prime")
    data = c.fetchall()
    conn.close()
    return data

def aggiorna_costo_unitario(table, row):
    """Aggiorna la colonna costo unitario per una riga specifica"""
    prezzo = table.item(row, 1).text() if table.item(row, 1) else ""
    quantita = table.item(row, 2).text() if table.item(row, 2) else ""
    combo = table.cellWidget(row, 3)
    unita = combo.currentText() if combo else ""

    costo, unita_std = calcola_costo_unitario(prezzo, quantita, unita)

    item = QTableWidgetItem(f"{costo} {unita_std}")
    item.setBackground(QBrush(QColor(60, 60, 60)))
    item.setForeground(QBrush(QColor(200, 200, 200)))
    item.setFlags(item.flags() & ~Qt.ItemIsEditable)
    table.setItem(row, 4, item)

def save_data(table):
    """Salva tutti i dati della tabella nel database"""
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("DELETE FROM materie_prime")

    for row in range(table.rowCount()):
        nome = table.item(row, 0).text() if table.item(row, 0) else ""
        prezzo = parse_float(table.item(row, 1).text() if table.item(row, 1) else "0")
        quantita = parse_float(table.item(row, 2).text() if table.item(row, 2) else "0")
        combo = table.cellWidget(row, 3)
        unita = combo.currentText() if combo else ""

        if nome:
            c.execute("INSERT INTO materie_prime (nome, prezzo, quantita, unita) VALUES (?, ?, ?, ?)",
                     (nome, prezzo, quantita, unita))

    conn.commit()
    conn.close()

def esporta_excel(table):
    """Esporta i dati della tabella in Excel"""
    file_path, _ = QFileDialog.getSaveFileName(None, "Salva come Excel", "", "Excel Files (*.xlsx)")
    if not file_path:
        return

    if not file_path.endswith('.xlsx'):
        file_path += '.xlsx'

    wb = Workbook()
    ws = wb.active
    ws.title = "Materie Prime"

    ws.append(["Materia Prima", "Prezzo Confezione (€)", "Quantità", "Unità"])

    for row in range(table.rowCount()):
        nome = table.item(row, 0).text() if table.item(row, 0) else ""
        prezzo = table.item(row, 1).text() if table.item(row, 1) else ""
        quantita = table.item(row, 2).text() if table.item(row, 2) else ""
        combo = table.cellWidget(row, 3)
        unita = combo.currentText() if combo else ""

        ws.append([nome, prezzo, quantita, unita])

    wb.save(file_path)
    QMessageBox.information(None, "Successo", f"Dati esportati in: {file_path}")

def importa_excel(table, unita_misura):
    """Importa dati da Excel nella tabella"""
    file_path, _ = QFileDialog.getOpenFileName(None, "Apri file Excel", "", "Excel Files (*.xlsx)")
    if not file_path:
        return

    try:
        wb = load_workbook(file_path)
        ws = wb.active

        table.setRowCount(0)

        for row_idx, row in enumerate(ws.iter_rows(min_row=2, values_only=True), 1):
            if row[0] is None:
                break

            table.insertRow(table.rowCount())
            current_row = table.rowCount() - 1

            table.setItem(current_row, 0, QTableWidgetItem(str(row[0]) if row[0] else ""))
            table.setItem(current_row, 1, QTableWidgetItem(str(parse_float(row[1])) if row[1] else ""))
            table.setItem(current_row, 2, QTableWidgetItem(str(parse_float(row[2])) if row[2] else ""))

            combo = QComboBox()
            combo.addItems(unita_misura)
            combo.blockSignals(True)
            if row[3]:
                combo.setCurrentText(str(row[3]))
            combo.blockSignals(False)
            combo.currentIndexChanged.connect(lambda r=current_row: (aggiorna_costo_unitario(table, r), save_data(table)))
            table.setCellWidget(current_row, 3, combo)

            aggiorna_costo_unitario(table, current_row)

        save_data(table)
        QMessageBox.information(None, "Successo", "Dati importati correttamente")
    except Exception as e:
        QMessageBox.critical(None, "Errore", f"Errore nell'importazione: {str(e)}")

def aggiungi_riga(table, unita):
    row = table.rowCount()
    table.insertRow(row)
    table.setItem(row, 0, QTableWidgetItem(""))
    table.setItem(row, 1, QTableWidgetItem(""))
    table.setItem(row, 2, QTableWidgetItem(""))
    combo = QComboBox()
    combo.addItems(unita)
    combo.blockSignals(True)
    combo.setCurrentText("grammi")
    combo.blockSignals(False)
    combo.currentIndexChanged.connect(lambda r=row: (aggiorna_costo_unitario(table, r), save_data(table)))
    table.setCellWidget(row, 3, combo)
    aggiorna_costo_unitario(table, row)

app = QApplication(sys.argv)
window = QWidget()
window.setWindowTitle("App Candele")
window.setGeometry(100, 100, 900, 500)

init_db()

tabs = QTabWidget()

# Tab 1: Materie prime
tab1 = QWidget()
layout1 = QVBoxLayout()

table = QTableWidget()
table.setColumnCount(5)
table.setHorizontalHeaderLabels(["Materia Prima", "Prezzo Confezione (€)", "Quantità", "Unità", "Costo Unitario"])

unita_misura = ["grammi", "chili", "litri", "ml", "pezzi"]

data = load_data()
table.setRowCount(len(data) if data else 5)

for row, record in enumerate(data):
    id_db, nome, prezzo, quantita, unita = record
    table.setItem(row, 0, QTableWidgetItem(nome))
    table.setItem(row, 1, QTableWidgetItem(str(prezzo)))
    table.setItem(row, 2, QTableWidgetItem(str(quantita)))

    combo = QComboBox()
    combo.addItems(unita_misura)
    combo.blockSignals(True)
    combo.setCurrentText(unita)
    combo.blockSignals(False)
    combo.currentIndexChanged.connect(lambda r=row: (aggiorna_costo_unitario(table, r), save_data(table)))
    table.setCellWidget(row, 3, combo)

    aggiorna_costo_unitario(table, row)

if not data:
    for row in range(5):
        table.setItem(row, 0, QTableWidgetItem(""))
        table.setItem(row, 1, QTableWidgetItem(""))
        table.setItem(row, 2, QTableWidgetItem(""))
        combo = QComboBox()
        combo.addItems(unita_misura)
        combo.blockSignals(True)
        combo.setCurrentText("grammi")
        combo.blockSignals(False)
        combo.currentIndexChanged.connect(lambda r=row: (aggiorna_costo_unitario(table, r), save_data(table)))
        table.setCellWidget(row, 3, combo)
        aggiorna_costo_unitario(table, row)

table.setEditTriggers(QTableWidget.DoubleClicked | QTableWidget.SelectedClicked)

def on_cell_changed(row, col):
    if col in [1, 2]:  # Se cambia prezzo o quantità
        aggiorna_costo_unitario(table, row)
    save_data(table)

table.cellChanged.connect(on_cell_changed)

button_layout = QHBoxLayout()

button_add = QPushButton("Aggiungi riga")
button_add.clicked.connect(lambda: aggiungi_riga(table, unita_misura))

button_export = QPushButton("Esporta dati in Excel")
button_export.clicked.connect(lambda: esporta_excel(table))

button_import = QPushButton("Importa dati da Excel")
button_import.clicked.connect(lambda: importa_excel(table, unita_misura))

button_layout.addWidget(button_add)
button_layout.addWidget(button_export)
button_layout.addWidget(button_import)

layout1.addWidget(table)
layout1.addLayout(button_layout)
tab1.setLayout(layout1)

# Tab 2: Prodotti
tab2 = QWidget()
layout2 = QVBoxLayout()
tab2.setLayout(layout2)

# Tab 3: Prezzi
tab3 = QWidget()
layout3 = QVBoxLayout()
tab3.setLayout(layout3)

tabs.addTab(tab1, "Materie prime")
tabs.addTab(tab2, "Prodotti")
tabs.addTab(tab3, "Prezzi")

main_layout = QVBoxLayout()
main_layout.addWidget(tabs)
window.setLayout(main_layout)

def closeEvent(event):
    save_data(table)
    event.accept()

window.closeEvent = closeEvent

window.show()
sys.exit(app.exec_())
