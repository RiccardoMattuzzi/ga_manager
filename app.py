# ============================================================================
# INCLUDE
# ============================================================================
import sys
import sqlite3
from pathlib import Path
from PyQt5.QtWidgets import (QApplication, QMainWindow, QTabWidget, QWidget, QVBoxLayout, QHBoxLayout,
                             QTableWidget, QTableWidgetItem, QComboBox, QStyledItemDelegate,
                             QLineEdit, QPushButton, QFileDialog, QMessageBox)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QColor, QBrush
from openpyxl import Workbook, load_workbook

# ============================================================================
# VARIABILI
# ============================================================================
NOME_TAB_1 = "Materie Prime"
NOME_TAB_2 = "Prodotti"
NOME_TAB_3 = "Costi"

COLONNA_NOME = "Nome"
COLONNA_COSTO = "Costo Confezione (€)"
COLONNA_QUANTITA = "Quantità"
COLONNA_UNITA = "Unità di Misura"

UNITA_MISURA = ["kg", "grammi", "litri", "ml", "pezzo/i", "m", "cm"]

# ============================================================================
# COSTANTI
# ============================================================================
WINDOW_WIDTH = 1200
WINDOW_HEIGHT = 600
WINDOW_TITLE = "App Candele"

# Fattori di conversione (unità base: kg, litri, m)
FATTORI_CONVERSIONE = {
    "kg": 1.0,
    "grammi": 1000.0,
    "litri": 1.0,
    "ml": 1000.0,
    "m": 1.0,
    "cm": 100.0,
    "pezzo/i": 1.0
}

# Database
PROJECT_DIR = Path(__file__).parent.absolute()
DATA_DIR = PROJECT_DIR / "data"
DB_PATH = DATA_DIR / "materie_prime.db"
EXPORT_DIR = PROJECT_DIR / "exports"

# ============================================================================
# FUNZIONI E CLASSI
# ============================================================================

def parse_float(value):
    """Converte stringa a float, accetta virgola e punto"""
    try:
        return float(str(value).replace(",", "."))
    except:
        return 0.0

def init_db():
    """Inizializza il database"""
    DATA_DIR.mkdir(exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS materie_prime
                 (id INTEGER PRIMARY KEY,
                  nome TEXT,
                  costo REAL,
                  quantita REAL,
                  unita TEXT)''')
    conn.commit()
    conn.close()

def load_data():
    """Carica i dati dal database"""
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("SELECT id, nome, costo, quantita, unita FROM materie_prime")
    data = c.fetchall()
    conn.close()
    return data

def save_data(table):
    """Salva tutti i dati della tabella nel database"""
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("DELETE FROM materie_prime")
    
    for row in range(table.rowCount()):
        nome = table.item(row, 0).text() if table.item(row, 0) else ""
        costo = parse_float(table.item(row, 1).text() if table.item(row, 1) else "0")
        quantita = parse_float(table.item(row, 2).text() if table.item(row, 2) else "0")
        combo = table.cellWidget(row, 3)
        unita = combo.currentText() if combo else ""
        
        if nome:  # Salva solo se il nome non è vuoto
            c.execute("INSERT INTO materie_prime (nome, costo, quantita, unita) VALUES (?, ?, ?, ?)",
                     (nome, costo, quantita, unita))
    
    conn.commit()
    conn.close()

def esporta_excel(table):
    """Esporta la tabella in un file Excel"""
    file_path, _ = QFileDialog.getSaveFileName(None, "Salva come Excel", "", "Excel Files (*.xlsx)")
    if not file_path:
        return
    
    if not file_path.endswith('.xlsx'):
        file_path += '.xlsx'
    
    wb = Workbook()
    ws = wb.active
    ws.title = "Materie Prime"
    
    # Header
    ws.append([COLONNA_NOME, COLONNA_COSTO, COLONNA_QUANTITA, COLONNA_UNITA])
    
    # Dati
    for row in range(table.rowCount()):
        nome = table.item(row, 0).text() if table.item(row, 0) else ""
        costo = table.item(row, 1).text() if table.item(row, 1) else ""
        quantita = table.item(row, 2).text() if table.item(row, 2) else ""
        combo = table.cellWidget(row, 3)
        unita = combo.currentText() if combo else ""
        
        if nome:  # Esporta solo righe con nome
            ws.append([nome, parse_float(costo), parse_float(quantita), unita])
    
    wb.save(file_path)
    QMessageBox.information(None, "Successo", f"Dati esportati in: {file_path}")

def importa_excel(table):
    """Importa dati da un file Excel"""
    file_path, _ = QFileDialog.getOpenFileName(None, "Apri file Excel", "", "Excel Files (*.xlsx)")
    if not file_path:
        return
    
    try:
        wb = load_workbook(file_path)
        ws = wb.active
        
        # Pulisci la tabella
        table.setRowCount(0)
        
        # Leggi i dati dal file (salta header)
        for row_idx, row in enumerate(ws.iter_rows(min_row=2, values_only=True), 0):
            if row[0] is None:  # Ferma alla prima riga vuota
                break
            
            table.insertRow(table.rowCount())
            current_row = table.rowCount() - 1
            
            # Colonna 0: Nome
            table.setItem(current_row, 0, QTableWidgetItem(str(row[0]) if row[0] else ""))
            
            # Colonna 1: Costo
            table.setItem(current_row, 1, QTableWidgetItem(str(parse_float(row[1])) if row[1] else ""))
            
            # Colonna 2: Quantità
            table.setItem(current_row, 2, QTableWidgetItem(str(parse_float(row[2])) if row[2] else ""))
            
            # Colonna 3: Unità
            combo = QComboBox()
            combo.addItems(UNITA_MISURA)
            combo.blockSignals(True)
            if row[3] and str(row[3]) in UNITA_MISURA:
                combo.setCurrentText(str(row[3]))
            combo.blockSignals(False)
            combo.currentIndexChanged.connect(lambda: save_data(table))
            table.setCellWidget(current_row, 3, combo)
        
        save_data(table)
        QMessageBox.information(None, "Successo", "Dati importati correttamente")
    except Exception as e:
        QMessageBox.critical(None, "Errore", f"Errore nell'importazione: {str(e)}")

class NumericDelegate(QStyledItemDelegate):
    """Delegate che converte automaticamente virgole in punti durante la digitazione"""
    
    def createEditor(self, parent, option, index):
        editor = QLineEdit(parent)
        editor.textChanged.connect(lambda text: self._converti_virgola(editor))
        return editor
    
    def setEditorData(self, editor, index):
        value = index.model().data(index, Qt.EditRole)
        editor.setText(str(value) if value else "")
        editor.textChanged.connect(lambda text: self._converti_virgola(editor))
    
    def setModelData(self, editor, model, index):
        text = editor.text().replace(",", ".")
        model.setData(index, text, Qt.EditRole)
    
    def _converti_virgola(self, editor):
        cursor_pos = editor.cursorPosition()
        text = editor.text().replace(",", ".")
        editor.blockSignals(True)
        editor.setText(text)
        editor.blockSignals(False)
        editor.setCursorPosition(cursor_pos)

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(WINDOW_TITLE)
        self.setGeometry(100, 100, WINDOW_WIDTH, WINDOW_HEIGHT)
        
        # Inizializza database
        init_db()
        
        # Crea widget principale
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # Crea layout principale
        main_layout = QVBoxLayout()
        central_widget.setLayout(main_layout)
        
        # Crea tab widget
        tabs = QTabWidget()
        main_layout.addWidget(tabs)
        
        # Tab 1: Materie Prime
        tab1 = QWidget()
        tab1_layout = QVBoxLayout()
        tab1.setLayout(tab1_layout)
        
        self.table_materie = QTableWidget()
        self.setup_tabella_materie()
        tab1_layout.addWidget(self.table_materie)
        
        # Layout pulsanti
        button_layout = QHBoxLayout()
        
        button_export = QPushButton("Esporta in Excel")
        button_export.clicked.connect(lambda: esporta_excel(self.table_materie))
        
        button_import = QPushButton("Importa da Excel")
        button_import.clicked.connect(lambda: importa_excel(self.table_materie))
        
        button_layout.addWidget(button_export)
        button_layout.addWidget(button_import)
        
        tab1_layout.addLayout(button_layout)
        
        # Tab 2 e 3 vuoti
        tab2 = QWidget()
        tab3 = QWidget()
        
        tabs.addTab(tab1, NOME_TAB_1)
        tabs.addTab(tab2, NOME_TAB_2)
        tabs.addTab(tab3, NOME_TAB_3)
    
    def setup_tabella_materie(self):
        """Configura la tabella delle materie prime"""
        self.table_materie.setColumnCount(4)
        self.table_materie.setHorizontalHeaderLabels([
            COLONNA_NOME,
            COLONNA_COSTO,
            COLONNA_QUANTITA,
            COLONNA_UNITA
        ])
        
        # Carica i dati dal database
        data = load_data()
        self.table_materie.setRowCount(len(data) if data else 5)
        
        # Configura il delegate numerico per la colonna costo
        numeric_delegate = NumericDelegate()
        self.table_materie.setItemDelegateForColumn(1, numeric_delegate)
        
        # Popola le righe con dati dal DB o vuote
        for row, record in enumerate(data) if data else enumerate([]):
            id_db, nome, costo, quantita, unita = record
            self.table_materie.setItem(row, 0, QTableWidgetItem(nome))
            self.table_materie.setItem(row, 1, QTableWidgetItem(str(costo)))
            self.table_materie.setItem(row, 2, QTableWidgetItem(str(quantita)))
            
            combo = QComboBox()
            combo.addItems(UNITA_MISURA)
            combo.blockSignals(True)
            combo.setCurrentText(unita)
            combo.blockSignals(False)
            combo.currentIndexChanged.connect(lambda: save_data(self.table_materie))
            self.table_materie.setCellWidget(row, 3, combo)
        
        # Se non ci sono dati, aggiungi righe vuote
        if not data:
            for row in range(5):
                self.table_materie.setItem(row, 0, QTableWidgetItem(""))
                self.table_materie.setItem(row, 1, QTableWidgetItem(""))
                self.table_materie.setItem(row, 2, QTableWidgetItem(""))
                
                combo = QComboBox()
                combo.addItems(UNITA_MISURA)
                combo.currentIndexChanged.connect(lambda: save_data(self.table_materie))
                self.table_materie.setCellWidget(row, 3, combo)
        
        # Consenti editing
        self.table_materie.setEditTriggers(QTableWidget.DoubleClicked | QTableWidget.SelectedClicked)
        
        # Salva quando una cella cambia
        self.table_materie.cellChanged.connect(lambda: save_data(self.table_materie))
        
        # Ridimensiona colonne
        self.table_materie.resizeColumnsToContents()

# ============================================================================
# MAIN
# ============================================================================

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())