# ============================================================================
# INCLUDE
# ============================================================================
import sys
import sqlite3
from pathlib import Path
from PyQt5.QtWidgets import (QApplication, QMainWindow, QTabWidget, QWidget, QVBoxLayout, QHBoxLayout,
                             QTableWidget, QTableWidgetItem, QComboBox, QStyledItemDelegate,
                             QLineEdit, QPushButton, QFileDialog, QMessageBox, QListWidget,
                             QListWidgetItem, QLabel, QSpinBox, QDoubleSpinBox)
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

TIPI_PRODOTTO = ["Candela soia", "Jesmonite", "Candela soia + involucro jesmonite", "Lampada"]

# ============================================================================
# COSTANTI
# ============================================================================
WINDOW_WIDTH = 1400
WINDOW_HEIGHT = 700
WINDOW_TITLE = "App Candele"

FATTORI_CONVERSIONE = {
    "kg": 1.0,
    "grammi": 1000.0,
    "litri": 1.0,
    "ml": 1000.0,
    "m": 1.0,
    "cm": 100.0,
    "pezzo/i": 1.0
}

PROJECT_DIR = Path(__file__).parent.absolute()
DATA_DIR = PROJECT_DIR / "data"
DB_PATH = DATA_DIR / "materie_prime.db"
EXPORT_DIR = PROJECT_DIR / "exports"

MAX_PROFONDITA_RICORSIONE = 5

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
    
    c.execute('''CREATE TABLE IF NOT EXISTS prodotti
                 (id INTEGER PRIMARY KEY,
                  nome TEXT,
                  tipo TEXT)''')
    
    c.execute('''CREATE TABLE IF NOT EXISTS composizione
                 (id INTEGER PRIMARY KEY,
                  prodotto_id INTEGER,
                  elemento_id INTEGER,
                  elemento_tipo TEXT,
                  quantita REAL,
                  FOREIGN KEY(prodotto_id) REFERENCES prodotti(id))''')
    
    conn.commit()
    conn.close()

def load_materie_prime():
    """Carica tutte le materie prime dal database"""
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("SELECT id, nome, costo, quantita, unita FROM materie_prime")
    data = c.fetchall()
    conn.close()
    return data

def get_materia_prima_by_id(materia_id):
    """Carica una materia prima specifica"""
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("SELECT id, nome, costo, quantita, unita FROM materie_prime WHERE id = ?", (materia_id,))
    data = c.fetchone()
    conn.close()
    return data

def save_materie_prime(table):
    """Salva tutte le materie prime nel database"""
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("DELETE FROM materie_prime")
    
    for row in range(table.rowCount()):
        nome = table.item(row, 0).text() if table.item(row, 0) else ""
        costo = parse_float(table.item(row, 1).text() if table.item(row, 1) else "0")
        quantita = parse_float(table.item(row, 2).text() if table.item(row, 2) else "0")
        combo = table.cellWidget(row, 3)
        unita = combo.currentText() if combo else ""
        
        if nome:
            c.execute("INSERT INTO materie_prime (nome, costo, quantita, unita) VALUES (?, ?, ?, ?)",
                     (nome, costo, quantita, unita))
    
    conn.commit()
    conn.close()

def load_prodotti():
    """Carica tutti i prodotti"""
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("SELECT id, nome, tipo FROM prodotti ORDER BY nome")
    data = c.fetchall()
    conn.close()
    return data

def get_prodotto_by_id(prodotto_id):
    """Carica un prodotto specifico"""
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("SELECT id, nome, tipo FROM prodotti WHERE id = ?", (prodotto_id,))
    data = c.fetchone()
    conn.close()
    return data

def save_prodotto(prodotto_id, nome, tipo):
    """Salva o aggiorna un prodotto"""
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    
    if prodotto_id is None:
        c.execute("INSERT INTO prodotti (nome, tipo) VALUES (?, ?)", (nome, tipo))
        conn.commit()
        new_id = c.lastrowid
        conn.close()
        return new_id
    else:
        c.execute("UPDATE prodotti SET nome = ?, tipo = ? WHERE id = ?", (nome, tipo, prodotto_id))
        conn.commit()
        conn.close()
        return prodotto_id

def delete_prodotto(prodotto_id):
    """Elimina un prodotto e la sua composizione"""
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("DELETE FROM composizione WHERE prodotto_id = ?", (prodotto_id,))
    c.execute("DELETE FROM prodotti WHERE id = ?", (prodotto_id,))
    conn.commit()
    conn.close()

def load_composizione(prodotto_id):
    """Carica la composizione di un prodotto"""
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("SELECT id, elemento_id, elemento_tipo, quantita FROM composizione WHERE prodotto_id = ? ORDER BY id",
             (prodotto_id,))
    data = c.fetchall()
    conn.close()
    return data

def save_composizione(prodotto_id, composizione_data):
    """Salva la composizione di un prodotto"""
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("DELETE FROM composizione WHERE prodotto_id = ?", (prodotto_id,))
    
    for elemento_tipo, elemento_id, quantita in composizione_data:
        if elemento_id:
            c.execute("INSERT INTO composizione (prodotto_id, elemento_id, elemento_tipo, quantita) VALUES (?, ?, ?, ?)",
                     (prodotto_id, elemento_id, elemento_tipo, quantita))
    
    conn.commit()
    conn.close()

def aggiungi_elemento_composizione(prodotto_id):
    """Aggiunge una riga vuota alla composizione"""
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("INSERT INTO composizione (prodotto_id, elemento_id, elemento_tipo, quantita) VALUES (?, NULL, ?, ?)",
             (prodotto_id, "materia_prima", 0))
    conn.commit()
    conn.close()

def elimina_elemento_composizione(elemento_id):
    """Elimina un elemento dalla composizione"""
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("DELETE FROM composizione WHERE id = ?", (elemento_id,))
    conn.commit()
    conn.close()

def calcola_costo_prodotto(prodotto_id, visited=None, profondita=0):
    """Calcola il costo totale di un prodotto ricorsivamente"""
    if visited is None:
        visited = set()
    
    if profondita > MAX_PROFONDITA_RICORSIONE:
        return 0.0
    
    if prodotto_id in visited:
        return 0.0
    
    visited.add(prodotto_id)
    
    costo_totale = 0.0
    composizione = load_composizione(prodotto_id)
    
    for comp_id, elemento_id, elemento_tipo, quantita in composizione:
        quantita = parse_float(quantita)
        
        if elemento_tipo == "materia_prima":
            materia = get_materia_prima_by_id(elemento_id)
            if materia:
                id_m, nome, costo, qta, unita = materia
                costo_unitario = costo / qta if qta > 0 else 0
                costo_totale += costo_unitario * quantita
        
        elif elemento_tipo == "prodotto":
            costo_sub = calcola_costo_prodotto(elemento_id, visited.copy(), profondita + 1)
            costo_totale += costo_sub * quantita
    
    return costo_totale

def detecta_ciclo(prodotto_id, elemento_id, elemento_tipo):
    """Verifica se aggiungere questo elemento creerebbe un ciclo"""
    if elemento_tipo != "prodotto":
        return False
    
    if elemento_id == prodotto_id:
        return True
    
    composizione = load_composizione(elemento_id)
    for comp_id, el_id, el_tipo, qta in composizione:
        if el_tipo == "prodotto" and detecta_ciclo(prodotto_id, el_id, el_tipo):
            return True
    
    return False

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
    
    ws.append([COLONNA_NOME, COLONNA_COSTO, COLONNA_QUANTITA, COLONNA_UNITA])
    
    for row in range(table.rowCount()):
        nome = table.item(row, 0).text() if table.item(row, 0) else ""
        costo = table.item(row, 1).text() if table.item(row, 1) else ""
        quantita = table.item(row, 2).text() if table.item(row, 2) else ""
        combo = table.cellWidget(row, 3)
        unita = combo.currentText() if combo else ""
        
        if nome:
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
        
        table.setRowCount(0)
        
        for row_idx, row in enumerate(ws.iter_rows(min_row=2, values_only=True), 0):
            if row[0] is None:
                break
            
            table.insertRow(table.rowCount())
            current_row = table.rowCount() - 1
            
            table.setItem(current_row, 0, QTableWidgetItem(str(row[0]) if row[0] else ""))
            table.setItem(current_row, 1, QTableWidgetItem(str(parse_float(row[1])) if row[1] else ""))
            table.setItem(current_row, 2, QTableWidgetItem(str(parse_float(row[2])) if row[2] else ""))
            
            combo = QComboBox()
            combo.addItems(UNITA_MISURA)
            combo.blockSignals(True)
            if row[3] and str(row[3]) in UNITA_MISURA:
                combo.setCurrentText(str(row[3]))
            combo.blockSignals(False)
            combo.currentIndexChanged.connect(lambda: save_materie_prime(table))
            table.setCellWidget(current_row, 3, combo)
        
        save_materie_prime(table)
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
        
        init_db()
        
        self.composizione_ids = {}  # Mappa row -> comp_id
        
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        main_layout = QVBoxLayout()
        central_widget.setLayout(main_layout)
        
        tabs = QTabWidget()
        main_layout.addWidget(tabs)
        
        # Tab 1: Materie Prime
        tab1 = QWidget()
        tab1_layout = QVBoxLayout()
        tab1.setLayout(tab1_layout)
        
        self.table_materie = QTableWidget()
        self.setup_tabella_materie()
        tab1_layout.addWidget(self.table_materie)
        
        button_layout = QHBoxLayout()
        button_export = QPushButton("Esporta in Excel")
        button_export.clicked.connect(lambda: esporta_excel(self.table_materie))
        button_import = QPushButton("Importa da Excel")
        button_import.clicked.connect(lambda: importa_excel(self.table_materie))
        button_layout.addWidget(button_export)
        button_layout.addWidget(button_import)
        tab1_layout.addLayout(button_layout)
        
        # Tab 2: Prodotti
        tab2 = QWidget()
        tab2_layout = QHBoxLayout()
        tab2.setLayout(tab2_layout)
        
        # Panel sinistro: lista prodotti
        left_layout = QVBoxLayout()
        left_layout.addWidget(QLabel("Prodotti:"))
        
        self.lista_prodotti = QListWidget()
        self.lista_prodotti.itemSelectionChanged.connect(self.on_prodotto_selezionato)
        left_layout.addWidget(self.lista_prodotti)
        
        btn_layout_sx = QHBoxLayout()
        btn_add_prodotto = QPushButton("Aggiungi")
        btn_add_prodotto.clicked.connect(self.aggiungi_prodotto)
        btn_del_prodotto = QPushButton("Elimina")
        btn_del_prodotto.clicked.connect(self.elimina_prodotto)
        btn_layout_sx.addWidget(btn_add_prodotto)
        btn_layout_sx.addWidget(btn_del_prodotto)
        left_layout.addLayout(btn_layout_sx)
        
        # Panel destro: dettagli prodotto
        right_layout = QVBoxLayout()
        
        self.current_prodotto_id = None
        
        right_layout.addWidget(QLabel("Nome:"))
        self.input_nome_prodotto = QLineEdit()
        self.input_nome_prodotto.textChanged.connect(self.on_change_dettagli)
        right_layout.addWidget(self.input_nome_prodotto)
        
        right_layout.addWidget(QLabel("Tipo:"))
        self.combo_tipo_prodotto = QComboBox()
        self.combo_tipo_prodotto.addItems(TIPI_PRODOTTO)
        self.combo_tipo_prodotto.currentIndexChanged.connect(self.on_change_dettagli)
        right_layout.addWidget(self.combo_tipo_prodotto)
        
        right_layout.addWidget(QLabel("Composizione:"))
        self.table_composizione = QTableWidget()
        self.setup_tabella_composizione()
        right_layout.addWidget(self.table_composizione)
        
        btn_layout_dx = QHBoxLayout()
        btn_add_elemento = QPushButton("Aggiungi materia/prodotto")
        btn_add_elemento.clicked.connect(self.aggiungi_elemento_composizione)
        btn_del_elemento = QPushButton("Elimina riga")
        btn_del_elemento.clicked.connect(self.elimina_elemento_composizione)
        btn_layout_dx.addWidget(btn_add_elemento)
        btn_layout_dx.addWidget(btn_del_elemento)
        right_layout.addLayout(btn_layout_dx)
        
        self.label_costo_totale = QLabel("Costo totale: 0.00 €")
        right_layout.addWidget(self.label_costo_totale)
        
        left_widget = QWidget()
        left_widget.setLayout(left_layout)
        left_widget.setMinimumWidth(250)
        
        right_widget = QWidget()
        right_widget.setLayout(right_layout)
        
        tab2_layout.addWidget(left_widget, 1)
        tab2_layout.addWidget(right_widget, 2)
        
        self.carica_lista_prodotti()
        
        # Tab 3: vuoto
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
        
        data = load_materie_prime()
        self.table_materie.setRowCount(len(data) if data else 5)
        
        numeric_delegate = NumericDelegate()
        self.table_materie.setItemDelegateForColumn(1, numeric_delegate)
        
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
            combo.currentIndexChanged.connect(lambda: save_materie_prime(self.table_materie))
            self.table_materie.setCellWidget(row, 3, combo)
        
        if not data:
            for row in range(5):
                self.table_materie.setItem(row, 0, QTableWidgetItem(""))
                self.table_materie.setItem(row, 1, QTableWidgetItem(""))
                self.table_materie.setItem(row, 2, QTableWidgetItem(""))
                
                combo = QComboBox()
                combo.addItems(UNITA_MISURA)
                combo.currentIndexChanged.connect(lambda: save_materie_prime(self.table_materie))
                self.table_materie.setCellWidget(row, 3, combo)
        
        self.table_materie.setEditTriggers(QTableWidget.DoubleClicked | QTableWidget.SelectedClicked)
        self.table_materie.cellChanged.connect(lambda: save_materie_prime(self.table_materie))
        self.table_materie.resizeColumnsToContents()
    
    def setup_tabella_composizione(self):
        """Configura la tabella di composizione"""
        self.table_composizione.setColumnCount(4)
        self.table_composizione.setHorizontalHeaderLabels(["Tipo", "Elemento", "Quantità", "Costo Unitario"])
        self.table_composizione.setRowCount(0)
    
    def carica_lista_prodotti(self):
        """Carica la lista di prodotti nella UI"""
        self.lista_prodotti.clear()
        prodotti = load_prodotti()
        
        for prod_id, nome, tipo in prodotti:
            costo = calcola_costo_prodotto(prod_id)
            item_text = f"{nome} ({tipo}) - {costo:.2f}€"
            item = QListWidgetItem(item_text)
            item.setData(Qt.UserRole, prod_id)
            self.lista_prodotti.addItem(item)
    
    def on_prodotto_selezionato(self):
        """Quando selezioni un prodotto, carica i dettagli"""
        item = self.lista_prodotti.currentItem()
        if not item:
            return
        
        self.current_prodotto_id = item.data(Qt.UserRole)
        prodotto = get_prodotto_by_id(self.current_prodotto_id)
        
        if prodotto:
            id_p, nome, tipo = prodotto
            self.input_nome_prodotto.blockSignals(True)
            self.combo_tipo_prodotto.blockSignals(True)
            
            self.input_nome_prodotto.setText(nome)
            self.combo_tipo_prodotto.setCurrentText(tipo)
            
            self.input_nome_prodotto.blockSignals(False)
            self.combo_tipo_prodotto.blockSignals(False)
            
            self.carica_composizione()
            self.aggiorna_costo_totale()
    
    def carica_composizione(self):
        """Carica la composizione del prodotto selezionato"""
        if not self.current_prodotto_id:
            return
        
        self.table_composizione.blockSignals(True)
        self.table_composizione.setRowCount(0)
        self.composizione_ids = {}
        composizione = load_composizione(self.current_prodotto_id)
        
        materie = {(m[0], "materia_prima"): m[1] for m in load_materie_prime()}
        prodotti = {(p[0], "prodotto"): p[1] for p in load_prodotti()}
        
        for comp_id, elemento_id, elemento_tipo, quantita in composizione:
            row = self.table_composizione.rowCount()
            self.table_composizione.insertRow(row)
            self.composizione_ids[row] = comp_id
            
            # Colonna 0: tipo
            combo_tipo = QComboBox()
            combo_tipo.addItems(["materia_prima", "prodotto"])
            combo_tipo.setCurrentText(elemento_tipo)
            combo_tipo.currentIndexChanged.connect(lambda t=row: self.on_change_composizione(t))
            self.table_composizione.setCellWidget(row, 0, combo_tipo)
            
            # Colonna 1: elemento
            combo_elemento = QComboBox()
            self.popola_combo_elemento(combo_elemento, elemento_tipo)
            combo_elemento.setProperty("last_tipo", elemento_tipo)
            if elemento_id:
                for i in range(combo_elemento.count()):
                    if combo_elemento.itemData(i) == elemento_id:
                        combo_elemento.setCurrentIndex(i)
                        break
            combo_elemento.currentIndexChanged.connect(lambda t=row: self.on_change_composizione(t))
            self.table_composizione.setCellWidget(row, 1, combo_elemento)
            
            # Colonna 2: quantità
            spin = QDoubleSpinBox()
            spin.setValue(parse_float(quantita))
            spin.setMinimum(0)
            spin.setMaximum(10000)
            spin.setSingleStep(0.1)
            spin.valueChanged.connect(lambda t=row: self.on_change_composizione(t))
            self.table_composizione.setCellWidget(row, 2, spin)
            
            # Colonna 3: costo unitario (read-only)
            self.aggiorna_costo_unitario_riga(row, elemento_id, elemento_tipo)
        
        self.table_composizione.blockSignals(False)
    
    def popola_combo_elemento(self, combo, elemento_tipo):
        """Popola un combobox con materie prime o prodotti"""
        combo.clear()
        combo.addItem("-- Seleziona --", None)
        
        if elemento_tipo == "materia_prima":
            materie = load_materie_prime()
            for mat_id, nome, costo, quantita, unita in materie:
                combo.addItem(nome, mat_id)
        else:
            prodotti = load_prodotti()
            for prod_id, nome, tipo in prodotti:
                if prod_id != self.current_prodotto_id:  # Evita di aggiungere se stesso
                    combo.addItem(f"{nome} ({tipo})", prod_id)
    
    def aggiorna_costo_unitario_riga(self, row, elemento_id, elemento_tipo):
        """Aggiorna il costo unitario di una riga"""
        costo_str = "0.00"
        
        if elemento_tipo == "materia_prima" and elemento_id:
            materia = get_materia_prima_by_id(elemento_id)
            if materia:
                id_m, nome, costo, qta, unita = materia
                costo_unitario = costo / qta if qta > 0 else 0
                costo_str = f"{costo_unitario:.2f}"
        
        elif elemento_tipo == "prodotto" and elemento_id:
            costo_prod = calcola_costo_prodotto(elemento_id)
            costo_str = f"{costo_prod:.2f}"
        
        item = QTableWidgetItem(costo_str)
        item.setFlags(item.flags() & ~Qt.ItemIsEditable)
        item.setBackground(QBrush(QColor(200, 200, 200)))
        self.table_composizione.setItem(row, 3, item)
    
    def on_change_composizione(self, row):
        """Quando cambia la composizione, aggiorna il costo"""
        # Verifica che i widget esistano
        combo_tipo = self.table_composizione.cellWidget(row, 0)
        combo_elemento = self.table_composizione.cellWidget(row, 1)
        
        if not combo_tipo or not combo_elemento:
            return
        
        elemento_tipo = combo_tipo.currentText()
        elemento_id = combo_elemento.currentData()
        
        # Se il tipo cambia, ripopola il combobox elemento
        old_tipo = combo_elemento.property("last_tipo")
        if old_tipo != elemento_tipo:
            combo_elemento.blockSignals(True)
            self.popola_combo_elemento(combo_elemento, elemento_tipo)
            combo_elemento.blockSignals(False)
            combo_elemento.setProperty("last_tipo", elemento_tipo)
        
        self.aggiorna_costo_unitario_riga(row, elemento_id, elemento_tipo)
        self.aggiorna_costo_totale()
        self.salva_composizione()
    
    def aggiorna_costo_totale(self):
        """Aggiorna il costo totale del prodotto"""
        if not self.current_prodotto_id:
            self.label_costo_totale.setText("Costo totale: 0.00 €")
            return
        
        costo = calcola_costo_prodotto(self.current_prodotto_id)
        self.label_costo_totale.setText(f"Costo totale: {costo:.2f} €")
    
    def salva_composizione(self):
        """Salva la composizione dal table"""
        if not self.current_prodotto_id:
            return
        
        composizione_data = []
        for row in range(self.table_composizione.rowCount()):
            combo_tipo = self.table_composizione.cellWidget(row, 0)
            combo_elemento = self.table_composizione.cellWidget(row, 1)
            spin_quantita = self.table_composizione.cellWidget(row, 2)
            
            elemento_tipo = combo_tipo.currentText()
            elemento_id = combo_elemento.currentData()
            quantita = spin_quantita.value()
            
            if elemento_id:
                # Verifica ciclo
                if detecta_ciclo(self.current_prodotto_id, elemento_id, elemento_tipo):
                    QMessageBox.warning(None, "Errore", "Aggiungere questo elemento creerebbe un ciclo!")
                    continue
                
                composizione_data.append((elemento_tipo, elemento_id, quantita))
        
        save_composizione(self.current_prodotto_id, composizione_data)
    
    def on_change_dettagli(self):
        """Quando cambiano i dettagli del prodotto"""
        if not self.current_prodotto_id:
            return
        
        nome = self.input_nome_prodotto.text()
        tipo = self.combo_tipo_prodotto.currentText()
        
        if nome:
            save_prodotto(self.current_prodotto_id, nome, tipo)
            self.carica_lista_prodotti()
            self.aggiorna_costo_totale()
    
    def aggiungi_prodotto(self):
        """Aggiunge un nuovo prodotto"""
        prod_id = save_prodotto(None, "Nuovo Prodotto", TIPI_PRODOTTO[0])
        self.carica_lista_prodotti()
        
        # Seleziona il nuovo prodotto
        for i in range(self.lista_prodotti.count()):
            if self.lista_prodotti.item(i).data(Qt.UserRole) == prod_id:
                self.lista_prodotti.setCurrentRow(i)
                break
    
    def elimina_prodotto(self):
        """Elimina il prodotto selezionato"""
        item = self.lista_prodotti.currentItem()
        if not item:
            return
        
        if QMessageBox.question(None, "Conferma", "Eliminare questo prodotto?") == QMessageBox.Yes:
            prod_id = item.data(Qt.UserRole)
            delete_prodotto(prod_id)
            self.carica_lista_prodotti()
            self.table_composizione.setRowCount(0)
            self.current_prodotto_id = None
    
    def aggiungi_elemento_composizione(self):
        """Aggiunge una riga alla composizione"""
        if not self.current_prodotto_id:
            return
        
        aggiungi_elemento_composizione(self.current_prodotto_id)
        self.carica_composizione()
    
    def elimina_elemento_composizione(self):
        """Elimina una riga dalla composizione"""
        row = self.table_composizione.currentRow()
        if row < 0:
            return
        
        comp_id = self.composizione_ids.get(row)
        
        if comp_id:
            elimina_elemento_composizione(comp_id)
            self.carica_composizione()
            self.aggiorna_costo_totale()

# ============================================================================
# MAIN
# ============================================================================

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())