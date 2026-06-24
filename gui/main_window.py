from PySide6.QtWidgets import (QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
                               QTabWidget, QLabel, QPushButton, QLineEdit, QComboBox,
                               QTableWidget, QTableWidgetItem, QHeaderView, QSpinBox,
                               QMessageBox, QFileDialog, QDoubleSpinBox)
from PySide6.QtCore import Qt, Signal
from .widgets import FloatLineEdit
from io import ExcelHandler

class MainWindow(QMainWindow):
    """Main application window"""
    
    material_updated = Signal()  # Signal emitted when materials are updated
    
    def __init__(self, material_manager, product_manager, cost_calculator, db_handler):
        super().__init__()
        self.material_manager = material_manager
        self.product_manager = product_manager
        self.cost_calculator = cost_calculator
        self.db_handler = db_handler
        self.excel_handler = ExcelHandler()
        
        self.setWindowTitle("Product Manager")
        self.setGeometry(100, 100, 1400, 800)
        
        self._setup_ui()
        self._connect_signals()
        self._load_initial_data()
    
    def _setup_ui(self):
        """Setup user interface"""
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        main_layout = QVBoxLayout()
        
        # Tab widget
        tabs = QTabWidget()
        
        # Tab 1: Raw materials
        tab1 = QWidget()
        self.tab1_layout = QVBoxLayout()
        self._setup_materials_tab(tab1)
        tabs.addTab(tab1, "Materie Prime")
        
        # Tab 2: Products
        tab2 = QWidget()
        self.tab2_layout = QVBoxLayout()
        self._setup_products_tab(tab2)
        tabs.addTab(tab2, "Prodotti")
        
        main_layout.addWidget(tabs)
        
        # Bottom buttons
        bottom_layout = QHBoxLayout()
        
        export_btn = QPushButton("Esporta in Excel")
        export_btn.clicked.connect(self._export_excel)
        bottom_layout.addWidget(export_btn)
        
        bottom_layout.addStretch()
        
        main_layout.addLayout(bottom_layout)
        central_widget.setLayout(main_layout)
    
    def _setup_materials_tab(self, tab):
        """Setup raw materials panel"""
        layout = QVBoxLayout()
        
        # Input section
        input_layout = QHBoxLayout()
        
        # Name
        input_layout.addWidget(QLabel("Nome:"))
        self.mat_name_input = QLineEdit()
        input_layout.addWidget(self.mat_name_input)
        
        # Unit
        input_layout.addWidget(QLabel("Unità:"))
        self.mat_unit_combo = QComboBox()
        self.mat_unit_combo.addItems(["kg", "L", "m", "pz"])
        input_layout.addWidget(self.mat_unit_combo)
        
        # Cost
        input_layout.addWidget(QLabel("Costo unitario:"))
        self.mat_cost_input = FloatLineEdit()
        input_layout.addWidget(self.mat_cost_input)
        
        # Add button
        add_btn = QPushButton("Aggiungi")
        add_btn.clicked.connect(self._add_material)
        input_layout.addWidget(add_btn)
        
        layout.addLayout(input_layout)
        
        # Materials table
        layout.addWidget(QLabel("Elenco materie prime:"))
        self.materials_table = QTableWidget()
        self.materials_table.setColumnCount(5)
        self.materials_table.setHorizontalHeaderLabels(["ID", "Nome", "Unità", "Costo unitario", "Azione"])
        self.materials_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        layout.addWidget(self.materials_table)
        
        tab.setLayout(layout)
    
    def _setup_products_tab(self, tab):
        """Setup products panel"""
        layout = QHBoxLayout()
        
        # Left side - Products list
        left_layout = QVBoxLayout()
        left_layout.addWidget(QLabel("Prodotti:"))
        
        # Add product section
        add_prod_layout = QHBoxLayout()
        self.prod_name_input = QLineEdit()
        self.prod_name_input.setPlaceholderText("Nome prodotto...")
        add_prod_layout.addWidget(self.prod_name_input)
        
        self.prod_semilavorato_cb = QComboBox()
        self.prod_semilavorato_cb.addItems(["No", "Sì"])
        add_prod_layout.addWidget(QLabel("Semilavorato:"))
        add_prod_layout.addWidget(self.prod_semilavorato_cb)
        
        add_prod_btn = QPushButton("Aggiungi Prodotto")
        add_prod_btn.clicked.connect(self._add_product)
        add_prod_layout.addWidget(add_prod_btn)
        left_layout.addLayout(add_prod_layout)
        
        # Products table
        self.products_table = QTableWidget()
        self.products_table.setColumnCount(4)
        self.products_table.setHorizontalHeaderLabels(["ID", "Nome", "Costo", "Azione"])
        self.products_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.products_table.itemSelectionChanged.connect(self._on_product_selected)
        left_layout.addWidget(self.products_table)
        
        layout.addLayout(left_layout)
        
        # Right side - Composition
        right_layout = QVBoxLayout()
        right_layout.addWidget(QLabel("Composizione del prodotto:"))
        
        # Add component section
        add_comp_layout = QHBoxLayout()
        
        add_comp_layout.addWidget(QLabel("Componente:"))
        self.comp_type_combo = QComboBox()
        self.comp_type_combo.addItems(["Materia Prima", "Semilavorato"])
        self.comp_type_combo.currentIndexChanged.connect(self._on_component_type_changed)
        add_comp_layout.addWidget(self.comp_type_combo)
        
        self.comp_item_combo = QComboBox()
        add_comp_layout.addWidget(self.comp_item_combo)
        
        add_comp_layout.addWidget(QLabel("Quantità:"))
        self.comp_quantity_input = FloatLineEdit()
        add_comp_layout.addWidget(self.comp_quantity_input)
        
        add_comp_btn = QPushButton("Aggiungi")
        add_comp_btn.clicked.connect(self._add_component_to_product)
        add_comp_layout.addWidget(add_comp_btn)
        
        right_layout.addLayout(add_comp_layout)
        
        # Composition table
        self.composition_table = QTableWidget()
        self.composition_table.setColumnCount(5)
        self.composition_table.setHorizontalHeaderLabels(["ID", "Nome", "Quantità", "Tipo", "Azione"])
        self.composition_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        right_layout.addWidget(self.composition_table)
        
        layout.addLayout(right_layout)
        tab.setLayout(layout)
    
    def _connect_signals(self):
        """Connect signals and slots"""
        self.material_updated.connect(self._on_material_updated)
    
    def _load_initial_data(self):
        """Load initial data from database"""
        self._refresh_materials_table()
        self._refresh_products_table()
        self._refresh_component_combos()
    
    def _add_material(self):
        """Add a new raw material"""
        name = self.mat_name_input.text().strip()
        unit = self.mat_unit_combo.currentText()
        cost = self.mat_cost_input.get_value()
        
        if not name:
            QMessageBox.warning(self, "Errore", "Inserisci un nome per la materia prima")
            return
        
        if cost <= 0:
            QMessageBox.warning(self, "Errore", "Il costo deve essere maggiore di 0")
            return
        
        self.material_manager.add(name, unit, cost)
        self.mat_name_input.clear()
        self.mat_cost_input.clear()
        
        self._refresh_materials_table()
        self._refresh_component_combos()
        self.material_updated.emit()
    
    def _refresh_materials_table(self):
        """Refresh materials table"""
        self.materials_table.setRowCount(0)
        
        for material in self.material_manager.get_all():
            row = self.materials_table.rowCount()
            self.materials_table.insertRow(row)
            
            self.materials_table.setItem(row, 0, QTableWidgetItem(material.id))
            self.materials_table.setItem(row, 1, QTableWidgetItem(material.name))
            self.materials_table.setItem(row, 2, QTableWidgetItem(material.unit))
            self.materials_table.setItem(row, 3, QTableWidgetItem(f"{material.cost_per_unit:.2f}"))
            
            delete_btn = QPushButton("Elimina")
            delete_btn.clicked.connect(lambda checked, mid=material.id: self._delete_material(mid))
            self.materials_table.setCellWidget(row, 4, delete_btn)
    
    def _delete_material(self, material_id):
        """Delete a material"""
        reply = QMessageBox.question(self, "Conferma", "Eliminare questa materia prima?")
        if reply == QMessageBox.Yes:
            self.material_manager.delete(material_id)
            self._refresh_materials_table()
            self._refresh_component_combos()
            self.material_updated.emit()
    
    def _add_product(self):
        """Add a new product"""
        name = self.prod_name_input.text().strip()
        is_semilavorato = self.prod_semilavorato_cb.currentText() == "Sì"
        
        if not name:
            QMessageBox.warning(self, "Errore", "Inserisci un nome per il prodotto")
            return
        
        self.product_manager.add(name, is_semilavorato)
        self.prod_name_input.clear()
        
        self._refresh_products_table()
        self._refresh_component_combos()
    
    def _refresh_products_table(self):
        """Refresh products table"""
        self.products_table.setRowCount(0)
        self.cost_calculator.update_all_costs()
        
        for product in self.product_manager.get_all():
            row = self.products_table.rowCount()
            self.products_table.insertRow(row)
            
            self.products_table.setItem(row, 0, QTableWidgetItem(product.id))
            self.products_table.setItem(row, 1, QTableWidgetItem(product.name))
            self.products_table.setItem(row, 2, QTableWidgetItem(f"{product.cost:.2f}"))
            
            delete_btn = QPushButton("Elimina")
            delete_btn.clicked.connect(lambda checked, pid=product.id: self._delete_product(pid))
            self.products_table.setCellWidget(row, 3, delete_btn)
    
    def _delete_product(self, product_id):
        """Delete a product"""
        reply = QMessageBox.question(self, "Conferma", "Eliminare questo prodotto?")
        if reply == QMessageBox.Yes:
            self.product_manager.delete(product_id)
            self._refresh_products_table()
            self._refresh_component_combos()
    
    def _on_product_selected(self):
        """Handle product selection"""
        selected_rows = self.products_table.selectedIndexes()
        if not selected_rows:
            return
        
        row = selected_rows[0].row()
        product_id = self.products_table.item(row, 0).text()
        self._refresh_composition_table(product_id)
    
    def _on_component_type_changed(self):
        """Update component combo based on type"""
        component_type = self.comp_type_combo.currentText()
        self.comp_item_combo.clear()
        
        if component_type == "Materia Prima":
            for material in self.material_manager.get_all():
                self.comp_item_combo.addItem(f"{material.name} ({material.unit})", material.id)
        else:  # Semilavorato
            for product in self.product_manager.get_all():
                if product.is_semilavorato:
                    self.comp_item_combo.addItem(product.name, product.id)
    
    def _refresh_component_combos(self):
        """Refresh component combo boxes"""
        self._on_component_type_changed()
    
    def _add_component_to_product(self):
        """Add component to product composition"""
        selected_rows = self.products_table.selectedIndexes()
        if not selected_rows:
            QMessageBox.warning(self, "Errore", "Seleziona un prodotto prima")
            return
        
        row = selected_rows[0].row()
        product_id = self.products_table.item(row, 0).text()
        
        component_type = "raw_material" if self.comp_type_combo.currentText() == "Materia Prima" else "product"
        component_id = self.comp_item_combo.currentData()
        quantity = self.comp_quantity_input.get_value()
        
        if quantity <= 0:
            QMessageBox.warning(self, "Errore", "La quantità deve essere maggiore di 0")
            return
        
        composition = self.product_manager.get_composition(product_id)
        composition.add_component(component_id, quantity, component_type)
        self.db_handler.save_composition(composition)
        
        self.comp_quantity_input.clear()
        self._refresh_composition_table(product_id)
        self._refresh_products_table()
    
    def _refresh_composition_table(self, product_id):
        """Refresh composition table for selected product"""
        self.composition_table.setRowCount(0)
        
        composition = self.product_manager.get_composition(product_id)
        if not composition:
            return
        
        for component_id, quantity, component_type in composition.get_components():
            row = self.composition_table.rowCount()
            self.composition_table.insertRow(row)
            
            if component_type == "raw_material":
                material = self.material_manager.get(component_id)
                name = material.name if material else "Unknown"
            else:
                product = self.product_manager.get(component_id)
                name = product.name if product else "Unknown"
            
            self.composition_table.setItem(row, 0, QTableWidgetItem(component_id))
            self.composition_table.setItem(row, 1, QTableWidgetItem(name))
            self.composition_table.setItem(row, 2, QTableWidgetItem(f"{quantity}"))
            self.composition_table.setItem(row, 3, QTableWidgetItem(component_type))
            
            delete_btn = QPushButton("Rimuovi")
            delete_btn.clicked.connect(lambda checked, pid=product_id, cid=component_id: self._remove_component(pid, cid))
            self.composition_table.setCellWidget(row, 4, delete_btn)
    
    def _remove_component(self, product_id, component_id):
        """Remove component from product"""
        composition = self.product_manager.get_composition(product_id)
        composition.remove_component(component_id)
        self.db_handler.save_composition(composition)
        
        self._refresh_composition_table(product_id)
        self._refresh_products_table()
    
    def _on_material_updated(self):
        """Handle material update"""
        self._refresh_products_table()
    
    def _export_excel(self):
        """Export data to Excel"""
        filename, _ = QFileDialog.getSaveFileName(self, "Esporta in Excel", "", "Excel Files (*.xlsx)")
        if filename:
            self.excel_handler.export(
                filename,
                self.material_manager.get_all(),
                self.product_manager.get_all(),
                self.product_manager.compositions,
                self.cost_calculator
            )
            QMessageBox.information(self, "Successo", "Dati esportati con successo!")