import sqlite3
from pathlib import Path
from models import RawMaterial, Product, Composition

class DatabaseHandler:
    """Handles all database operations with SQLite"""
    
    def __init__(self, db_path):
        self.db_path = db_path
        self.conn = None
        self.setup_connection()
        self.setup_tables()
    
    def setup_connection(self):
        """Create database connection"""
        self.conn = sqlite3.connect(str(self.db_path))
        self.conn.row_factory = sqlite3.Row
    
    def setup_tables(self):
        """Create tables if they don't exist"""
        cursor = self.conn.cursor()
        
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS raw_materials (
                id TEXT PRIMARY KEY,
                name TEXT NOT NULL,
                unit TEXT NOT NULL,
                cost_per_unit REAL NOT NULL
            )
        ''')
        
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS products (
                id TEXT PRIMARY KEY,
                name TEXT NOT NULL,
                is_semilavorato BOOLEAN NOT NULL DEFAULT 0
            )
        ''')
        
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS compositions (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                product_id TEXT NOT NULL,
                component_id TEXT NOT NULL,
                quantity REAL NOT NULL,
                component_type TEXT NOT NULL,
                FOREIGN KEY(product_id) REFERENCES products(id)
            )
        ''')
        
        self.conn.commit()
    
    def load_materials(self):
        """Load all raw materials from database"""
        cursor = self.conn.cursor()
        cursor.execute('SELECT * FROM raw_materials')
        materials = {}
        for row in cursor.fetchall():
            material = RawMaterial(row['id'], row['name'], row['unit'], row['cost_per_unit'])
            materials[row['id']] = material
        return materials
    
    def save_material(self, material):
        """Save or update a raw material"""
        cursor = self.conn.cursor()
        cursor.execute('''
            INSERT OR REPLACE INTO raw_materials (id, name, unit, cost_per_unit)
            VALUES (?, ?, ?, ?)
        ''', (material.id, material.name, material.unit, material.cost_per_unit))
        self.conn.commit()
    
    def delete_material(self, material_id):
        """Delete a raw material"""
        cursor = self.conn.cursor()
        cursor.execute('DELETE FROM raw_materials WHERE id = ?', (material_id,))
        self.conn.commit()
    
    def load_products(self):
        """Load all products from database"""
        cursor = self.conn.cursor()
        cursor.execute('SELECT * FROM products')
        products = {}
        for row in cursor.fetchall():
            product = Product(row['id'], row['name'], bool(row['is_semilavorato']))
            products[row['id']] = product
        return products
    
    def save_product(self, product):
        """Save or update a product"""
        cursor = self.conn.cursor()
        cursor.execute('''
            INSERT OR REPLACE INTO products (id, name, is_semilavorato)
            VALUES (?, ?, ?)
        ''', (product.id, product.name, product.is_semilavorato))
        self.conn.commit()
    
    def delete_product(self, product_id):
        """Delete a product and its compositions"""
        cursor = self.conn.cursor()
        cursor.execute('DELETE FROM compositions WHERE product_id = ?', (product_id,))
        cursor.execute('DELETE FROM products WHERE id = ?', (product_id,))
        self.conn.commit()
    
    def load_compositions(self):
        """Load all compositions from database"""
        cursor = self.conn.cursor()
        cursor.execute('SELECT * FROM compositions')
        compositions = {}
        for row in cursor.fetchall():
            product_id = row['product_id']
            if product_id not in compositions:
                compositions[product_id] = Composition(product_id)
            compositions[product_id].add_component(
                row['component_id'],
                row['quantity'],
                row['component_type']
            )
        return compositions
    
    def save_composition(self, composition):
        """Save or update a composition"""
        cursor = self.conn.cursor()
        cursor.execute('DELETE FROM compositions WHERE product_id = ?', (composition.product_id,))
        
        for component_id, quantity, component_type in composition.get_components():
            cursor.execute('''
                INSERT INTO compositions (product_id, component_id, quantity, component_type)
                VALUES (?, ?, ?, ?)
            ''', (composition.product_id, component_id, quantity, component_type))
        
        self.conn.commit()
    
    def close(self):
        """Close database connection"""
        if self.conn:
            self.conn.close()