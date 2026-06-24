import sys
from PySide6.QtWidgets import QApplication
from config import DB_PATH
from io import DatabaseHandler
from managers import MaterialManager, ProductManager, CostCalculator
from gui import MainWindow

def main():
    # Initialize database
    db_handler = DatabaseHandler(DB_PATH)
    
    # Initialize managers
    material_manager = MaterialManager(db_handler)
    product_manager = ProductManager(db_handler)
    cost_calculator = CostCalculator(material_manager, product_manager)
    
    # Load data from database
    material_manager.load_all()
    product_manager.load_all()
    
    # Create GUI
    app = QApplication(sys.argv)
    window = MainWindow(material_manager, product_manager, cost_calculator, db_handler)
    window.show()
    
    sys.exit(app.exec())

if __name__ == "__main__":
    main()