from models import Product, Composition

class ProductManager:
    """Manages all products and their compositions"""
    
    def __init__(self, db_handler):
        self.db = db_handler
        self.products = {}  # id -> Product
        self.compositions = {}  # product_id -> Composition
        self.next_id = 1
    
    def load_all(self):
        """Load all products and compositions from database"""
        self.products = self.db.load_products()
        self.compositions = self.db.load_compositions()
        if self.products:
            self.next_id = max(int(p.id) for p in self.products.values()) + 1
    
    def add(self, name, is_semilavorato=False):
        """Add a new product and return its id"""
        product_id = str(self.next_id)
        self.next_id += 1
        product = Product(product_id, name, is_semilavorato)
        self.products[product_id] = product
        self.compositions[product_id] = Composition(product_id)
        self.db.save_product(product)
        return product_id
    
    def update(self, product_id, name=None, is_semilavorato=None):
        """Update an existing product"""
        if product_id not in self.products:
            return False
        
        product = self.products[product_id]
        if name is not None:
            product.name = name
        if is_semilavorato is not None:
            product.is_semilavorato = is_semilavorato
        
        self.db.save_product(product)
        return True
    
    def get(self, product_id):
        """Get a product by id"""
        return self.products.get(product_id)
    
    def get_all(self):
        """Get all products"""
        return list(self.products.values())
    
    def get_composition(self, product_id):
        """Get composition of a product"""
        return self.compositions.get(product_id)
    
    def delete(self, product_id):
        """Delete a product"""
        if product_id in self.products:
            del self.products[product_id]
            del self.compositions[product_id]
            self.db.delete_product(product_id)
            return True
        return False