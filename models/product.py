class Product:
    """Represents a product (finished or semi-finished)"""
    
    def __init__(self, id, name, is_semilavorato=False):
        self.id = id
        self.name = name
        self.is_semilavorato = is_semilavorato  # True if can be used as input
        self.cost = 0.0  # Calculated dynamically
    
    def __repr__(self):
        return f"Product(id={self.id}, name={self.name}, cost={self.cost})"