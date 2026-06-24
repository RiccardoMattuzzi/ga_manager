class Composition:
    """Represents the composition of a product"""
    
    def __init__(self, product_id):
        self.product_id = product_id
        self.items = []  # List of (component_id, quantity, component_type)
        # component_type = "raw_material" or "product" (for semi-finished)
    
    def add_component(self, component_id, quantity, component_type):
        """Add a component to the composition"""
        self.items.append((component_id, quantity, component_type))
    
    def remove_component(self, component_id):
        """Remove a component from the composition"""
        self.items = [(cid, q, ct) for cid, q, ct in self.items if cid != component_id]
    
    def update_quantity(self, component_id, quantity):
        """Update quantity of a component"""
        for i, (cid, q, ct) in enumerate(self.items):
            if cid == component_id:
                self.items[i] = (cid, quantity, ct)
                break
    
    def get_components(self):
        """Return list of (component_id, quantity, component_type)"""
        return self.items