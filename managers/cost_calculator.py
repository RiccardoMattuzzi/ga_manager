class CostCalculator:
    """Calculates product costs based on compositions"""
    
    def __init__(self, material_manager, product_manager):
        self.materials = material_manager
        self.products = product_manager
        self.cache = {}  # Cache for calculated costs
    
    def calculate_product_cost(self, product_id):
        """
        Calculate total cost of a product recursively.
        Handles both raw materials and semi-finished products.
        """
        composition = self.products.get_composition(product_id)
        if not composition:
            return 0.0
        
        total_cost = 0.0
        
        for component_id, quantity, component_type in composition.get_components():
            if component_type == "raw_material":
                material = self.materials.get(component_id)
                if material:
                    total_cost += material.cost_per_unit * quantity
            
            elif component_type == "product":
                # Semi-finished product - recursive calculation
                product_cost = self.calculate_product_cost(component_id)
                total_cost += product_cost * quantity
        
        return round(total_cost, 2)
    
    def update_all_costs(self):
        """Recalculate costs for all products"""
        for product in self.products.get_all():
            product.cost = self.calculate_product_cost(product.id)