class RawMaterial:
    """Represents a raw material with cost and unit"""
    
    def __init__(self, id, name, unit, cost_per_unit):
        self.id = id  # Internal identifier
        self.name = name  # "cera", "stoppini", etc.
        self.unit = unit  # "kg", "L", "m", "pz"
        self.cost_per_unit = cost_per_unit  # Float
    
    def __repr__(self):
        return f"RawMaterial(id={self.id}, name={self.name}, unit={self.unit}, cost={self.cost_per_unit})"