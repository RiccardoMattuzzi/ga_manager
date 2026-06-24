from models import RawMaterial

class MaterialManager:
    """Manages all raw materials"""
    
    def __init__(self, db_handler):
        self.db = db_handler
        self.materials = {}  # id -> RawMaterial
        self.next_id = 1
    
    def load_all(self):
        """Load all materials from database"""
        self.materials = self.db.load_materials()
        if self.materials:
            self.next_id = max(int(m.id) for m in self.materials.values()) + 1
    
    def add(self, name, unit, cost_per_unit):
        """Add a new material and return its id"""
        material_id = str(self.next_id)
        self.next_id += 1
        material = RawMaterial(material_id, name, unit, cost_per_unit)
        self.materials[material_id] = material
        self.db.save_material(material)
        return material_id
    
    def update(self, material_id, name=None, unit=None, cost_per_unit=None):
        """Update an existing material"""
        if material_id not in self.materials:
            return False
        
        material = self.materials[material_id]
        if name is not None:
            material.name = name
        if unit is not None:
            material.unit = unit
        if cost_per_unit is not None:
            material.cost_per_unit = cost_per_unit
        
        self.db.save_material(material)
        return True
    
    def get(self, material_id):
        """Get a material by id"""
        return self.materials.get(material_id)
    
    def get_all(self):
        """Get all materials"""
        return list(self.materials.values())
    
    def delete(self, material_id):
        """Delete a material"""
        if material_id in self.materials:
            del self.materials[material_id]
            self.db.delete_material(material_id)
            return True
        return False