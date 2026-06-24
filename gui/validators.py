from PySide6.QtCore import Qt
from PySide6.QtGui import QValidator

class CommaToPointValidator(QValidator):
    """
    Validator that automatically replaces commas with points.
    Allows only valid float numbers.
    """
    
    def validate(self, string, pos):
        # Replace comma with point
        corrected = string.replace(',', '.')
        
        # Allow empty string
        if not corrected:
            return QValidator.Acceptable, corrected, pos
        
        # Allow single point
        if corrected.count('.') > 1:
            return QValidator.Invalid, corrected, pos
        
        # Check if it starts with point
        if corrected.startswith('.'):
            return QValidator.Intermediate, corrected, pos
        
        # Try to parse as float
        try:
            float(corrected)
            return QValidator.Acceptable, corrected, pos
        except ValueError:
            return QValidator.Invalid, corrected, pos
    
    def fixup(self, string):
        return string.replace(',', '.')