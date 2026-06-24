from PySide6.QtWidgets import (QLineEdit, QComboBox, QTableWidget, 
                               QTableWidgetItem, QPushButton, QSpinBox, QDoubleSpinBox)
from PySide6.QtCore import Qt
from .validators import CommaToPointValidator

class FloatLineEdit(QLineEdit):
    """
    QLineEdit that accepts only valid float numbers.
    Automatically converts commas to points.
    """
    
    def __init__(self, parent=None):
        super().__init__(parent)
        validator = CommaToPointValidator()
        self.setValidator(validator)
        self.textChanged.connect(self._on_text_changed)
    
    def _on_text_changed(self, text):
        """Replace comma with point on every change"""
        if ',' in text:
            corrected = text.replace(',', '.')
            self.blockSignals(True)
            self.setText(corrected)
            self.blockSignals(False)
    
    def get_value(self):
        """Get float value from input"""
        try:
            return float(self.text()) if self.text() else 0.0
        except ValueError:
            return 0.0