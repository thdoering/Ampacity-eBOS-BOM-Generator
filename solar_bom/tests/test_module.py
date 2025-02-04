import unittest
import os
import sys
from pathlib import Path

# Add the project root directory to Python path
project_root = Path(__file__).parent.parent
sys.path.append(str(project_root))

from src.models.module import ModuleSpec
from src.utils.file_handlers import parse_pan_file

class TestModuleSpec(unittest.TestCase):
    def test_valid_module_creation(self):
        """Test creating a ModuleSpec with valid parameters"""
        try:
            module = ModuleSpec(
                model="Test Module",
                width=1.303,
                isc=18.42,
                imp=17.34,
                wattage=600.0,
                voc=41.7,
                vmp=34.6
            )
            
            # Verify all attributes
            self.assertEqual(module.model, "Test Module")
            self.assertEqual(module.width, 1.303)
            self.assertEqual(module.isc, 18.42)
            self.assertEqual(module.imp, 17.34)
            self.assertEqual(module.wattage, 600.0)
            self.assertEqual(module.voc, 41.7)
            self.assertEqual(module.vmp, 34.6)
            
        except Exception as e:
            self.fail(f"ModuleSpec creation raised exception: {str(e)}")

    def test_invalid_module_parameters(self):
        """Test ModuleSpec validation with invalid parameters"""
        # Test negative width
        with self.assertRaises(ValueError):
            ModuleSpec(
                model="Test Module",
                width=-1.0,
                isc=18.42,
                imp=17.34,
                wattage=600.0,
                voc=41.7,
                vmp=34.6
            )
        
        # Test zero current
        with self.assertRaises(ValueError):
            ModuleSpec(
                model="Test Module",
                width=1.303,
                isc=0,
                imp=17.34,
                wattage=600.0,
                voc=41.7,
                vmp=34.6
            )
        
        # Test Imp > Isc (invalid)
        with self.assertRaises(ValueError):
            ModuleSpec(
                model="Test Module",
                width=1.303,
                isc=18.42,
                imp=20.0,  # Greater than Isc
                wattage=600.0,
                voc=41.7,
                vmp=34.6
            )
        
        # Test Vmp > Voc (invalid)
        with self.assertRaises(ValueError):
            ModuleSpec(
                model="Test Module",
                width=1.303,
                isc=18.42,
                imp=17.34,
                wattage=600.0,
                voc=41.7,
                vmp=45.0  # Greater than Voc
            )

class TestPanFileParser(unittest.TestCase):
    def setUp(self):
        """Create a sample PAN file content for testing"""
        self.sample_pan_content = """
PVObject_=pvModule
Version=8.0.3
Flags=$00800743
PVObject_Commercial=pvCommercial
Manufacturer=Trina solar
Model=TSM-DEG-20C-20-600 Vertex
Width=1.303
Height=2.172
Depth=0.033
Weight=34.900
NPieces=100
Technol=mtSiMono
NCelS=60
NCelP=2
NDiode=3
SubModuleLayout=slTwinHalfCells
FrontSurface=fsARCoating
GRef=1000
TRef=25.0
PNom=600.0
Isc=18.420
Voc=41.70
Imp=17.340
Vmp=34.60
End of PVObject pvModule
"""

    def test_parse_valid_pan_file(self):
        """Test parsing valid PAN file content"""
        try:
            module = parse_pan_file(self.sample_pan_content)
            
            # Verify parsed values
            self.assertEqual(module.model, "TSM-DEG-20C-20-600 Vertex")
            self.assertEqual(module.width, 1.303)
            self.assertEqual(module.isc, 18.420)
            self.assertEqual(module.imp, 17.340)
            self.assertEqual(module.wattage, 600.0)
            self.assertEqual(module.voc, 41.70)
            self.assertEqual(module.vmp, 34.60)
            
        except Exception as e:
            self.fail(f"PAN file parsing raised exception: {str(e)}")

    def test_parse_invalid_pan_file(self):
        """Test parsing invalid PAN file content"""
        # Test missing required field
        invalid_content = """
PVObject_=pvModule
Model=Test Module
Width=1.303
# Missing other required fields
End of PVObject pvModule
"""
        with self.assertRaises(ValueError):
            parse_pan_file(invalid_content)
        
        # Test invalid numeric value
        invalid_content = """
PVObject_=pvModule
Model=Test Module
Width=invalid
Isc=18.420
Imp=17.340
PNom=600.0
Voc=41.70
Vmp=34.60
End of PVObject pvModule
"""
        with self.assertRaises(ValueError):
            parse_pan_file(invalid_content)

def run_tests():
    """Run all tests"""
    unittest.main(argv=[''], verbosity=2)

if __name__ == '__main__':
    run_tests()