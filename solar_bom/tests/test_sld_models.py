"""
Unit tests for SLD data models
"""

import unittest
import sys
from pathlib import Path

# Add src to path
sys.path.insert(0, str(Path(__file__).parent.parent))

from src.models.sld import (
    SLDElement, SLDConnection, SLDDiagram, SLDAnnotation,
    SLDElementType, ConnectionPortType, ConnectionPort
)


class TestSLDModels(unittest.TestCase):
    
    def test_sld_element_creation(self):
        """Test creating an SLD element"""
        element = SLDElement(
            element_id="PV-01",
            element_type=SLDElementType.PV_BLOCK,
            x=100,
            y=200,
            label="Block A - 2.5 MW"
        )
        
        self.assertEqual(element.element_id, "PV-01")
        self.assertEqual(element.element_type, SLDElementType.PV_BLOCK)
        self.assertEqual(element.x, 100)
        self.assertEqual(element.y, 200)
    
    def test_connection_port_creation(self):
        """Test creating connection ports"""
        port = ConnectionPort(
            port_id="dc_pos_1",
            port_type=ConnectionPortType.DC_POSITIVE,
            side="right",
            offset=0.25,
            max_current=30.0
        )
        
        self.assertEqual(port.port_type, ConnectionPortType.DC_POSITIVE)
        self.assertEqual(port.side, "right")
    
    def test_sld_connection_creation(self):
        """Test creating connections between elements"""
        connection = SLDConnection(
            connection_id="CONN-01",
            from_element="PV-01",
            from_port="dc_pos",
            to_element="INV-01",
            to_port="dc_pos_input",
            cable_size="4/0 AWG"
        )
        
        self.assertEqual(connection.cable_size, "4/0 AWG")
        self.assertEqual(connection.cable_type, "DC")
    
    def test_orthogonal_path_calculation(self):
        """Test orthogonal path calculation"""
        connection = SLDConnection(
            connection_id="CONN-01",
            from_element="PV-01",
            from_port="dc_pos",
            to_element="INV-01",
            to_port="dc_pos_input"
        )
        
        path = connection.calculate_orthogonal_path((100, 100), (300, 200))
        
        # Should have 3 points for a right-angle turn
        self.assertEqual(len(path), 3)
        self.assertEqual(path[0], (100, 100))  # Start point
        self.assertEqual(path[1], (300, 100))  # Turn point
        self.assertEqual(path[2], (300, 200))  # End point
    
    def test_sld_diagram_creation(self):
        """Test creating a complete SLD diagram"""
        diagram = SLDDiagram(
            project_id="test_project",
            diagram_name="Test SLD"
        )
        
        # Add PV block
        pv_element = SLDElement(
            element_id="PV-01",
            element_type=SLDElementType.PV_BLOCK,
            x=100,
            y=100,
            label="Block A"
        )
        
        # Add ports to PV block
        pv_element.ports.append(ConnectionPort(
            port_id="dc_pos",
            port_type=ConnectionPortType.DC_POSITIVE,
            side="right",
            offset=0.3
        ))
        pv_element.ports.append(ConnectionPort(
            port_id="dc_neg",
            port_type=ConnectionPortType.DC_NEGATIVE,
            side="right",
            offset=0.7
        ))
        
        diagram.add_element(pv_element)
        
        # Add inverter
        inv_element = SLDElement(
            element_id="INV-01",
            element_type=SLDElementType.INVERTER,
            x=400,
            y=100,
            label="Inverter 1"
        )
        
        # Add ports to inverter
        inv_element.ports.append(ConnectionPort(
            port_id="dc_pos_input",
            port_type=ConnectionPortType.DC_POSITIVE,
            side="left",
            offset=0.3
        ))
        inv_element.ports.append(ConnectionPort(
            port_id="dc_neg_input",
            port_type=ConnectionPortType.DC_NEGATIVE,
            side="left",
            offset=0.7
        ))
        
        diagram.add_element(inv_element)
        
        self.assertEqual(len(diagram.elements), 2)
        self.assertIsNotNone(diagram.get_element("PV-01"))
        self.assertIsNotNone(diagram.get_element("INV-01"))
    
    def test_connection_validation(self):
        """Test connection validation logic"""
        diagram = SLDDiagram(
            project_id="test_project",
            diagram_name="Test SLD"
        )
        
        # Create two elements with ports
        pv = SLDElement(
            element_id="PV-01",
            element_type=SLDElementType.PV_BLOCK,
            x=100, y=100
        )
        pv.ports.append(ConnectionPort(
            port_id="dc_pos",
            port_type=ConnectionPortType.DC_POSITIVE,
            side="right",
            offset=0.5
        ))
        
        inv = SLDElement(
            element_id="INV-01",
            element_type=SLDElementType.INVERTER,
            x=400, y=100
        )
        inv.ports.append(ConnectionPort(
            port_id="dc_pos_input",
            port_type=ConnectionPortType.DC_POSITIVE,
            side="left",
            offset=0.5
        ))
        inv.ports.append(ConnectionPort(
            port_id="dc_neg_input",
            port_type=ConnectionPortType.DC_NEGATIVE,
            side="left",
            offset=0.7
        ))
        
        diagram.add_element(pv)
        diagram.add_element(inv)
        
        # Valid connection (DC+ to DC+)
        valid, msg = diagram.validate_connection(
            "PV-01", "dc_pos",
            "INV-01", "dc_pos_input"
        )
        self.assertTrue(valid)
        
        # Invalid connection (DC+ to DC-)
        valid, msg = diagram.validate_connection(
            "PV-01", "dc_pos",
            "INV-01", "dc_neg_input"
        )
        self.assertFalse(valid)
        self.assertIn("positive", msg.lower())
    
    def test_auto_layout(self):
        """Test automatic layout functionality"""
        diagram = SLDDiagram(project_id="test")
        
        # Add multiple PV blocks
        for i in range(3):
            diagram.add_element(SLDElement(
                element_id=f"PV-{i+1}",
                element_type=SLDElementType.PV_BLOCK,
                x=0, y=0  # Start at origin
            ))
        
        # Add multiple inverters
        for i in range(2):
            diagram.add_element(SLDElement(
                element_id=f"INV-{i+1}",
                element_type=SLDElementType.INVERTER,
                x=0, y=0  # Start at origin
            ))
        
        # Run auto-layout
        diagram.auto_layout()
        
        # Check PV blocks are in left zone
        for element in diagram.elements:
            if element.element_type == SLDElementType.PV_BLOCK:
                self.assertGreaterEqual(element.x, diagram.pv_zone_x_start)
                self.assertLessEqual(element.x, diagram.pv_zone_x_end)
            elif element.element_type == SLDElementType.INVERTER:
                self.assertGreaterEqual(element.x, diagram.inverter_zone_x_start)
                self.assertLessEqual(element.x, diagram.inverter_zone_x_end)
        
        # Check vertical spacing
        pv_blocks = [e for e in diagram.elements 
                     if e.element_type == SLDElementType.PV_BLOCK]
        if len(pv_blocks) > 1:
            # Elements should have different y positions
            y_positions = [e.y for e in pv_blocks]
            self.assertEqual(len(y_positions), len(set(y_positions)))
    
    def test_serialization(self):
        """Test to_dict and from_dict methods"""
        # Create a diagram with elements
        diagram = SLDDiagram(project_id="test")
        
        element = SLDElement(
            element_id="PV-01",
            element_type=SLDElementType.PV_BLOCK,
            x=100,
            y=200,
            label="Test Block",
            power_kw=2500
        )
        element.ports.append(ConnectionPort(
            port_id="dc_pos",
            port_type=ConnectionPortType.DC_POSITIVE,
            side="right",
            offset=0.5
        ))
        
        diagram.add_element(element)
        
        connection = SLDConnection(
            connection_id="CONN-01",
            from_element="PV-01",
            from_port="dc_pos",
            to_element="INV-01",
            to_port="dc_input",
            cable_size="2/0 AWG"
        )
        diagram.add_connection(connection)
        
        annotation = SLDAnnotation(
            annotation_id="ANN-01",
            text="2.5 MW",
            x=100,
            y=180,
            element_id="PV-01"
        )
        diagram.annotations.append(annotation)
        
        # Serialize to dict
        data = diagram.to_dict()
        
        # Deserialize from dict
        diagram2 = SLDDiagram.from_dict(data)
        
        # Verify
        self.assertEqual(diagram2.project_id, "test")
        self.assertEqual(len(diagram2.elements), 1)
        self.assertEqual(len(diagram2.connections), 1)
        self.assertEqual(len(diagram2.annotations), 1)
        
        element2 = diagram2.elements[0]
        self.assertEqual(element2.element_id, "PV-01")
        self.assertEqual(element2.label, "Test Block")
        self.assertEqual(element2.power_kw, 2500)
        self.assertEqual(len(element2.ports), 1)
    
    def test_element_removal(self):
        """Test removing elements and their connections"""
        diagram = SLDDiagram(project_id="test")
        
        # Add elements
        diagram.add_element(SLDElement(
            element_id="PV-01",
            element_type=SLDElementType.PV_BLOCK,
            x=100, y=100
        ))
        diagram.add_element(SLDElement(
            element_id="INV-01",
            element_type=SLDElementType.INVERTER,
            x=400, y=100
        ))
        
        # Add connection
        diagram.add_connection(SLDConnection(
            connection_id="CONN-01",
            from_element="PV-01",
            from_port="dc_pos",
            to_element="INV-01",
            to_port="dc_input"
        ))
        
        # Add annotation
        diagram.annotations.append(SLDAnnotation(
            annotation_id="ANN-01",
            text="Test",
            x=100, y=80,
            element_id="PV-01"
        ))
        
        self.assertEqual(len(diagram.elements), 2)
        self.assertEqual(len(diagram.connections), 1)
        self.assertEqual(len(diagram.annotations), 1)
        
        # Remove PV-01
        removed = diagram.remove_element("PV-01")
        
        self.assertTrue(removed)
        self.assertEqual(len(diagram.elements), 1)
        self.assertEqual(len(diagram.connections), 0)  # Connection should be removed
        self.assertEqual(len(diagram.annotations), 0)  # Annotation should be removed
    
    def test_get_port_position(self):
        """Test calculating absolute port positions"""
        element = SLDElement(
            element_id="PV-01",
            element_type=SLDElementType.PV_BLOCK,
            x=100,
            y=200,
            width=80,
            height=60
        )
        
        # Add ports on different sides
        element.ports.extend([
            ConnectionPort("top_port", ConnectionPortType.DC_POSITIVE, "top", 0.5),
            ConnectionPort("bottom_port", ConnectionPortType.DC_NEGATIVE, "bottom", 0.5),
            ConnectionPort("left_port", ConnectionPortType.DC_POSITIVE, "left", 0.5),
            ConnectionPort("right_port", ConnectionPortType.DC_NEGATIVE, "right", 0.5)
        ])
        
        # Test port positions
        top_pos = element.get_port_position("top_port")
        self.assertEqual(top_pos, (140, 200))  # x + width*0.5, y
        
        bottom_pos = element.get_port_position("bottom_port")
        self.assertEqual(bottom_pos, (140, 260))  # x + width*0.5, y + height
        
        left_pos = element.get_port_position("left_port")
        self.assertEqual(left_pos, (100, 230))  # x, y + height*0.5
        
        right_pos = element.get_port_position("right_port")
        self.assertEqual(right_pos, (180, 230))  # x + width, y + height*0.5


if __name__ == '__main__':
    unittest.main()