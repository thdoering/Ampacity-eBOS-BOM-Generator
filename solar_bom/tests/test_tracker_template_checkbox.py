import pytest
import tkinter as tk
from unittest.mock import Mock, MagicMock, patch
import sys
import os

# Add the src directory to the path so we can import our modules
sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(__file__)), 'src'))

from ui.tracker_creator import TrackerTemplateCreator
from models.module import ModuleSpec, ModuleType
from models.tracker import TrackerTemplate, ModuleOrientation


class TestTrackerTemplateCheckbox:
    """Test the tracker template checkbox functionality"""
    
    @pytest.fixture
    def root(self):
        """Create a Tk root window for testing"""
        root = tk.Tk()
        yield root
        root.destroy()
    
    @pytest.fixture
    def mock_project(self):
        """Create a mock project with enabled_templates list"""
        project = Mock()
        project.enabled_templates = []
        return project
    
    @pytest.fixture
    def tracker_creator(self, root, mock_project):
        """Create a TrackerTemplateCreator instance for testing"""
        # Mock the load_templates method to return test templates
        with patch.object(TrackerTemplateCreator, 'load_templates') as mock_load:
            mock_load.return_value = {
                "Test Manufacturer - Test Template 1": {
                    "module_orientation": "portrait",
                    "modules_per_string": 28,
                    "strings_per_tracker": 2,
                    "module_spacing_m": 0.02,
                    "motor_gap_m": 1.0,
                    "motor_position_after_string": 0,
                    "motor_placement_type": "between_strings",
                    "motor_string_index": 1,
                    "motor_split_north": 14,
                    "motor_split_south": 14,
                    "module_spec": {
                        "manufacturer": "Test Manufacturer",
                        "model": "Test Model",
                        "type": "monofacial",
                        "length_mm": 2100,
                        "width_mm": 1050,
                        "depth_mm": 35,
                        "weight_kg": 25,
                        "wattage": 500,
                        "vmp": 40.5,
                        "imp": 12.35,
                        "voc": 48.2,
                        "isc": 13.1,
                        "max_system_voltage": 1500
                    }
                }
            }
            
            creator = TrackerTemplateCreator(
                root,
                current_project=mock_project,
                on_template_enabled_changed=Mock()
            )
            return creator
    
    def test_toggle_item_enabled_incomplete(self, tracker_creator):
        """Test that the current toggle_item_enabled method is incomplete"""
        # Get the first template item from the tree
        # The tree structure is Manufacturer -> Model -> String Size -> Template
        manufacturer_items = tracker_creator.template_tree.get_children()
        assert len(manufacturer_items) > 0, "No manufacturer items in tree"
        
        model_items = tracker_creator.template_tree.get_children(manufacturer_items[0])
        assert len(model_items) > 0, "No model items in tree"
        
        string_items = tracker_creator.template_tree.get_children(model_items[0])
        assert len(string_items) > 0, "No string size items in tree"
        
        template_items = tracker_creator.template_tree.get_children(string_items[0])
        assert len(template_items) > 0, "No template items in tree"
        
        template_item = template_items[0]
        
        # Check initial state - should have checkbox value
        values = tracker_creator.template_tree.item(template_item, 'values')
        assert values[0] in ['☐', '☑'], "Template should have checkbox value"
        initial_checkbox = values[0]
        
        # Try to toggle - this should fail with current implementation
        tracker_creator.toggle_item_enabled(template_item)
        
        # Check if state changed (it shouldn't with current broken implementation)
        values_after = tracker_creator.template_tree.item(template_item, 'values')
        assert values_after[0] == initial_checkbox, "Checkbox state should not change with incomplete implementation"
        
    def test_checkbox_click_event(self, tracker_creator):
        """Test that clicking on the checkbox column triggers toggle_item_enabled"""
        # Mock the identify methods
        tracker_creator.template_tree.identify_row = Mock(return_value="test_item")
        tracker_creator.template_tree.identify_column = Mock(return_value="#1")  # Checkbox column
        
        # Mock toggle_item_enabled to track if it was called
        tracker_creator.toggle_item_enabled = Mock()
        
        # Create a mock event
        event = Mock()
        event.x = 100
        event.y = 50
        
        # Trigger the click event
        tracker_creator.on_tree_click(event)
        
        # Verify toggle_item_enabled was called
        tracker_creator.toggle_item_enabled.assert_called_once_with("test_item")
    
    def test_template_key_mapping(self, tracker_creator):
        """Test that tree_item_to_template mapping is created correctly"""
        # Check that mapping exists
        assert hasattr(tracker_creator, 'tree_item_to_template'), "Should have tree_item_to_template mapping"
        
        # Get a template item
        manufacturer_items = tracker_creator.template_tree.get_children()
        model_items = tracker_creator.template_tree.get_children(manufacturer_items[0])
        string_items = tracker_creator.template_tree.get_children(model_items[0])
        template_items = tracker_creator.template_tree.get_children(string_items[0])
        
        if template_items:
            template_item = template_items[0]
            # Check that this item is in the mapping
            assert template_item in tracker_creator.tree_item_to_template, "Template item should be in mapping"
            # Check that it maps to the correct template key
            assert tracker_creator.tree_item_to_template[template_item] == "Test Manufacturer - Test Template 1"


class TestTrackerTemplateCheckboxFix:
    """Test the fixed checkbox functionality"""
    
    @pytest.fixture
    def root(self):
        """Create a Tk root window for testing"""
        root = tk.Tk()
        yield root
        root.destroy()
    
    @pytest.fixture
    def mock_project(self):
        """Create a mock project with enabled_templates list"""
        project = Mock()
        project.enabled_templates = []
        return project
    
    def test_toggle_enabled_state(self, root, mock_project):
        """Test that toggle_item_enabled properly toggles the checkbox state"""
        # This test will fail until we implement the fix
        # We'll update this after implementing the fix
        pass