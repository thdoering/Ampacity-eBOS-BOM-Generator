from typing import Dict, Any, Optional
import json
import os
from ..models.module import ModuleSpec
from ..models.inverter import InverterSpec
from ..utils.pan_parser import parse_pan_file as _parse_pan_file
from ..models.module import ModuleSpec, ModuleType

def parse_pan_file(content: str) -> ModuleSpec:
    """
    Parse PAN file content to extract module specifications.
    
    Args:
        content (str): Raw content of .pan file
        
    Returns:
        ModuleSpec: Module specification object
        
    Raises:
        ValueError: If required fields are missing or invalid
    """
    try:
        # Use the implementation from pan_parser.py 
        params = _parse_pan_file(content)
        
        # Create and return a ModuleSpec object
        return ModuleSpec(
            manufacturer=params['manufacturer'],
            model=params['model'],
            type=ModuleType.MONO_PERC,  # Default type
            length_mm=params['length_mm'],
            width_mm=params['width_mm'],
            depth_mm=params['depth_mm'],
            weight_kg=params['weight_kg'],
            wattage=params['wattage'],
            vmp=params['vmp'],
            imp=params['imp'],
            voc=params['voc'],
            isc=params['isc'],
            max_system_voltage=params['max_system_voltage'],
            efficiency=params['efficiency'],
            temperature_coefficient=params['temperature_coefficient']
        )
    except ValueError as e:
        raise ValueError(f"Failed to parse PAN file: {str(e)}")
    except KeyError as e:
        raise ValueError(f"Missing required parameter in PAN file: {str(e)}")

def parse_ond_file(content: str) -> InverterSpec:
    """
    Parse OND file content to extract inverter specifications.
    
    Args:
        content (str): Raw content of .ond file
        
    Returns:
        InverterSpec: Inverter specification object
        
    Raises:
        ValueError: If required fields are missing or invalid
    """
    # Note: This is a placeholder implementation. Actual OND file format
    # needs to be specified for proper implementation
    
    # Initialize parameters dictionary
    params = {}
    
    # Split content into lines and process
    lines = content.split('\n')
    
    # Parameter mapping from OND file to our names
    param_mapping = {
        'Model': 'model',
        'Manufacturer': 'manufacturer',
        'Vmax': 'max_voltage',
        'Vmin_mppt': 'min_voltage',
        'Imax_per_mppt': 'max_current_per_mppt',
        'Num_mppt': 'num_mppt',
        'Pmax': 'max_power',
        'Efficiency': 'efficiency'
    }
    
    # Extract parameters (placeholder implementation)
    for line in lines:
        line = line.strip()
        if ':' in line:  # Assuming OND uses : as separator
            key, value = line.split(':', 1)
            key = key.strip()
            value = value.strip()
            
            if key in param_mapping:
                params[param_mapping[key]] = value
    
    # Add default efficiency if not specified
    if 'efficiency' not in params:
        params['efficiency'] = '0.98'  # 98% default efficiency
    
    try:
        return InverterSpec(
            model=params.get('model', 'Unknown'),
            manufacturer=params.get('manufacturer', 'Unknown'),
            max_voltage=float(params['max_voltage']),
            min_voltage=float(params['min_voltage']),
            max_current_per_mppt=float(params['max_current_per_mppt']),
            num_mppt=int(params['num_mppt']),
            max_power=float(params['max_power']),
            efficiency=float(params['efficiency'])
        )
    except KeyError as e:
        raise ValueError(f"Missing required parameter: {str(e)}")
    except ValueError as e:
        raise ValueError(f"Invalid parameter value: {str(e)}")

def save_json_file(data: Dict[str, Any], filepath: str, create_dirs: bool = True) -> None:
    """
    Save data to JSON file.
    
    Args:
        data: Dictionary to save
        filepath: Path to save file
        create_dirs: Whether to create directories if they don't exist
        
    Raises:
        IOError: If file cannot be written
    """
    try:
        if create_dirs:
            os.makedirs(os.path.dirname(filepath), exist_ok=True)
        
        with open(filepath, 'w') as f:
            json.dump(data, f, indent=4)
    except Exception as e:
        raise IOError(f"Failed to save file: {str(e)}")

def load_json_file(filepath: str) -> Dict[str, Any]:
    """
    Load data from JSON file.
    
    Args:
        filepath: Path to JSON file
        
    Returns:
        Dictionary containing file contents
        
    Raises:
        IOError: If file cannot be read
        json.JSONDecodeError: If file contains invalid JSON
    """
    try:
        with open(filepath, 'r') as f:
            return json.load(f)
    except json.JSONDecodeError as e:
        raise json.JSONDecodeError(f"Invalid JSON file: {str(e)}", e.doc, e.pos)
    except Exception as e:
        raise IOError(f"Failed to load file: {str(e)}")

def ensure_data_directory() -> None:
    """
    Ensure all required data directories exist.
    Creates directories if they don't exist.
    """
    directories = [
        'data',
        'data/blocks',
        'data/templates'
    ]
    
    for directory in directories:
        os.makedirs(directory, exist_ok=True)

def cleanup_filename(filename: str) -> str:
    """
    Clean up filename to ensure it's valid.
    
    Args:
        filename: Original filename
        
    Returns:
        Cleaned filename
    """
    # Replace invalid characters
    invalid_chars = '<>:"/\\|?*'
    for char in invalid_chars:
        filename = filename.replace(char, '_')
    
    # Ensure filename isn't too long
    max_length = 255
    name, ext = os.path.splitext(filename)
    if len(filename) > max_length:
        return name[:max_length-len(ext)] + ext
    
    return filename

def get_file_extension(filepath: str) -> Optional[str]:
    """
    Get file extension without the dot.
    
    Args:
        filepath: Path to file
        
    Returns:
        File extension or None if no extension
    """
    ext = os.path.splitext(filepath)[1]
    return ext[1:] if ext else None

def validate_file_type(filepath: str, allowed_extensions: list) -> bool:
    """
    Validate file has allowed extension.
    
    Args:
        filepath: Path to file
        allowed_extensions: List of allowed extensions without dots
        
    Returns:
        True if file type is allowed, False otherwise
    """
    ext = get_file_extension(filepath)
    return ext and ext.lower() in [x.lower() for x in allowed_extensions]

def get_app_base_path():
    """Get base path for the application, works in both script and frozen modes"""
    import sys
    import os
    if getattr(sys, 'frozen', False):
        # Running as compiled exe
        return os.path.dirname(sys.executable)
    else:
        # Running as script
        return os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))