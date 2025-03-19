def parse_pan_file(content: str) -> dict:
    """
    Parse PVsyst .pan file content into dictionary of parameters
    
    Args:
        content (str): Raw content of .pan file
        
    Returns:
        dict: Dictionary of module parameters
        
    Raises:
        ValueError: If required fields are missing or invalid
    """
    params = {}
    current_section = None
    
    # Parse raw .pan file contents
    for line in content.splitlines():
        line = line.strip()
        if not line or line.startswith('End of'):
            continue
            
        if line.startswith('PVObject_'):
            current_section = line.split('=')[1]
            continue
            
        if '=' not in line:
            continue
            
        key, value = line.split('=', 1)
        key = key.strip()
        value = value.strip()
        
        # Convert numeric values
        try:
            if '.' in value:
                value = float(value)
            elif value.isdigit():
                value = int(value)
        except ValueError:
            pass
            
        params[key] = value
    
    # Define required fields for validation
    required_fields = ['Model', 'Width', 'Isc', 'Imp', 'PNom', 'Voc', 'Vmp']
    
    # Check for missing required fields
    missing_fields = [field for field in required_fields if field not in params]
    if missing_fields:
        raise ValueError(f"Missing required parameters in PAN file: {', '.join(missing_fields)}")
            
    # Return processed parameters
    try:
        return {
            'manufacturer': params.get('Manufacturer', ''),
            'model': params.get('Model', ''),
            'length_mm': float(params.get('Height', 0)) * 1000,  # Convert m to mm
            'width_mm': float(params.get('Width', 0)) * 1000,    # Convert m to mm
            'depth_mm': float(params.get('Depth', 0.04)) * 1000, # Convert m to mm
            'weight_kg': float(params.get('Weight', 25)),
            'wattage': float(params.get('PNom', 0)),
            'vmp': float(params.get('Vmp', 0)),
            'imp': float(params.get('Imp', 0)),
            'voc': float(params.get('Voc', 0)),
            'isc': float(params.get('Isc', 0)),
            'max_system_voltage': float(params.get('VMaxIEC', 1500)),
            'efficiency': None,  # Not directly available in PAN file
            'temperature_coefficient': float(params.get('muPmpReq', 0)) if 'muPmpReq' in params else None
        }
    except (ValueError, TypeError) as e:
        raise ValueError(f"Invalid parameter value in PAN file: {str(e)}")