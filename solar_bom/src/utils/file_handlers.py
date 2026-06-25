from typing import Dict, Any, Optional
import json
import os
from ..models.module import ModuleSpec, ModuleType
from ..models.inverter import InverterSpec
from ..utils.pan_parser import parse_pan_file as _parse_pan_file

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
        # Handle backward compatibility for temperature coefficient
        temp_coeff_pmax = params.get('temperature_coefficient_pmax')
        if temp_coeff_pmax is None:
            # Fall back to old single temperature_coefficient field
            temp_coeff_pmax = params.get('temperature_coefficient')
            
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
            temperature_coefficient_pmax=temp_coeff_pmax,
            temperature_coefficient_voc=params.get('temperature_coefficient_voc'),
            temperature_coefficient_isc=params.get('temperature_coefficient_isc')
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
        NotImplementedError: OND file parsing is not yet implemented
    """
    # OND file parsing is not yet implemented.
    # The InverterSpec dataclass requires fields like inverter_type, rated_power_kw,
    # mppt_channels, etc. that need proper mapping from the OND file format.
    raise NotImplementedError(
        "OND file parsing is not yet supported. "
        "Please add inverters manually using the Inverter Manager."
    )

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


def get_user_data_dir():
    """Get the per-user writable base directory for user data and projects.

    Frozen (shipped exe): %APPDATA%\\Solar eBOS BOM Generator, so a user's data
    and projects persist across versions regardless of where the exe lives.
    Dev (script): the repo root, so the existing dev workflow and gitignored dev
    files are unchanged. The directory is created if it doesn't exist.
    """
    import sys
    if getattr(sys, 'frozen', False):
        appdata = os.environ.get('APPDATA') or os.path.expanduser('~')
        base = os.path.join(appdata, 'Solar eBOS BOM Generator')
    else:
        base = get_app_base_path()
    os.makedirs(base, exist_ok=True)
    return base


def get_user_data_path(filename):
    """Full path to a user-writable data file under <user_data_dir>/data/.

    Use for runtime-editable libraries/templates/pricing (NOT read-only factory
    or NEC reference files, which stay bundled). The data/ dir is created if needed.
    """
    data_dir = os.path.join(get_user_data_dir(), 'data')
    os.makedirs(data_dir, exist_ok=True)
    return os.path.join(data_dir, filename)


def get_user_projects_dir():
    """Per-user writable projects directory under <user_data_dir>/projects/.

    Created if it doesn't exist.
    """
    projects_dir = os.path.join(get_user_data_dir(), 'projects')
    os.makedirs(projects_dir, exist_ok=True)
    return projects_dir


def get_bundled_data_path(filename):
    """Full path to shipped reference data under the bundle's data/ folder.

    Frozen: <_MEIPASS>/data/<filename> — so the file reflects the *installed
    version* and updates each time a new build is shipped. Dev: <repo>/data/.
    Use for centrally-maintained reference data (pricing, part catalogs) that
    should track the installed version, NOT per-user edits.
    """
    import sys
    if getattr(sys, 'frozen', False):
        base = os.path.join(sys._MEIPASS, 'data')
    else:
        base = os.path.join(get_app_base_path(), 'data')
    return os.path.join(base, filename)


# User-writable data files that live in the per-user data/ dir (templates the
# user edits/adds at runtime). Catalogs/pricing/NEC/factory are NOT here — those
# read from the bundle so shipped updates win.
_USER_DATA_FILES = [
    'module_templates.json',
    'tracker_templates.json',
    'inverters.json',
]


def initialize_user_data():
    """Populate the per-user data/projects locations on first run.

    On a shipped build the user data dir (%APPDATA%) starts empty. This (1)
    migrates any projects and user templates the user already had next to the
    exe so existing work is preserved, then (2) seeds still-missing templates
    from the bundle. Existing destination files are never overwritten. No-op in
    dev mode, where the user data dir is the repo itself.
    """
    import sys
    import shutil

    user_data_dir = get_user_data_dir()
    app_base = get_app_base_path()

    # Dev mode (user data dir == repo): nothing to migrate or seed.
    if os.path.normcase(os.path.abspath(user_data_dir)) == os.path.normcase(os.path.abspath(app_base)):
        return

    user_data_subdir = os.path.join(user_data_dir, 'data')
    os.makedirs(user_data_subdir, exist_ok=True)
    user_projects_dir = get_user_projects_dir()

    def _copy_if_missing(src, dst):
        try:
            if os.path.isfile(src) and not os.path.exists(dst):
                shutil.copy2(src, dst)
        except Exception as e:
            print(f"initialize_user_data: could not copy {src} -> {dst}: {e}")

    # (1) Migrate existing work from the folder next to the exe (the prior cwd).
    legacy_data = os.path.join(app_base, 'data')
    for name in _USER_DATA_FILES:
        _copy_if_missing(os.path.join(legacy_data, name), os.path.join(user_data_subdir, name))

    legacy_projects = os.path.join(app_base, 'projects')
    if os.path.isdir(legacy_projects):
        for entry in os.listdir(legacy_projects):
            if entry.endswith('.json'):  # project files; skip .recent_projects (rebuilds itself)
                _copy_if_missing(os.path.join(legacy_projects, entry),
                                 os.path.join(user_projects_dir, entry))

    # (2) Seed any still-missing templates from the bundle.
    bundle_data = os.path.join(sys._MEIPASS, 'data') if getattr(sys, 'frozen', False) else legacy_data
    for name in _USER_DATA_FILES:
        _copy_if_missing(os.path.join(bundle_data, name), os.path.join(user_data_subdir, name))