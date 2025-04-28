import os
import sys
import shutil
from pathlib import Path

# Add necessary directories to path
def setup_environment():
    # Get the base directory (either script location or executable location)
    if getattr(sys, 'frozen', False):
        # Running as compiled exe
        base_dir = Path(sys._MEIPASS)
    else:
        # Running as script
        base_dir = Path(__file__).parent
    
    # Ensure data directories exist
    os.makedirs(os.path.join(base_dir, 'data'), exist_ok=True)
    os.makedirs(os.path.join(base_dir, 'projects'), exist_ok=True)
    
    # Copy template files if they don't exist yet
    template_files = ['module_templates.json', 'tracker_templates.json']
    for file in template_files:
        source = os.path.join(base_dir, 'data', file)
        dest = os.path.join(os.getcwd(), 'data', file) 
        if os.path.exists(source) and not os.path.exists(dest):
            os.makedirs(os.path.dirname(dest), exist_ok=True)
            shutil.copy(source, dest)
    
    return base_dir

# Main function
if __name__ == "__main__":
    base_dir = setup_environment()
    
    # Import the main application after setup
    sys.path.append(str(base_dir))
    from main import main
    main()