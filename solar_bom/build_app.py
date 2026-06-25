import os
import sys
from pathlib import Path
from version import get_version
from src.utils.file_handlers import get_app_base_path, initialize_user_data

# Add necessary directories to path
def setup_environment():
    # Get the base directory (either script location or executable location)
    if getattr(sys, 'frozen', False):
        # Running as compiled exe
        base_dir = Path(sys._MEIPASS)
    else:
        # Running as script
        base_dir = Path(__file__).parent

    # Anchor the working directory to the app's own folder. Without this, a
    # taskbar-pinned (or otherwise launched) exe can run with cwd = C:\Windows\
    # System32, where the app's relative data/ and projects/ access is denied.
    app_dir = get_app_base_path()
    os.chdir(app_dir)

    # Ensure the per-user data/projects locations exist, migrating any existing
    # work next to the exe and seeding templates from the bundle on first run.
    initialize_user_data()

    return base_dir

# Print version info at startup
print(f"Solar eBOS BOM Generator {get_version()}")

# Main function
if __name__ == "__main__":
    base_dir = setup_environment()
    
    # Import the main application after setup
    sys.path.append(str(base_dir))
    from main import main
    main()