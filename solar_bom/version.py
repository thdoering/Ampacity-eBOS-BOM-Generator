# Version information
VERSION_MAJOR = 2
VERSION_MINOR = 1
VERSION_PATCH = 0

# Get version string
def get_version():
    return f"{VERSION_MAJOR}.{VERSION_MINOR}.{VERSION_PATCH}"

# Get full version info with optional info
def get_version_info(include_build_date=True):
    import datetime
    version = get_version()
    
    if include_build_date:
        build_date = datetime.datetime.now().strftime("%Y-%m-%d")
        return f"v{version} (Build: {build_date})"
    
    return f"v{version}"