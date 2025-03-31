import sys
import subprocess
import platform

# Get Python version for wheel naming
py_version = f"cp{sys.version_info.major}{sys.version_info.minor}"

# Define GitHub release base URL
base_url = "https://github.com/cgohlke/pymol-open-source-wheels/releases/latest/download"

# Define expected file name pattern
wheel_filename = f"pymol-3.1.0-{py_version}-{py_version}-win_amd64.whl"

# Construct full download URL
wheel_url = f"{base_url}/{wheel_filename}"

# Check if running on Windows (Gohlke's wheels are Windows-only)
if platform.system() != "Windows":
    print("Error: These PyMOL wheels are only available for Windows.")
    sys.exit(1)

# Install the wheel directly using pip
try:
    subprocess.run(["pip", "install", wheel_url], check=True)
    print(f"Successfully installed {wheel_filename}")
except subprocess.CalledProcessError:
    print(f"Error: Failed to install {wheel_filename}. Ensure the URL exists and your Python version is supported.")
