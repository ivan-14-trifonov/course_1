import subprocess
import json
import sys
import os
from datetime import datetime
from pathlib import Path

# Variable for the folder path - change this to the desired path
FOLDER_PATH = r"E:\2022"


def run_powershell_command(folder_path):
    """
    Runs the PowerShell command to get file/folder information recursively
    and returns the results as a Python list of dictionaries.
    """
    # PowerShell command to get file/folder information
    ps_command = (
        f'Get-ChildItem -Path "{folder_path}" -Recurse -Force | '
        'Select-Object FullName, CreationTime, LastWriteTime, LastAccessTime, Length | '
        'ConvertTo-Json -Depth 10'
    )
    
    try:
        # Run PowerShell command
        result = subprocess.run(
            ['powershell', '-Command', ps_command],
            capture_output=True,
            text=True,
            check=True
        )
        
        # Parse JSON output
        output = result.stdout.strip()
        if output:
            try:
                parsed_data = json.loads(output)
            except json.JSONDecodeError:
                print("Could not parse JSON output from PowerShell command.")
                return []
            
            # Ensure we have a list
            if isinstance(parsed_data, dict):
                parsed_data = [parsed_data]
                
            return parsed_data
        else:
            return []
            
    except subprocess.CalledProcessError as e:
        print(f"Error executing PowerShell command: {e}")
        print(f"Error output: {e.stderr}")
        return []
    except FileNotFoundError:
        print("PowerShell not found. This script is designed to run on Windows where PowerShell is available.")
        print("The PowerShell command will not work on macOS/Linux.")
        return []


def get_file_info_cross_platform(path=FOLDER_PATH):
    """
    Cross-platform alternative to PowerShell command for getting file information.
    Returns similar structure to PowerShell command.
    """
    results = []
    
    if not os.path.exists(path):
        print(f"Path {path} does not exist. This script needs to be run on Windows with that path available.")
        return []
    
    try:
        path_obj = Path(path)
        for item in path_obj.rglob('*'):  # Recursively get all files and directories
            try:
                stat = item.stat()
                # On Windows, we can get creation time, but on Unix systems, st_ctime is change time
                creation_time = datetime.fromtimestamp(stat.st_ctime).isoformat()
                modified_time = datetime.fromtimestamp(stat.st_mtime).isoformat()
                accessed_time = datetime.fromtimestamp(stat.st_atime).isoformat()
                
                item_info = {
                    "FullName": str(item),
                    "CreationTime": creation_time,
                    "LastWriteTime": modified_time,
                    "LastAccessTime": accessed_time,
                    "Length": stat.st_size if item.is_file() else None  # Size is only for files
                }
                results.append(item_info)
            except (OSError, IOError):
                # Skip files that can't be accessed
                continue
                
        return results
    except Exception as e:
        print(f"Error accessing path {path}: {e}")
        return []


def main():
    print(f"Attempting to run PowerShell command to get file/folder information from {FOLDER_PATH}...")
    
    # Check if running on Windows
    if sys.platform == "win32":
        results = run_powershell_command(FOLDER_PATH)
        if results:
            print(f"Successfully retrieved {len(results)} items using PowerShell.")
        else:
            print("Trying cross-platform alternative...")
            results = get_file_info_cross_platform()
    else:
        print(f"This script is designed to run on Windows where the {FOLDER_PATH} path and PowerShell are available.")
        print("Using cross-platform alternative to demonstrate functionality...")
        # Using a mock path for demonstration - change as needed
        import tempfile
        with tempfile.TemporaryDirectory() as tmpdir:
            results = get_file_info_cross_platform(tmpdir)
    
    print(f"Retrieved {len(results)} items")
    
    # Print all results
    for i, item in enumerate(results):
        print(f"\nItem {i+1}:")
        print(f"  FullName: {item.get('FullName', 'N/A')}")
        print(f"  CreationTime: {item.get('CreationTime', 'N/A')}")
        print(f"  LastWriteTime: {item.get('LastWriteTime', 'N/A')}")
        print(f"  LastAccessTime: {item.get('LastAccessTime', 'N/A')}")
        print(f"  Length: {item.get('Length', 'N/A')}")
    
    # The full results array is stored in 'results'
    # You can now process this array as needed
    print(f"\nThe results array contains {len(results)} file/folder information objects.")
    return results


if __name__ == "__main__":
    results = main()