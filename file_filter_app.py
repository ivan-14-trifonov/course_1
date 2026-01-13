import subprocess
import json
import sys
import os
from datetime import datetime
from pathlib import Path
from tkinter import *
from tkinter import ttk, filedialog
from tkcalendar import DateEntry
import re


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


def filter_results(results, name_pattern=None, creation_date_from=None, creation_date_to=None,
                   modified_date_from=None, modified_date_to=None, 
                   accessed_date_from=None, accessed_date_to=None):
    """
    Filter the results based on name pattern and date ranges
    """
    filtered_results = []
    
    for item in results:
        # Name filtering
        if name_pattern:
            if not re.search(name_pattern, os.path.basename(item.get('FullName', '')), re.IGNORECASE):
                continue
        
        # Convert date strings to datetime objects for comparison
        creation_time = datetime.fromisoformat(item.get('CreationTime', '').replace('Z', '+00:00'))
        modified_time = datetime.fromisoformat(item.get('LastWriteTime', '').replace('Z', '+00:00'))
        accessed_time = datetime.fromisoformat(item.get('LastAccessTime', '').replace('Z', '+00:00'))
        
        # Creation time filtering
        if creation_date_from and creation_time < datetime.combine(creation_date_from, datetime.min.time()):
            continue
        if creation_date_to and creation_time > datetime.combine(creation_date_to, datetime.max.time()):
            continue
            
        # Modified time filtering
        if modified_date_from and modified_time < datetime.combine(modified_date_from, datetime.min.time()):
            continue
        if modified_date_to and modified_time > datetime.combine(modified_date_to, datetime.max.time()):
            continue
            
        # Accessed time filtering
        if accessed_date_from and accessed_time < datetime.combine(accessed_date_from, datetime.min.time()):
            continue
        if accessed_date_to and accessed_time > datetime.combine(accessed_date_to, datetime.max.time()):
            continue
        
        filtered_results.append(item)
    
    return filtered_results


class FileFilterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("File Explorer with Filtering")
        self.root.geometry("1200x800")
        
        # Variables for filters
        self.name_var = StringVar()
        self.creation_from_var = StringVar()
        self.creation_to_var = StringVar()
        self.modified_from_var = StringVar()
        self.modified_to_var = StringVar()
        self.accessed_from_var = StringVar()
        self.accessed_to_var = StringVar()
        
        # Original data
        self.original_results = []
        self.filtered_results = []
        
        # Create UI
        self.create_widgets()
        
        # Load initial data
        self.load_data()
    
    def create_widgets(self):
        # Main frame
        main_frame = Frame(self.root)
        main_frame.pack(fill=BOTH, expand=True, padx=10, pady=10)
        
        # Control panel
        control_frame = LabelFrame(main_frame, text="Filters", padx=10, pady=10)
        control_frame.pack(fill=X, pady=(0, 10))
        
        # Name filter
        Label(control_frame, text="Name Pattern:").grid(row=0, column=0, sticky=W, padx=(0, 5))
        self.name_entry = Entry(control_frame, textvariable=self.name_var, width=20)
        self.name_entry.grid(row=0, column=1, padx=(0, 10))
        
        # Creation date filters
        Label(control_frame, text="Creation From:").grid(row=0, column=2, sticky=W, padx=(10, 5))
        self.creation_from_entry = DateEntry(control_frame, textvariable=self.creation_from_var, 
                                             date_pattern='yyyy-mm-dd', width=12)
        self.creation_from_entry.grid(row=0, column=3, padx=(0, 5))
        
        Label(control_frame, text="To:").grid(row=0, column=4, sticky=W, padx=(5, 5))
        self.creation_to_entry = DateEntry(control_frame, textvariable=self.creation_to_var, 
                                           date_pattern='yyyy-mm-dd', width=12)
        self.creation_to_entry.grid(row=0, column=5, padx=(0, 10))
        
        # Modification date filters
        Label(control_frame, text="Modified From:").grid(row=1, column=0, sticky=W, padx=(0, 5))
        self.modified_from_entry = DateEntry(control_frame, textvariable=self.modified_from_var, 
                                             date_pattern='yyyy-mm-dd', width=12)
        self.modified_from_entry.grid(row=1, column=1, padx=(0, 5))
        
        Label(control_frame, text="To:").grid(row=1, column=2, sticky=W, padx=(5, 5))
        self.modified_to_entry = DateEntry(control_frame, textvariable=self.modified_to_var, 
                                           date_pattern='yyyy-mm-dd', width=12)
        self.modified_to_entry.grid(row=1, column=3, padx=(0, 5))
        
        # Access date filters
        Label(control_frame, text="Accessed From:").grid(row=2, column=0, sticky=W, padx=(0, 5))
        self.accessed_from_entry = DateEntry(control_frame, textvariable=self.accessed_from_var, 
                                             date_pattern='yyyy-mm-dd', width=12)
        self.accessed_from_entry.grid(row=2, column=1, padx=(0, 5))
        
        Label(control_frame, text="To:").grid(row=2, column=2, sticky=W, padx=(5, 5))
        self.accessed_to_entry = DateEntry(control_frame, textvariable=self.accessed_to_var, 
                                           date_pattern='yyyy-mm-dd', width=12)
        self.accessed_to_entry.grid(row=2, column=3, padx=(0, 5))
        
        # Buttons
        button_frame = Frame(control_frame)
        button_frame.grid(row=0, column=6, rowspan=3, padx=(20, 0))
        
        Button(button_frame, text="Apply Filters", command=self.apply_filters).pack(pady=2)
        Button(button_frame, text="Clear Filters", command=self.clear_filters).pack(pady=2)
        Button(button_frame, text="Browse Folder", command=self.browse_folder).pack(pady=2)
        
        # Results count label
        self.results_label = Label(control_frame, text="Results: 0")
        self.results_label.grid(row=3, column=0, columnspan=7, sticky=W, pady=(10, 0))
        
        # Treeview for displaying results
        tree_frame = Frame(main_frame)
        tree_frame.pack(fill=BOTH, expand=True)
        
        # Create treeview
        columns = ('FullName', 'CreationTime', 'LastWriteTime', 'LastAccessTime', 'Length')
        self.tree = ttk.Treeview(tree_frame, columns=columns, show='headings', height=20)
        
        # Define headings
        self.tree.heading('FullName', text='Full Name')
        self.tree.heading('CreationTime', text='Creation Time')
        self.tree.heading('LastWriteTime', text='Modification Time')
        self.tree.heading('LastAccessTime', text='Access Time')
        self.tree.heading('Length', text='Size (bytes)')
        
        # Define column widths
        self.tree.column('FullName', width=400)
        self.tree.column('CreationTime', width=150)
        self.tree.column('LastWriteTime', width=150)
        self.tree.column('LastAccessTime', width=150)
        self.tree.column('Length', width=100)
        
        # Add scrollbar
        scrollbar = ttk.Scrollbar(tree_frame, orient=VERTICAL, command=self.tree.yview)
        self.tree.configure(yscroll=scrollbar.set)
        scrollbar.pack(side=RIGHT, fill=Y)
        self.tree.pack(side=LEFT, fill=BOTH, expand=True)
    
    def load_data(self):
        print(f"Loading data from {FOLDER_PATH}...")
        
        # Check if running on Windows
        if sys.platform == "win32":
            self.original_results = run_powershell_command(FOLDER_PATH)
            if not self.original_results:
                print("Trying cross-platform alternative...")
                self.original_results = get_file_info_cross_platform()
        else:
            print(f"This script is designed to run on Windows where the {FOLDER_PATH} path and PowerShell are available.")
            print("Using cross-platform alternative to demonstrate functionality...")
            # For demo purposes, we could create sample data here
            self.original_results = get_file_info_cross_platform()
        
        print(f"Loaded {len(self.original_results)} items")
        self.filtered_results = self.original_results[:]
        self.update_treeview()
    
    def browse_folder(self):
        """Allow user to select a folder to scan"""
        folder_selected = filedialog.askdirectory(initialdir=FOLDER_PATH, title="Select folder to scan")
        if folder_selected:
            global FOLDER_PATH
            FOLDER_PATH = folder_selected
            self.load_data()
    
    def apply_filters(self):
        """Apply the selected filters to the data"""
        # Get values from GUI
        name_pattern = self.name_var.get().strip()
        creation_from = self.creation_from_entry.get_date() if self.creation_from_entry.get() else None
        creation_to = self.creation_to_entry.get_date() if self.creation_to_entry.get() else None
        modified_from = self.modified_from_entry.get_date() if self.modified_from_entry.get() else None
        modified_to = self.modified_to_entry.get_date() if self.modified_to_entry.get() else None
        accessed_from = self.accessed_from_entry.get_date() if self.accessed_from_entry.get() else None
        accessed_to = self.accessed_to_entry.get_date() if self.accessed_to_entry.get() else None
        
        # Apply filters
        self.filtered_results = filter_results(
            self.original_results,
            name_pattern=name_pattern or None,
            creation_date_from=creation_from,
            creation_date_to=creation_to,
            modified_date_from=modified_from,
            modified_date_to=modified_to,
            accessed_date_from=accessed_from,
            accessed_date_to=accessed_to
        )
        
        self.update_treeview()
        
        # Update results count
        total_count = len(self.original_results)
        filtered_count = len(self.filtered_results)
        self.results_label.config(text=f"Results: {filtered_count} of {total_count}")
    
    def clear_filters(self):
        """Clear all filters and reset to original data"""
        self.name_var.set("")
        self.creation_from_var.set("")
        self.creation_to_var.set("")
        self.modified_from_var.set("")
        self.modified_to_var.set("")
        self.accessed_from_var.set("")
        self.accessed_to_var.set("")
        
        self.filtered_results = self.original_results[:]
        self.update_treeview()
        
        # Update results count
        self.results_label.config(text=f"Results: {len(self.filtered_results)}")
    
    def update_treeview(self):
        """Update the treeview with filtered results"""
        # Clear existing items
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # Insert filtered results
        for item in self.filtered_results:
            # Format the row for display
            formatted_item = (
                item.get('FullName', ''),
                item.get('CreationTime', '')[:19],  # Show only date and time without timezone
                item.get('LastWriteTime', '')[:19],
                item.get('LastAccessTime', '')[:19],
                str(item.get('Length', '')) if item.get('Length') is not None else ''
            )
            self.tree.insert('', END, values=formatted_item)


def main():
    root = Tk()
    app = FileFilterApp(root)
    root.mainloop()


if __name__ == "__main__":
    # Install required packages if not already installed
    try:
        import tkinter
        import tkcalendar
    except ImportError:
        print("Required packages not found. Please install them:")
        print("pip install tkinter")
        print("pip install tkcalendar")
        sys.exit(1)
    
    main()