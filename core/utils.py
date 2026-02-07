"""
Core utility functions for the Invoice Inspector application.
"""

import os
import subprocess
import platform
from pathlib import Path


def open_file(file_path: str | Path) -> bool:
    """
    Opens a file using the system's default application.
    
    Args:
        file_path: Path to the file to open (str or Path object)
        
    Returns:
        True if successful, False otherwise
    """
    try:
        file_path = Path(file_path)
        
        if not file_path.exists():
            print(f"File not found: {file_path}")
            return False
        
        system = platform.system()
        
        if system == 'Windows':
            os.startfile(str(file_path))
        elif system == 'Darwin':  # macOS
            subprocess.run(['open', str(file_path)], check=True)
        else:  # Linux
            subprocess.run(['xdg-open', str(file_path)], check=True)
        
        return True
        
    except Exception as e:
        print(f"Error opening file: {e}")
        return False


def open_file_location(file_path: str | Path) -> bool:
    """
    Opens the folder containing the file and selects it in Explorer.
    
    Args:
        file_path: Path to the file
        
    Returns:
        True if successful, False otherwise
    """
    try:
        file_path = Path(file_path)
        
        if not file_path.exists():
            print(f"File not found: {file_path}")
            return False
        
        system = platform.system()
        
        if system == 'Windows':
            subprocess.run(['explorer', '/select,', str(file_path)], check=True)
        elif system == 'Darwin':  # macOS
            subprocess.run(['open', '-R', str(file_path)], check=True)
        else:  # Linux
            subprocess.run(['xdg-open', str(file_path.parent)], check=True)
        
        return True
        
    except Exception as e:
        print(f"Error opening file location: {e}")
        return False


def import_file(source_path: str | Path, target_dir: str | Path, overwrite: bool = False) -> Path | None:
    """
    Imports (copies) a file to the target directory.
    
    Args:
        source_path: Path to the source file
        target_dir: Target directory to copy the file to
        overwrite: If True, overwrite existing files. If False, skip duplicates.
        
    Returns:
        Path to the copied file, or None if failed
    """
    import shutil
    
    try:
        source_path = Path(source_path)
        target_dir = Path(target_dir)
        
        if not source_path.exists():
            print(f"Source file not found: {source_path}")
            return None
        
        # Create target directory if it doesn't exist
        target_dir.mkdir(parents=True, exist_ok=True)
        
        target_path = target_dir / source_path.name
        
        # Handle duplicates
        if target_path.exists() and not overwrite:
            print(f"File already exists (skipping): {target_path.name}")
            return target_path
        
        shutil.copy2(source_path, target_path)
        print(f"Imported: {source_path.name}")
        return target_path
        
    except Exception as e:
        print(f"Error importing file: {e}")
        return None


def import_files(source_paths: list, target_dir: str | Path, overwrite: bool = False) -> list:
    """
    Imports (copies) multiple files to the target directory.
    
    Args:
        source_paths: List of paths to source files
        target_dir: Target directory to copy files to
        overwrite: If True, overwrite existing files. If False, skip duplicates.
        
    Returns:
        List of successfully imported file paths
    """
    imported = []
    for path in source_paths:
        result = import_file(path, target_dir, overwrite)
        if result:
            imported.append(result)
    return imported


def delete_file(file_path: str | Path, to_trash: bool = True) -> bool:
    """
    Deletes a file.
    
    Args:
        file_path: Path to the file to delete
        to_trash: If True, move to recycle bin (Windows) instead of permanent delete
        
    Returns:
        True if successful, False otherwise
    """
    try:
        file_path = Path(file_path)
        
        if not file_path.exists():
            print(f"File not found: {file_path}")
            return False
        
        if to_trash and platform.system() == 'Windows':
            # Use send2trash if available, otherwise fall back to permanent delete
            try:
                import send2trash
                send2trash.send2trash(str(file_path))
                print(f"Moved to Recycle Bin: {file_path.name}")
                return True
            except ImportError:
                # send2trash not installed, do permanent delete
                pass
        
        # Permanent delete
        file_path.unlink()
        print(f"Deleted: {file_path.name}")
        return True
        
    except Exception as e:
        print(f"Error deleting file: {e}")
        return False


def delete_files(file_paths: list, to_trash: bool = True) -> int:
    """
    Deletes multiple files.
    
    Args:
        file_paths: List of paths to delete
        to_trash: If True, move to recycle bin instead of permanent delete
        
    Returns:
        Number of successfully deleted files
    """
    deleted_count = 0
    for path in file_paths:
        if delete_file(path, to_trash):
            deleted_count += 1
    return deleted_count
