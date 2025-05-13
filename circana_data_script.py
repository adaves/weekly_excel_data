import openpyxl
import os
import re
import glob
import shutil
from datetime import datetime, timedelta

def extract_date_from_filename(filename: str) -> tuple:
    """
    Extract date from filename in various formats.
    
    Args:
        filename: The Excel file name
        
    Returns:
        Tuple of (month, day, year)
    """
    # Try MM.DD.YY format (e.g., 04.27.25)
    dot_pattern = r'(\d{2})\.(\d{2})\.(\d{2})'
    match = re.search(dot_pattern, filename)
    if match:
        month, day, year = map(int, match.groups())
        # Assume 20xx for year
        return month, day, 2000 + year
    
    # Try MMDDYY format without dots (e.g., 042025)
    no_dot_pattern = r'WE\s+(\d{2})(\d{2})(\d{2})'
    match = re.search(no_dot_pattern, filename)
    if match:
        month, day, year = map(int, match.groups())
        return month, day, 2000 + year
    
    # Fallback: use previous Sunday's date
    today = datetime.now()
    days_since_sunday = today.weekday() + 1  # +1 because weekday() has Monday=0, Sunday=6
    previous_sunday = today - timedelta(days=days_since_sunday)
    return previous_sunday.month, previous_sunday.day, previous_sunday.year

def format_date(date_tuple: tuple) -> str:
    """
    Format date tuple as mm-dd-yyyy string.
    
    Args:
        date_tuple: Tuple of (month, day, year)
        
    Returns:
        Formatted date string
    """
    month, day, year = date_tuple
    return f"{month:02d}-{day:02d}-{year}"

def create_new_filename(original_filename: str) -> str:
    """
    Create new filename with date prefix.
    
    Args:
        original_filename: Original Excel filename
        
    Returns:
        New filename with date prefix
    """
    date_tuple = extract_date_from_filename(original_filename)
    formatted_date = format_date(date_tuple)
    return f"{formatted_date}_{original_filename}"

def get_output_path(filename: str, output_dir: str = "modified_excel_workbooks") -> str:
    """
    Get output path in specified directory.
    
    Args:
        filename: The filename to save
        output_dir: Directory to save to (default: modified_excel_workbooks)
        
    Returns:
        Full output path
    """
    os.makedirs(output_dir, exist_ok=True)
    return os.path.join(output_dir, filename)

def get_archive_path(filename: str, archive_dir: str = "archived_data") -> str:
    """
    Get path for archiving original file.
    
    Args:
        filename: The filename to archive
        archive_dir: Directory to archive to (default: archived_data)
        
    Returns:
        Full archive path
    """
    os.makedirs(archive_dir, exist_ok=True)
    return os.path.join(archive_dir, filename)

def find_excel_files(directory: str = ".") -> list:
    """
    Find all Excel files in the specified directory.
    Excludes files in modified_excel_workbooks and archived_data directories.
    
    Args:
        directory: Directory to search for Excel files (default: current directory)
        
    Returns:
        List of full paths to Excel files
    """
    excel_files = []
    
    # Get all Excel files
    for ext in ["*.xlsx", "*.xls"]:
        pattern = os.path.join(directory, ext)
        excel_files.extend(glob.glob(pattern))
    
    # Filter out files from special directories
    filtered_files = []
    for file_path in excel_files:
        if "modified_excel_workbooks" not in file_path and "archived_data" not in file_path:
            filtered_files.append(file_path)
    
    return filtered_files

def unhide_all_sheets(excel_path: str, output_path: str = None) -> None:
    """
    Unhide all sheets in the given Excel workbook and save to output path.
    
    Args:
        excel_path: Path to the Excel workbook
        output_path: Path to save the modified workbook (optional)
    """
    wb = openpyxl.load_workbook(excel_path)
    for sheet in wb.worksheets:
        sheet.sheet_state = 'visible'
    
    if output_path:
        wb.save(output_path)
    else:
        wb.save(excel_path)

def process_excel_file(excel_path: str, output_dir: str = "modified_excel_workbooks", 
                       archive_dir: str = "archived_data") -> str:
    """
    Process Excel file: unhide sheets, save with date prefix, and archive original.
    
    Args:
        excel_path: Path to the Excel workbook
        output_dir: Directory to save processed files
        archive_dir: Directory to archive original files
        
    Returns:
        Path to the processed file
    """
    # Get filenames
    filename = os.path.basename(excel_path)
    new_filename = create_new_filename(filename)
    output_path = get_output_path(new_filename, output_dir)
    archive_path = get_archive_path(filename, archive_dir)
    
    # Process the file
    unhide_all_sheets(excel_path, output_path)
    
    # Move the original file instead of copying
    shutil.move(excel_path, archive_path)
    
    return output_path

def process_all_excel_files(directory: str = ".", output_dir: str = "modified_excel_workbooks", 
                           archive_dir: str = "archived_data") -> dict:
    """
    Process all Excel files in the specified directory.
    
    Args:
        directory: Directory to search for Excel files
        output_dir: Directory to save processed files
        archive_dir: Directory to archive original files
        
    Returns:
        Dictionary with results for each file: {file_path: {"status": bool, "output": str, "error": str}}
    """
    excel_files = find_excel_files(directory)
    results = {}
    
    for file_path in excel_files:
        results[file_path] = {"status": False, "output": None, "error": None}
        
        try:
            output_path = process_excel_file(file_path, output_dir, archive_dir)
            results[file_path]["status"] = True
            results[file_path]["output"] = output_path
        except Exception as e:
            results[file_path]["error"] = str(e)
    
    return results

if __name__ == "__main__":
    print("Starting Excel processing...")
    results = process_all_excel_files()
    
    # Print results summary
    processed_count = sum(1 for r in results.values() if r["status"])
    failed_count = len(results) - processed_count
    
    print(f"\nProcessing complete!")
    print(f"Files processed successfully: {processed_count}")
    print(f"Files failed: {failed_count}")
    
    # Print details for each file
    if results:
        print("\nDetailed results:")
        for file_path, result in results.items():
            if result["status"]:
                print(f"✓ {os.path.basename(file_path)} -> {os.path.basename(result['output'])}")
            else:
                print(f"✗ {os.path.basename(file_path)}: {result['error']}")