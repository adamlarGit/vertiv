import os
from pathlib import Path

# --- CONFIGURATION ---
TARGET_DIRECTORY = "EXCEL"
EXCEL_EXTENSIONS = {".xlsx", ".xls", ".xlsm", ".xlsb"}
# ---------------------

def clear_excel_files(directory):
    path = Path(directory)
    if not path.exists():
        print(f"Error: Directory '{directory}' does not exist.")
        return

    print(f"Scanning: {path.absolute()}")
    
    # 1. DRY RUN / COLLECTION
    files_to_delete = []
    for file_path in path.rglob("*"):
        if file_path.is_file() and file_path.suffix.lower() in EXCEL_EXTENSIONS:
            files_to_delete.append(file_path)

    if not files_to_delete:
        print("No Excel files found.")
        return

    print("\n--- FILES IDENTIFIED FOR DELETION ---")
    for f in files_to_delete:
        print(f"  - {f}")
    
    print("-" * 30)
    print(f"Total files found: {len(files_to_delete)}")
    
    # 2. CONFIRMATION
    confirm = input("\nDo you want to delete these files? (yes/no): ").strip().lower()
    
    if confirm in ("yes", "y"):
        print("\n[DELETING] Starting cleanup...")
        deleted_count = 0
        for file_path in files_to_delete:
            try:
                file_path.unlink()
                print(f"Deleted: {file_path}")
                deleted_count += 1
            except Exception as e:
                print(f"Error deleting {file_path}: {e}")
        
        print("-" * 30)
        print(f"Cleanup complete. Successfully deleted {deleted_count} files.")
    else:
        print("\n[CANCELLED] No files were deleted.")

if __name__ == "__main__":
    clear_excel_files(TARGET_DIRECTORY)
