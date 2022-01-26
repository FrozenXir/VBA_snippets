import pandas as pd
import ctypes
import os

# output folder here:
folder_path = r"C:\Users\aapostu\Desktop\for Share\Tools for Schaffler\Contracts PDF and JPG split\docs\Output"
batches = 0

for subdir, dirs, files in os.walk(folder_path):
    for file in files:
        file_path = subdir + os.sep + file
        if "files" in subdir:
            break
        elif file_path.endswith(".xlsx"):
            csv_path = subdir + os.sep + file.replace("xlsx", "csv")
            excel_file = pd.read_excel(file_path, dtype=str, index_col=None)
            excel_file.to_csv(csv_path, encoding="utf-8", index=False)
            os.remove(subdir + os.sep + file)
            batches += 1

ctypes.windll.user32.MessageBoxW(
    0, f"Done!\n\nProcessed {batches} batches.", "Operation completed", 0
)
