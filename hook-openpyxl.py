# hooks/hook-openpyxl.py
from PyInstaller.utils.hooks import collect_submodules

hiddenimports = collect_submodules('openpyxl')
hiddenimports.append('openpyxl.cell._writer')  # 确保包含