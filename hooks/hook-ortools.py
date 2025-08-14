# hook-ortools.py
from PyInstaller.utils.hooks import collect_dynamic_libs

# تقوم هذه الدالة بجمع كل ملفات DLL و PYD التي تحتاجها مكتبة ortools
binaries = collect_dynamic_libs('ortools')