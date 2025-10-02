$excluded = @(
  # Large scientific computing libraries (biggest impact on file size)
  "numpy", "scipy", "matplotlib", "pandas", "sklearn", "statsmodels", "seaborn",
  
  # Large GUI frameworks (if not using Qt)
  "PyQt5", "PyQt6", "PySide2", "PySide6", "PIL.ImageQt", "PIL.ImageQt4", "PIL.ImageQt5",
  
  # Large web frameworks (if not using)
  "flask", "django", "fastapi",
  
  # Large database libraries (if not using)
  "sqlalchemy", "pymongo", "psycopg2", "mysql.connector",
  
  # Large development tools
  "jupyter", "ipython", "notebook", "ipykernel", "jupyter_client"
  
  # Large build/package tools
  "setuptools", "pip", "wheel", "pkg_resources", "setuptools_scm",
  
  # Large cryptography libraries
  "cryptography",
  
  # Large testing frameworks
  "pytest", "nose",
  
  # Large compression (keep zipfile)
  "bz2", "lzma", "tarfile"
)
$import = @(

)

$excludes = ($excluded | ForEach-Object { "--exclude-module $_" }) -join " "
$hiddenImports = ($import | ForEach-Object { "--hidden-import $_" }) -join " "

$cmd = "python -O -m PyInstaller -w -F --additional-hooks-dir=hooks --distpath dist/python --collect-all platformdirs --copy-metadata=platformdirs $hiddenImports $excludes .\main.py"
Invoke-Expression $cmd

# python -O -m PyInstaller -w -F --additional-hooks-dir=hooks --distpath dist/python .\main.py
# pyi-archive_viewer dist/python/main.exe