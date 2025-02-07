import os
import importlib

module_dir = os.path.dirname(__file__)

# patches 폴더 내의 패치 파일만 자동으로 import
for filename in os.listdir(module_dir):
    if filename.endswith(".py") and filename != "__init__.py":
        module_name = f"src.patches.{filename[:-3]}"
        importlib.import_module(module_name)
