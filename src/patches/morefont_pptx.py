import os
from pptx.text.fonts import FontFiles

def _windows_font_directories():
    return [os.path.join(os.environ['WINDIR'], "Fonts") , os.path.join(os.environ['USERPROFILE'], "AppData", "Local", "Microsoft", "Windows", "Fonts")]

FontFiles._windows_font_directories = _windows_font_directories

if __name__ == "__main__":
    font_names = [f[0] for f in FontFiles._installed_fonts()]
    print(font_names)