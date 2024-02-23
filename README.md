# Automatic Printable NameTag Generator

## Introduction

This project aims to automatically generate nametags to load data from an Excel file and inserting it into a PowerPoint which is prinble format. It reads nametag data from a given Excel file and inserts information for each data into PowerPoint slides based on a provided sample template which designed in PowerPoint.

![Automatic Printable NameTag Generator Introduction](img/introduction.png)

## Requirements

- Python 3.x
- `openpyxl` library
- `python-pptx` library

You can install the required libraries using the following command:

```bash
pip install openpyxl python-pptx
```

## Usage

1. Prepare Excel file: Prepare an Excel file ([attendess_list.xlsx](template/attendess_list.xlsx), for example) containing the information to be included in the nametags. Each row should contain information for one nametag.
2. Prepare PowerPoint sample: Prepare a sample PowerPoint file ([nametag.pptx](template/nametage.pptx), for example) to be used when generating the nametags. For Each sample slide put different layouts and designs for the name badges.
3. Run the script: Execute `main.py` Python script to automatically generate the nametags. You can find the generated nametag PowerPoint `dist/generated-*.pptx`.
4. Check the result: The executed script will generate a new PowerPoint file. Open this file to review the generated nametags.

### Notes

- One silde must contain one sample nametag design.
- Text of sample nametag should be one of name of header. Otherwise leave it as is. The script subtitues header text to information text.

## Running the Script

```bash
python main.py -excel 'example/attendees_list-example.xlsx' -pptx 'example/nametag-example.pptx'
```

## Excel File Format

The Excel file containing the necessary information for the nametags should follow a format similar to the following:

| Sample Num | Campus    | Name          | Position |
| ---------- | --------- | ------------- | -------- |
| 0          | Ajou Univ | Kyunghyun Min | SoonJang |
| 1          | ABC Univ  | Jane Smith    | SoonWon  |

The 'Sample Num' column in the Excel file allows you to select from the provided sample templates. sample slide number is start from 0.
For basic tamplate, [attendess_list.xlsx](template/attendess_list.xlsx) file.

## File Structure

```
project-root/
│
├── main.py
├── example.py
├── src/
|   ├── gui.py
│   ├── analyze_slide.py
│   ├── draw.py
|   ├── morefont_pptx.py
|   ├── allow_eastaisa_typeface_pptx.py
│   ├── settable_pptx.py
│   └── utils.py
│
└── template/
    ├── attendees_list.xlsx
    └── nametag.pptx
```

### Description

- `main.py`: Main script
- `example.py`: Run example based on files in example folder
- `src/`: Directory containing source code files
  - `analyze_slide.py`: analyze shapes in slide. position, image, text, font etc. information are extracted
  - `draw.py`: Draw nametag in ppt
  - `morefont_pptx.py`: Extends the `python-pptx` library to add support for user local font directories
  - `allow_eastaisa_typeface_pptx.py`: Extends the `python-pptx` library to make run.font settable
  - `settable_pptx.py`: Extends the `python-pptx` library to add support for East Asian fonts
- `template/`: Directory containing template files
  - `attendees_list.xlsx`: Excel file containing the list of attendees information for nametag
  - `nametag.pptx`: PowerPoint template for sample nametags. Custom slide layout are applied

## License

This project is licensed under the MIT License. For more information, see the [LICENSE](LICENSE) file.
