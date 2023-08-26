


# Excel Translator with AI Models

This program translates specified Excel sheets and cell ranges in the desired translation direction (English-Turkish or Turkish-English) using the Helsinki-NLP--opus-mt-tr-en and Helsinki-NLP--opus-tatoeba-en-tr open-source NLP models. The program works through the command line and requires no internet access for model loading.

## Usage

1. Place your Excel files in the same directory as this program.
2. Run the program using Python: `python excel_translator.py`.
3. Follow the on-screen instructions to select translation direction, Excel file, sheets, and cell ranges.
4. Translated content will be saved in a new Excel file named 'translated_<original_file_name>.xlsx'.

## Requirements

- Python 3.x
- openpyxl library
- transformers library


---
<b> To install the necessary dependencies using the requirements.txt file, follow these instructions: </b>

  1. Open a terminal or command prompt on your system.
  2. Navigate to the directory where your Python script and requirements.txt file are located. You can use the cd command to change directories.
  3. Run the following command to install the required packages using pip:
     
     ```

     pip install -r requirements.txt

     ```


---

## Important Notes

- Do not move or delete model files. The program will re-download them if needed.
- If the models are already downloaded, run the program in the same directory as the models.
- Translation direction and other options will be prompted during runtime.

--- 

