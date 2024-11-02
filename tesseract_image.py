import os
import pytesseract
from PIL import Image
import openpyxl

pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
folder_path = r'####'

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Tesseract Outputs"
ws.append(["Filename", "Result 1"])

file_list = os.listdir(folder_path)
results_dict = {filename: [] for filename in file_list}

for filename in file_list:
    file_path = os.path.join(folder_path, filename)

    try:
        img = Image.open(file_path)
        text = pytesseract.image_to_string(img, lang='eng')
        results_dict[filename].append(text)
        ws.append([filename, text])

        print(f"Processed: {filename}")
    except Exception as e:
        print(f"Error processing {filename}: {e}")


# 엑셀 파일 저장
output_excel_path = r'####.xlsx'
wb.save(output_excel_path)

print(f"Results saved to {output_excel_path}")