import pytesseract
from PIL import Image
import os
import openpyxl
import pyautogui
from time import sleep

quantidade = 100
cont_geral = 0
i = 0
caminho = r"C:\Program Files\Tesseract-OCR"
pytesseract.pytesseract.tesseract_cmd = caminho + r'\tesseract.exe'

# Path to save the screenshot
screenshot_dir = r"C:/Users/AdminDell/Desktop/AUTOMATIZACAO_RELATORIO_GARCA/images"
if not os.path.exists(screenshot_dir):
    os.makedirs(screenshot_dir)
    
# Path to save the Excel file
excel_file_path = r"C:/Users/AdminDell/Desktop/AUTOMATIZACAO_RELATORIO_GARCA/results.xlsx"

def print_ip():
    global cont_geral
    left = 919
    top = 841
    width = 83
    height = 20
    screenshot = pyautogui.screenshot(region=(left, top, width, height))
    screenshot_path = os.path.join(screenshot_dir, f"ip_{cont_geral}.png")
    screenshot.save(screenshot_path)

    ip = pytesseract.image_to_string(screenshot_path, config='--psm 7').strip()
    print(ip)
    
    return ip

def print_data():
    global cont_geral
    left = 952
    top = 655
    width = 238
    height = 26
    screenshot = pyautogui.screenshot(region=(left, top, width, height))
    screenshot_path = os.path.join(screenshot_dir, f"results_{cont_geral}.png")
    screenshot.save(screenshot_path)

    data = pytesseract.image_to_string(screenshot_path, lang='por', config='--psm 7').strip()
    print(data)
    
    return data

def print_coord():
    global cont_geral
    left = 969
    top = 674
    width = 222
    height = 27
    screenshot_coord = pyautogui.screenshot(region=(left, top, width, height))
    screenshot_path_coord = os.path.join(screenshot_dir, f"coord_{cont_geral}.png")
    screenshot_coord.save(screenshot_path_coord)

    custom_config = r'--oem 3 --psm 7 -c tessedit_char_whitelist=" 0123456789.,NSWE"'
    coord = pytesseract.image_to_string(screenshot_path_coord, config=custom_config).strip()
    
    print("Coordenadas: ")
    print(coord)
    
    return coord


def print_logradouro():
    global cont_geral
    left = 737
    top = 797
    width = 378
    height = 21
    screenshot = pyautogui.screenshot(region=(left, top, width, height))
    screenshot_path = os.path.join(screenshot_dir, f"logradouro_{cont_geral}.png")
    screenshot.save(screenshot_path)

    logradouro = pytesseract.image_to_string(screenshot_path, lang='por', config='--psm 7').strip()
    print("Logradouro: "+ logradouro)
    
    return logradouro

def print_potencia():
    global cont_geral
    left = 1129
    top = 800
    width = 61
    height = 18
    screenshot = pyautogui.screenshot(region=(left, top, width, height))
    screenshot_path = os.path.join(screenshot_dir, f"potencia_{cont_geral}.png")
    screenshot.save(screenshot_path)

    potencia = pytesseract.image_to_string(screenshot_path, config='--psm 7').strip()
    print("Potência: "+ potencia)
    
    return potencia

def write_to_excel(data, coord, ip, logradouro, potencia):
    if not os.path.exists(excel_file_path):
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.title = "Resultado"
        sheet.append(["DATA", "COORD", "IP", "LOGRADOURO", "POTÊNCIA"])  # Adding headers
        workbook.save(excel_file_path)
    
    workbook = openpyxl.load_workbook(excel_file_path)
    sheet = workbook.active
    sheet.append([data, coord, ip, logradouro, potencia])
    workbook.save(excel_file_path)

while i < quantidade:
    if i == 0:
        data = print_data()
        coord = print_coord()
        ip = print_ip()
        logradouro = print_logradouro()
        potencia = print_potencia()

        write_to_excel(data, coord, ip, logradouro, potencia)
        sleep(1.5)
    else:
        arrow = pyautogui.locateCenterOnScreen('arrow.png', confidence=0.98)
        pyautogui.click(arrow.x, arrow.y)
        pyautogui.moveTo(255, 255)  # Move the mouse away from the arrow to avoid re-detection
        sleep(1.5)
        data = print_data()
        coord = print_coord()
        ip = print_ip()
        potencia = print_potencia()
        logradouro = print_logradouro()

        write_to_excel(data, coord, ip, logradouro, potencia)
    i += 1
    cont_geral += 1
