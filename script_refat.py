import pytesseract
from PIL import Image, ImageEnhance, ImageFilter
import os
import openpyxl
import pyautogui
from time import sleep

quantidade = 1000
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

def preprocess_image(image_path):
    img = Image.open(image_path)
    img = img.resize((img.width * 2, img.height * 2), Image.Resampling.LANCZOS)  # Aumentar a resolução
    img = img.convert('L')  # Converter para escala de cinza
    img = ImageEnhance.Contrast(img).enhance(2)  # Aumentar o contraste
    img = img.filter(ImageFilter.SHARPEN)  # Aplicar filtro de nitidez
    return img

def print_ip():
    global cont_geral
    left = 909
    top = 835
    width = 108
    height = 38
    screenshot = pyautogui.screenshot(region=(left, top, width, height))
    screenshot_path = os.path.join(screenshot_dir, f"ip_{cont_geral}.png")
    screenshot.save(screenshot_path)

    # Preprocess the image
    processed_img = preprocess_image(screenshot_path)

    # Perform OCR with a whitelist of characters
    custom_config = r'--psm 6 -c tessedit_char_whitelist=0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz'
    ip = pytesseract.image_to_string(processed_img, config=custom_config).strip()
    
    print(ip)
    return ip

def print_data():
    global cont_geral
    left = 950
    top = 655
    width = 240
    height = 23
    screenshot = pyautogui.screenshot(region=(left, top, width, height))
    screenshot_path = os.path.join(screenshot_dir, f"results_{cont_geral}.png")
    screenshot.save(screenshot_path)

    data = pytesseract.image_to_string(preprocess_image(screenshot_path), lang='por', config='--psm 7').strip()
    print(data)
    
    return data

def print_coord():
    global cont_geral
    left = 961
    top = 678
    width = 228
    height = 23
    screenshot_coord = pyautogui.screenshot(region=(left, top, width, height))
    screenshot_path_coord = os.path.join(screenshot_dir, f"coord_{cont_geral}.png")
    screenshot_coord.save(screenshot_path_coord)

    custom_config = r'--oem 3 --psm 7 -c tessedit_char_whitelist=" 0123456789.,NSWE"'
    coord = pytesseract.image_to_string(preprocess_image(screenshot_path_coord), config=custom_config).strip()
    
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
    
    custom_config = r'--oem 3 --psm 6 -c tessedit_char_whitelist=" ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789"'
    logradouro = pytesseract.image_to_string(preprocess_image(screenshot_path), lang='por', config=custom_config).strip()
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

    potencia = pytesseract.image_to_string(preprocess_image(screenshot_path), config='--psm 7').strip()
    print("Potência: "+ potencia)
    
    return potencia

def print_potencia2():
    global cont_geral
    left = 730
    top = 799
    width = 322
    height = 19
    screenshot = pyautogui.screenshot(region=(left, top, width, height))
    screenshot_path = os.path.join(screenshot_dir, f"potencia_{cont_geral}.png")
    screenshot.save(screenshot_path)

    potencia = pytesseract.image_to_string(preprocess_image(screenshot_path), config='--psm 7').strip()
    print("Potência2: "+ potencia)
    
    return potencia

def write_to_excel(ip,potencia, potencia2):
    if not os.path.exists(excel_file_path):
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.title = "Resultado"
        sheet.append(["IP", "POTÊNCIA", "POTÊNCIA2"])  # Adding headers
        workbook.save(excel_file_path)
    
    workbook = openpyxl.load_workbook(excel_file_path)
    sheet = workbook.active
    sheet.append([ip, potencia, potencia2])
    workbook.save(excel_file_path)

while i < quantidade or ip == "IP 02854" or ip == "IP02854" or ip == "ip02854" or ip == "1P02854":
    if i == 0:
        ip = print_ip()
        potencia = print_potencia()
        potencia2 = print_potencia2()

        write_to_excel(ip, potencia, potencia2)
        sleep(1.5)
    else:
        arrow = pyautogui.locateCenterOnScreen('arrow.png', confidence=0.98)
        pyautogui.click(arrow.x, arrow.y)
        pyautogui.moveTo(255, 255)  # Move the mouse away from the arrow to avoid re-detection
        sleep(1.5)
        ip = print_ip()
        potencia = print_potencia()
        potencia2 = print_potencia2()

        write_to_excel(ip, potencia, potencia2)
    i += 1
    cont_geral += 1
