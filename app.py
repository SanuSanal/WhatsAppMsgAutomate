import openpyxl as xl
from io import BytesIO
import time
import webbrowser
from keyboard import press_and_release
import win32clipboard
from PIL import Image


def send_to_clipboard(filepath):
    image = Image.open(filepath)

    output = BytesIO()
    image.convert("RGB").save(output, "BMP")
    data = output.getvalue()[14:]
    output.close()
    win32clipboard.OpenClipboard()
    win32clipboard.EmptyClipboard()
    win32clipboard.SetClipboardData(win32clipboard.CF_DIB, data)
    win32clipboard.CloseClipboard()


wb = xl.load_workbook('contacts.xlsx')
sheet = wb['Sheet1']

for row in range(2, sheet.max_row + 1):
    mobile_number = sheet.cell(row, 1).value
    message = sheet.cell(row, 2).value
    image_file_path = 'imgs\\' + sheet.cell(row, 3).value
    print(f'sending message, To: {mobile_number}, message: {message}, image location: {image_file_path}')
    webbrowser.open(f'whatsapp://send?phone={mobile_number}&text={message}')
    send_to_clipboard(image_file_path)
    time.sleep(1)
    press_and_release('ctrl+v')
    time.sleep(1)
    press_and_release('enter')