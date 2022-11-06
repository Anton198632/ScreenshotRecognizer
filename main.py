import keyboard as keyboard
import pytesseract
import win32con
import win32gui
import win32ui
from PIL import Image, ImageGrab, ImageTk
import tkinter as tk
import pyperclip

def recognize_text(im):
    pytesseract.pytesseract.tesseract_cmd = r'c:\Program Files\Tesseract-OCR\tesseract'
    recognize_t = pytesseract.image_to_string(im, lang='rus+eng')
    print(recognize_t)
    return recognize_t

def create_bitmap_for_save(x, y, width, height):
    # захватите ручку главного окна рабочего стола
    h_desktop = win32gui.GetDesktopWindow()

    # создание контекста устройства
    desktop_dc = win32gui.GetWindowDC(h_desktop)
    img_dc = win32ui.CreateDCFromHandle(desktop_dc)

    # создание контекста устройства на основе памяти
    mem_dc = img_dc.CreateCompatibleDC()

    # создание растрового объекта
    screenshot = win32ui.CreateBitmap()
    screenshot.CreateCompatibleBitmap(img_dc, width, height)
    mem_dc.SelectObject(screenshot)

    # скопируйте экран в контекст нашего запоминающего устройства
    mem_dc.BitBlt((0, 0), (width, height), img_dc, (x, y), win32con.SRCCOPY)

    bmpinfo = screenshot.GetInfo()
    bmpstr = screenshot.GetBitmapBits(True)
    im = Image.frombuffer('RGBA', (bmpinfo['bmWidth'], bmpinfo['bmHeight']), bmpstr, 'raw', 'RGBA', 0, 1)

    # Сохранение скриншота в файл
    # screenshot.SaveBitmapFile(mem_dc, 'd:\\2.jpg')
    # mem_dc.DeleteDC()
    # win32gui.DeleteObject(screenshot.GetHandle())

    return im


class GUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.withdraw()
        self.attributes('-fullscreen', True)

        self.canvas = tk.Canvas(self)
        self.canvas.pack(fill="both", expand=True)

        image = ImageGrab.grab()
        self.image = ImageTk.PhotoImage(image)
        self.photo = self.canvas.create_image(0, 0, image=self.image, anchor="nw")

        self.x, self.y = 0, 0
        self.rect, self.start_x, self.start_y = None, None, None
        self.deiconify()

        self.canvas.tag_bind(self.photo, "<ButtonPress-1>", self.on_button_press)
        self.canvas.tag_bind(self.photo, "<B1-Motion>", self.on_move_press)
        self.canvas.tag_bind(self.photo, "<ButtonRelease-1>", self.on_button_release)

    # при нажатии на кнопку мыши инициализируем прямоугольник
    def on_button_press(self, event):
        self.start_x = event.x
        self.start_y = event.y
        self.rect = self.canvas.create_rectangle(self.x, self.y, 1, 1, outline='red')

    # при перемещении мыши перерисовываем на холсте прямоугольник
    def on_move_press(self, event):
        curX, curY = (event.x, event.y)
        self.canvas.coords(self.rect, self.start_x, self.start_y, curX, curY)

    # при отпускании кнопки мыши
    def on_button_release(self, event):
        self.bbox = self.canvas.bbox(self.rect) # сохраняем выделенную область в поле bbox
        self.destroy()  # закрываем изображение экрана

    # возвращает координаты выделенной области
    def get_selected_area(self):
        return self.bbox


if __name__ == '__main__':

    print('Для выделения области распознания нажмите "Alt + S"')

    while True:
        if keyboard.is_pressed('alt+s'):
            root = GUI()
            root.mainloop()
            selected_area = root.get_selected_area()

            # создаем растровое изображение по выделенным координатм мыши
            img = create_bitmap_for_save(selected_area[0], selected_area[1], selected_area[2] - selected_area[0],
                                         selected_area[3] - selected_area[1])

            # распознаем текст
            r_text = recognize_text(img)
            r_text = r_text.replace("\n", "\t")

            pyperclip.copy(r_text)

            # записывеаем результат в файл
            with open('result.txt', 'a') as f:
                f.write(f'{r_text}\n')
