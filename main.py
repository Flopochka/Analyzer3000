import time
import math
import pygetwindow as gw
import datetime
import datetime
import os
import threading
import io
import sqlite3
import pyaudio
import numpy as np
import tkinter as tk

from pystray import MenuItem as MenuItem, Icon, Menu
from pynput.mouse import Controller, Listener
from screeninfo import get_monitors
from pynput import keyboard
from PIL import Image, ImageGrab, ImageDraw, ImageFont
from plyer import notification

# Получаем список всех экранов и создаём объект для мыши
monitors = get_monitors()
mouse = Controller()

# Подключаемся к базе данных SQLite (или создаём её, если не существует)
conn = sqlite3.connect('data.db')
cursor = conn.cursor()

# Создаём таблицу, если она не существует
cursor.execute('''
CREATE TABLE IF NOT EXISTS Windows (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT UNIQUE NOT NULL
);
''')

cursor.execute('''
CREATE TABLE IF NOT EXISTS Actions (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    timestamp TEXT NOT NULL,
    total_mouse_move_active_time REAL,
    total_mouse_scroll_active_time REAL,
    total_keyboard_active_time REAL,
    total_mouse_distance REAL,
    total_mouse_error_dist REAL,
    total_scroll_count INTEGER,
    scroll_err_count INTEGER,
    total_key_presses INTEGER,
    error_key_presses INTEGER,
    avg_vol REAL,
    screenshot BLOB
);
''')

cursor.execute('''
CREATE TABLE IF NOT EXISTS WindowData (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    action_id INTEGER,
    window_id INTEGER,
    total_screen_time REAL,
    active_mouse_move_seconds REAL,
    total_mouse_distance REAL,
    mouse_error_dist REAL,
    active_mouse_scroll_seconds REAL,
    total_scroll_count INTEGER,
    scroll_err_count INTEGER,
    active_keyboard_seconds REAL,
    total_key_presses INTEGER,
    error_key_presses INTEGER,
    FOREIGN KEY (action_id) REFERENCES Actions(id),
    FOREIGN KEY (window_id) REFERENCES Windows(id)
);
''')

conn.commit()

# Функция для захвата экрана и сжатия изображения
def capture_screen():
    # Снимаем скриншот активного экрана
    screenshot = ImageGrab.grab()

    # Сжимаем изображение до JPEG с качеством 70
    img_byte_array = io.BytesIO()
    screenshot.save(img_byte_array, format='JPEG', quality=50)
    img_byte_array = img_byte_array.getvalue()

    # Возвращаем сжатое изображение в формате BLOB
    return img_byte_array

# Функция для автоматического форматирования и создания записи
def format_data(data_dict):
    formatted_data = {}
    for key, value in data_dict.items():
        if isinstance(value, dict):  # Если значение - это вложенный словарь, рекурсивно применяем форматирование
            formatted_data[key] = format_data(value)
        else:
            formatted_data[key] = "{:.2f}".format(value)
    return formatted_data

# Функция для отображения уведомлений
def show_notification(title, message):
    notification.notify(
        title=title,
        message=message,
        app_name="Analyzer3000",
        timeout=5
    )
    
# Функция для создания скрытого окна
def create_hidden_window():
    hidden_window = tk.Tk()
    hidden_window.title("Совет")
    hidden_window.geometry("200x100")
    hidden_window.protocol("WM_DELETE_WINDOW", lambda: hidden_window.withdraw())
    hidden_window.withdraw()  # Прячем окно

    # Виджет для текста уведомления
    # notification_text = tk.Label(hidden_window, text="", anchor="w", padx=10, pady=10)
    # notification_text.pack(fill="both", expand=True)
    notification_text = tk.Text(height=10)
    notification_text.pack(anchor="w", fill="both")
    notification_text.insert("1.0", "мама мыла раму")

    def update_text(message):
        notification_text.delete("1.0", tk.END)
        notification_text.insert("1.0", message)
        # notification_text.config(text=message)

    # Функция для сворачивания окна
    def hide_window():
        hidden_window.withdraw()

    # Функция для отображения окна
    def show_window():
        hidden_window.deiconify()

    # Возвращаем скрытое окно и функцию для обновления текста
    return hidden_window, update_text, show_window

# Функция для добавления иконки в системный трей
def on_quit(icon, item):
    icon.stop()
    os._exit(0)  # Завершает процесс мгновенно

# Функция для записи данных в базу
def write_summary_to_db(total_mouse_distance, mouse_error_dist, active_m_secs, windows, active_s_secs, active_k_secs, total_vol, vol_count, total_scroll_count, scroll_err_count, total_key_presses, error_key_presses):
    now = datetime.datetime.now()
    timestamp = now.strftime("%Y-%m-%d %H:%M:%S")
    avg_vol = total_vol/vol_count

    # Захватываем скриншот и получаем сжатое изображение
    screenshot_blob = capture_screen()

    # Форматируем данные окон
    formatted_windows = {name: format_data(win) for name, win in windows.items()}

    # Вставляем данные в таблицу Actions
    cursor.execute('''
    INSERT INTO Actions (timestamp, total_mouse_move_active_time, total_mouse_scroll_active_time, total_keyboard_active_time,
                         total_mouse_distance, total_mouse_error_dist, total_scroll_count, scroll_err_count, total_key_presses, error_key_presses, avg_vol, screenshot)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    ''', (timestamp, active_m_secs, active_s_secs, active_k_secs, total_mouse_distance, mouse_error_dist, total_scroll_count, scroll_err_count, total_key_presses, error_key_presses, avg_vol,
          sqlite3.Binary(screenshot_blob)))
    action_id = cursor.lastrowid  # Получаем ID последней записи из Actions

    # Получаем ID окон
    for window_name in windows.keys():
        # Проверяем, существует ли окно в таблице
        cursor.execute('''
        SELECT id FROM Windows WHERE name = ?
        ''', (window_name,))
        
        window_id = cursor.fetchone()

        # Если окно существует, используем его id, если нет - вставляем новое окно и получаем id
        if window_id is None:
            cursor.execute('''
            INSERT INTO Windows (name) 
            VALUES (?)
            ''', (window_name,))
            # Получаем id только что вставленного окна
            window_id = cursor.lastrowid
        else:
            window_id = window_id[0]

        # Вставляем данные окна в таблицу WindowData
        cursor.execute('''
            INSERT INTO WindowData (action_id, window_id, total_screen_time, active_mouse_move_seconds, total_mouse_distance,
                                mouse_error_dist, active_mouse_scroll_seconds, total_scroll_count, scroll_err_count,
                                active_keyboard_seconds, total_key_presses, error_key_presses)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (action_id,
                window_id,
                window_times[window_name]["total_screen_time"],
                window_times[window_name]["active_mouse_move_seconds"],
                window_times[window_name]["total_mouse_distance"],
                window_times[window_name]["mouse_error_dist"],
                window_times[window_name]["active_mouse_scroll_seconds"],
                window_times[window_name]["total_scroll_count"],
                window_times[window_name]["scroll_err_count"],
                window_times[window_name]["active_keyboard_seconds"],
                window_times[window_name]["total_key_presses"],
                window_times[window_name]["error_key_presses"]
                ))

    # Подтверждаем изменения и сохраняем
    conn.commit()
    return 0, 0, 0, {}, 0, 0, 0, 0, 0, 0, 0, 0

def update_window_time(active_window):
    if active_window not in window_times:
        # Если окна нет, создаем новый объект
        window_times[active_window] = {
            "total_screen_time": 0.1,
            "active_mouse_move_seconds": 0,
            "total_mouse_distance": 0,
            "mouse_error_dist": 0,
            "active_mouse_scroll_seconds": 0,
            "total_scroll_count": 0,
            "scroll_err_count": 0,
            "active_keyboard_seconds": 0,
            "total_key_presses": 0,
            "error_key_presses": 0
        }
    else:
        # Если уже есть, просто увеличиваем счетчик
        window_times[active_window]["total_screen_time"] += 0.1

# Функция для определения активного экрана по координатам мыши
def get_active_screen(x, y):
    for monitor in monitors:
        # Если монитор имеет координаты (monitor.x, monitor.y)
        if monitor.x <= x < monitor.x + monitor.width and monitor.y <= y < monitor.y + monitor.height:
            return monitor
    return None

# Функция для определения активного окна
def get_active_window():
    try:
        win = gw.getActiveWindow()
        if win:
            return win.title
    except Exception as e:
        return f"Error: {e}"
    return "Нет активного окна"

# Функция для вычисления коэффициентов преобразования пикселей в мм для данного монитора
def get_conversion_factors(monitor):
    # Коэффициент по оси X: мм/px
    conversion_x = monitor.width_mm / monitor.width
    # Коэффициент по оси Y: мм/px
    conversion_y = monitor.height_mm / monitor.height
    return conversion_x, conversion_y

# Инициализация переменных для расчёта
last_pos = None          # Последняя позиция мыши (в пикселях)
prev_vector = (0, 0)     # Вектор предыдущего движения
prev_screen = None       # Последний активный монитор
error_save_point = None  # Точка, сохранённая при обнаружении обратного движения
error_save_time = None   # Время, когда была сохранена точка ошибки
mouse_error_distance = 0.0     # Суммарное расстояние ошибок (в мм)
total_mouse_distance_mm = 0.0  # Суммарное пройденное расстояние (в мм)
window_times = {}

# Переменные для активных секунд
active_mouse_move_flag = False      # Флаг, показывающий, что в текущую секунду было движение
active_mouse_scroll_flag = False
active_keyboard_flag = False
active_mouse_move_seconds = 0       # Счётчик активных секунд
active_mouse_scroll_seconds = 0 
active_keyboard_seconds = 0 
last_active_check = time.time()  # Время последней проверки активной секунды
last_write_check = time.time()

# Переменные
total_scroll_count = 0  # Общее количество скроллов
prev_scroll_dir = None  # Направление предыдущего скролла (1 - вверх, -1 - вниз)
scroll_err_dir = None  # Направление ошибки скролла
scroll_err_count = 0  # Количество ошибочных скроллов
scroll_err_temp = 0  # Временный счётчик ошибочных скроллов
last_scroll_time = None  # Время последнего скролла
scroll_timeout = 0.5  # Тайм-аут, через который считаем, что скролл остановился (в секундах)

# Параметры для записи звука
FORMAT = pyaudio.paInt16  # Формат записи
CHANNELS = 1              # Одноканальный звук (моно)
RATE = 44100              # Частота дискретизации (44.1 kHz)
CHUNK = 1024              # Размер фрагмента (как часто будем считывать)

# Переменные для накопления громкости
total_vol = 0
vol_count = 0

# Инициализация pyaudio
p = pyaudio.PyAudio()

# Функция для добавления громкости
def update_volume():
    global total_vol, vol_count

    # Запись данных с микрофона
    stream = p.open(format=FORMAT,
                    channels=CHANNELS,
                    rate=RATE,
                    input=True,
                    frames_per_buffer=CHUNK)
    
        # Чтение данных и вычисление RMS
    try:
        data = np.frombuffer(stream.read(CHUNK), dtype=np.int16)
        
        if data.size == 0:
            raise ValueError("Received empty data chunk")
        
        rms = np.sqrt(np.mean(np.square(data)))  # RMS (среднеквадратическое отклонение)
        
        # Обновление накопительных переменных
        total_vol += rms
        vol_count += 1

    except ValueError as e:
        print(f"Error reading microphone data: {e}")
    except Exception as e:
        print(f"Unexpected error: {e}")

    # Закрытие потока после записи
    stream.stop_stream()
    stream.close()

# Функция для обработки прокрутки
def on_scroll(x, y, dx, dy):
    global total_scroll_count, prev_scroll_dir, scroll_err_dir, scroll_err_count, last_scroll_time, scroll_err_temp, active_mouse_scroll_flag

    active_mouse_scroll_flag = True
    current_time = time.time()  # Текущее время

    # Определяем направление скролла
    if dy > 0:
        current_scroll_dir = 1  # Скроллинг вверх
    elif dy < 0:
        current_scroll_dir = -1  # Скроллинг вниз
    else:
        return  # Нет прокрутки

    # 1. Обновляем количество прокруток
    total_scroll_count += abs(dy)
    active_window = get_active_window()
    update_window_time(active_window)
    window_times[active_window]["total_scroll_count"] += abs(dy)

    # 2. Проверяем ошибочные скроллы (смена направления и остановка)
    if prev_scroll_dir is not None:
        # 1. Если сменилось направление скролла и прошло меньше 0.2 секунд
        if current_scroll_dir != prev_scroll_dir and (current_time - last_scroll_time) < scroll_timeout:
            # 2. Ошибка при смене направления
            if scroll_err_dir is None or scroll_err_dir == prev_scroll_dir:
                scroll_err_dir = prev_scroll_dir  # Запоминаем направление ошибки
                scroll_err_temp += 1  # Увеличиваем временный счётчик ошибок

        # 3. Проверяем, если количество ошибочных скроллов превышает 4
        if scroll_err_temp > 4:
            scroll_err_temp = 0  # Сбрасываем временный счётчик ошибок

    # 4. Проверяем на остановку (если прошло больше 0.2 секунд с последнего скролла)
    if last_scroll_time is not None and current_time - last_scroll_time > scroll_timeout:
        if scroll_err_temp > 0:
            scroll_err_count += scroll_err_temp  # Добавляем ошибки в общий счётчик
            window_times[active_window]["scroll_err_count"] += scroll_err_temp
        scroll_err_temp = 0  # Сбрасываем временные ошибки

    # Обновляем время последнего скролла
    last_scroll_time = current_time
    prev_scroll_dir = current_scroll_dir
    
# Переменные для клавиатуры
total_key_presses = 0  # Общее количество нажатий
error_key_presses = 0  # Количество ошибочных нажатий
prev_key = None  # Предыдущая нажатая клавиша
key_err_temp = 0  # Временный счётчик ошибок
key_err_count = 0  # Общий счётчик ошибок
last_key_time = None  # Время последнего нажатия
key_timeout = 0.2  # Время, после которого ошибка фиксируется
key_state = {}  # Словарь для хранения состояния клавиш (нажата ли клавиша)
window_times = {}  # Данные для каждого окна

# Глобальная переменная для состояния скрытого окна
window_is_open = False

# Функция обработки нажатия клавиш
def on_press(key):
    global total_key_presses, prev_key, key_err_temp, key_err_count, last_key_time, error_key_presses, window_times, key_state
    
    current_time = time.time()  # Текущее время
    key_name = str(key)  # Имя нажатой клавиши
    
    active_window = get_active_window()
    update_window_time(active_window)

    window_times[active_window]["total_key_presses"] += 1  # Обновляем для активного окна
    total_key_presses += 1  # Увеличиваем счётчик нажатий

    # Обновляем состояние клавиши
    key_state[key_name] = True  # Устанавливаем в True, что клавиша нажата

    # Проверяем на ошибки:
    # 1. Ошибочные нажатия: Backspace, Delete или комбинация клавиш (например, Alt + Z)
    if key_name in ['Key.backspace', 'Key.delete', '\x1a']:  # \x16 — это символ для Ctrl + V
        # Для Backspace и Delete считаем ошибку
        window_times[active_window]["error_key_presses"] += 1
        error_key_presses += 1

    # 2. Ошибочные переключения (например, Caps Lock или Alt + Shift)
    if key_name in ['Key.caps_lock', 'Key.shift', 'Key.alt_l', 'Key.alt_r', "Key.insert"]:
        if prev_key == key_name:  # Если клавиша уже была нажата
            key_err_temp += 1
        else:
            prev_key = key_name  # Запоминаем клавишу

    # 3. Проверяем, если прошло больше 0.2 сек с последнего нажатия
    if last_key_time is not None and current_time - last_key_time > key_timeout:
        if key_err_temp > 0:
            key_err_count += key_err_temp
        key_err_temp = 0  # Сбрасываем временный счётчик

    last_key_time = current_time

# Функция обработки отпускания клавиш
def on_release(key):
    global key_state

    key_name = str(key)
    # Устанавливаем состояние клавиши как False, когда она отпускается
    key_state[key_name] = False

    # Остановить слушатель при нажатии ESC
    if key == keyboard.Key.esc:
        return False
    
# Функция для добавления уведомления в трей
def create_tray_icon(hidden_window, update_text, show_window):
    icon = Icon("Analyzer3000", create_icon_image(), menu=Menu(MenuItem("Открыть", lambda: show_window()),
                                                     MenuItem("Выйти", on_quit)))

    icon.run()

# Функция для создания изображения иконки с буквой S
def create_icon_image():
    size = (64, 64)
    image = Image.new('RGBA', size, (0, 0, 0, 0))
    draw = ImageDraw.Draw(image)

    # Рисуем зелёный круг
    draw.ellipse((5, 5, 59, 59), fill=(0, 200, 0))

    # Добавляем букву "S"
    try:
        font = ImageFont.truetype("arial.ttf", 36)  # Стандартный шрифт Windows
    except IOError:
        font = ImageFont.load_default()

    text = "A"
    bbox = draw.textbbox((0, 0), text, font=font)  # Получаем координаты текста
    text_width = bbox[2] - bbox[0]
    text_height = bbox[3] - bbox[1]
    text_position = ((size[0] - text_width) // 2, (size[1] - text_height) // 2)

    draw.text(text_position, text, font=font, fill=(0, 0, 0))

    return image

# Функция для запуска слушателя клавиатуры в отдельном потоке
def start_listener_keyboard():
    with keyboard.Listener(on_press=on_press, on_release=on_release) as listener:
        listener.join()

# Функция для запуска Listener в отдельном потоке
def start_listener_mouse():
    with Listener(on_scroll=on_scroll) as listener:
        listener.join()

def main():
    hidden_window, update_text, show_window = create_hidden_window()
    # Запуск системного трей и скрытого окна
    tray_thread = threading.Thread(target=create_tray_icon, args=(hidden_window, update_text, show_window))
    tray_thread.start()
    update_text("Тестовый текст")
    hidden_window.mainloop()
    # Создаем и запускаем отдельный поток для прослушивания событий
    listener_thread_mouse = threading.Thread(target=start_listener_mouse)
    listener_thread_mouse.daemon = True  # Поток завершится при выходе из программы
    listener_thread_mouse.start()

    # Запуск слушателя в отдельном потоке
    listener_thread_keyboard = threading.Thread(target=start_listener_keyboard)
    listener_thread_keyboard.daemon = True  # Поток будет завершён, когда основная программа завершится
    listener_thread_keyboard.start()

    print("Для выхода нажмите (Ctr+С) в консоли")

    # Основной цикл программы, который будет работать параллельно
    try:
        global last_pos, prev_screen, total_mouse_distance_mm, mouse_error_distance, active_mouse_move_seconds, active_mouse_scroll_seconds
        global active_keyboard_seconds, active_mouse_move_flag, active_mouse_scroll_flag, active_keyboard_flag, prev_vector, error_save_point
        global error_save_time, last_active_check, last_write_check, window_times, total_vol, vol_count, total_scroll_count, scroll_err_count, total_key_presses, error_key_presses
        while True:
            current_time = time.time()
            x, y = mouse.position
            current_screen = get_active_screen(x, y)
            active_window = get_active_window()
            update_window_time(active_window)
            update_volume()
            
            if current_screen is None:
                print("Курсор вне известных экранов!")
                time.sleep(0.1)
                continue

            # Вычисляем коэффициенты преобразования для текущего экрана
            conv_x, conv_y = get_conversion_factors(current_screen)
            
            # Если это первая итерация, инициализируем переменные
            if last_pos is None:
                last_pos = (x, y)
                prev_screen = current_screen
                time.sleep(0.1)
                continue

            # Если произошёл переход между мониторами, используем среднее значение коэффициентов
            if current_screen != prev_screen:
                prev_conv_x, prev_conv_y = get_conversion_factors(prev_screen)
                conv_x = (prev_conv_x + conv_x) / 2
                conv_y = (prev_conv_y + conv_y) / 2
            # Иначе используем коэффициенты текущего монитора

            # Рассчитываем смещение (dx, dy) в пикселях
            dx = x - last_pos[0]
            dy = y - last_pos[1]
            pixel_distance = math.sqrt(dx**2 + dy**2)
            # Рассчитываем пройденное расстояние в мм, используя отдельные коэффициенты по X и Y
            movement_mm = math.sqrt((dx * conv_x)**2 + (dy * conv_y)**2)
            
            MIN_MOVE_MM = 1  # Порог в мм для учёта движения
            if movement_mm >= MIN_MOVE_MM:
                window_times[active_window]["total_mouse_distance"] += movement_mm
                total_mouse_distance_mm += movement_mm
                active_mouse_move_flag = True  # Фиксируем, что в текущую секунду было движение

            # Определяем текущий вектор движения
            current_vector = (dx, dy)
            # Если имеется предыдущий вектор и текущее движение примерно в обратном направлении, сохраняем точку
            if prev_vector != (0, 0):
                dot = prev_vector[0]*current_vector[0] + prev_vector[1]*current_vector[1]
                if dot < 0 and error_save_point is None:
                    error_save_point = last_pos
                    error_save_time = current_time
            
            # Если точка ошибки сохранена и мышь замирает (смещение менее 1 px) более 0.1 сек, учитываем ошибку
            if error_save_point is not None:
                if pixel_distance < 1 and (current_time - error_save_time) >= 0.1:
                    freeze_distance_mm = math.sqrt(((x - error_save_point[0])*conv_x)**2 + ((y - error_save_point[1])*conv_y)**2)
                    # Для ошибки можно взять среднее значение коэффициентов
                    if freeze_distance_mm < 20:
                        mouse_error_distance += 2 * freeze_distance_mm
                        window_times[active_window]["mouse_error_dist"] += 2 * freeze_distance_mm
                    error_save_point = None
                    error_save_time = None

            # Каждую секунду проверяем активность
            if current_time - last_active_check >= 0.5:
                if active_mouse_move_flag:
                    active_mouse_move_seconds += 0.5
                    window_times[active_window]["active_mouse_move_seconds"] += 0.5
                    active_mouse_move_flag = False  # Сбрасываем флаг после учёта
                if active_mouse_scroll_flag:
                    active_mouse_scroll_seconds += 0.5
                    window_times[active_window]["active_mouse_scroll_seconds"] += 0.5
                    active_mouse_scroll_flag = False  # Сбрасываем флаг после учёта
                if active_keyboard_flag:
                    active_keyboard_seconds += 0.5
                    window_times[active_window]["active_keyboard_seconds"] += 0.5
                    active_keyboard_flag = False  # Сбрасываем флаг после учёта
                last_active_check = current_time
            
            if current_time - last_write_check >= 10:
                total_mouse_distance_mm, mouse_error_distance, active_mouse_move_seconds, window_times, active_mouse_scroll_seconds, active_keyboard_seconds, total_vol, vol_count, total_scroll_count, scroll_err_count, total_key_presses, error_key_presses = write_summary_to_db(
                    total_mouse_distance_mm, mouse_error_distance, active_mouse_move_seconds, window_times, active_mouse_scroll_seconds, active_keyboard_seconds, total_vol, vol_count, total_scroll_count, scroll_err_count, total_key_presses, error_key_presses
                )
                last_update_time = current_time
                last_write_check = time.time()
            
            # Обновляем переменные для следующей итерации
            last_pos = (x, y)
            prev_vector = current_vector
            prev_screen = current_screen

            time.sleep(0.1)
            pass
    except KeyboardInterrupt:
        print("Программа завершена пользователем (Ctrl+C).")
    
if __name__ == "__main__":
    main()