import winreg
import os
import sys
import platform
import ctypes
import time
import urllib.request
import zipfile
import shutil
import json


# --- СИСТЕМНЫЕ УТИЛИТЫ ---
def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False


def clear_screen():
    os.system('cls' if os.name == 'nt' else 'clear')


def restart_explorer_generic():
    print("\n[!] Для применения изменений необходимо перезапустить проводник.")
    choice = input("    Перезапустить сейчас? (y/n): ").lower()
    if choice == 'y':
        print("[*] Перезапуск проводника...", end="")
        time.sleep(1)
        os.system("taskkill /f /im explorer.exe >nul 2>&1")
        os.system("start explorer.exe")
        print(" Готово.")
    else:
        print("[i] Изменения вступят в силу после перезагрузки ПК.")
    input("\nНажмите Enter для продолжения...")


# --- ЛОГИКА: ПРОВОДНИК И ПАНЕЛЬ ЗАДАЧ ---

def enable_classic_menu():
    key_path = r"Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}\InprocServer32"
    try:
        key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, key_path)
        winreg.SetValueEx(key, "", 0, winreg.REG_SZ, "")
        winreg.CloseKey(key)
        print("\n[+] УСПЕХ: Классическое меню (Win 10) включено!")
        restart_explorer_generic()
    except Exception as e:
        print(f"\n[-] ОШИБКА: {e}")
        input("Enter...")


def enable_modern_menu():
    key_parent = r"Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}"
    try:
        winreg.DeleteKey(winreg.HKEY_CURRENT_USER, key_parent + r"\InprocServer32")
        winreg.DeleteKey(winreg.HKEY_CURRENT_USER, key_parent)
        print("\n[+] УСПЕХ: Современное меню (Win 11) восстановлено!")
        restart_explorer_generic()
    except FileNotFoundError:
        print("\n[i] Современное меню уже активно.")
        input("Enter...")
    except Exception as e:
        print(f"\n[-] ОШИБКА: {e}")
        input("Enter...")


def toggle_file_extensions(show):
    key_path = r"Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
    try:
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_WRITE)
        val = 0 if show else 1
        winreg.SetValueEx(key, "HideFileExt", 0, winreg.REG_DWORD, val)
        state = "ВИДИМЫ" if show else "СКРЫТЫ"
        print(f"\n[+] Расширения файлов теперь {state}.")
        winreg.CloseKey(key)
        restart_explorer_generic()
    except Exception as e:
        print(f"\n[-] Ошибка: {e}")
        input("Enter...")


def toggle_hidden_items(show):
    key_path = r"Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
    try:
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_WRITE)
        val = 1 if show else 2
        winreg.SetValueEx(key, "Hidden", 0, winreg.REG_DWORD, val)
        state = "ВИДИМЫ" if show else "СКРЫТЫ"
        print(f"\n[+] Скрытые папки и файлы теперь {state}.")
        winreg.CloseKey(key)
        restart_explorer_generic()
    except Exception as e:
        print(f"\n[-] Ошибка: {e}")
        input("Enter...")


def toggle_taskbar_end_task(enable):
    # Для этой функции нужен подраздел TaskbarDeveloperSettings
    key_path = r"Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced\TaskbarDeveloperSettings"
    try:
        # Создаем ключ, если его нет (он может отсутствовать по умолчанию)
        key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, key_path)

        val = 1 if enable else 0
        winreg.SetValueEx(key, "TaskbarEndTask", 0, winreg.REG_DWORD, val)

        state = "ДОБАВЛЕНА" if enable else "УБРАНА"
        print(f"\n[+] Кнопка 'Завершить задачу' {state} в меню панели задач.")
        winreg.CloseKey(key)

        # Тут иногда перезапуск не нужен сразу, но для надежности лучше предложить
        restart_explorer_generic()
    except Exception as e:
        print(f"\n[-] Ошибка реестра: {e}")
        input("Enter...")


# --- ЛОГИКА: ПУСК ---

def toggle_bing_search(enable):
    key_path = r"Software\Policies\Microsoft\Windows\Explorer"
    try:
        key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, key_path)
        if enable:
            try:
                winreg.DeleteValue(key, "DisableSearchBoxSuggestions")
                print("\n[+] Bing поиск ВКЛЮЧЕН.")
            except:
                print("\n[i] Уже включен.")
        else:
            winreg.SetValueEx(key, "DisableSearchBoxSuggestions", 0, winreg.REG_DWORD, 1)
            print("\n[+] Bing поиск ОТКЛЮЧЕН.")
        winreg.CloseKey(key)
        restart_explorer_generic()
    except Exception as e:
        print(f"\n[-] Ошибка: {e}")
        input("Enter...")


# --- ЛОГИКА: ВИЗУАЛ (ПРОЗРАЧНОСТЬ) ---

def download_and_install_mica():
    install_dir = r"C:\Program Files\VizFix"
    print("\n--- УСТАНОВКА ПРОЗРАЧНОСТИ ---")

    if not os.path.exists(install_dir):
        try:
            os.makedirs(install_dir)
        except Exception as e:
            print(f"[-] Ошибка создания папки: {e}")
            input("Enter...")
            return

    print("[1/4] Поиск обновления...")
    try:
        api_url = "https://api.github.com/repos/Maplespe/ExplorerBlurMica/releases/latest"
        req = urllib.request.Request(api_url, headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req) as url:
            data = json.loads(url.read().decode())

        download_url = None
        for asset in data['assets']:
            if asset['name'].endswith('.zip'):
                download_url = asset['browser_download_url']
                break

        if not download_url:
            print("[-] Ссылка не найдена.")
            input("Enter...")
            return

        print("[2/4] Скачивание...")
        zip_path = os.path.join(install_dir, "vizfix_update.zip")
        urllib.request.urlretrieve(download_url, zip_path)

        print("[3/4] Распаковка...")
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            zip_ref.extractall(install_dir)
        os.remove(zip_path)

        real_dll_path = ""
        for root, dirs, files in os.walk(install_dir):
            if "ExplorerBlurMica.dll" in files:
                real_dll_path = os.path.join(root, "ExplorerBlurMica.dll")
                break

        if not real_dll_path:
            print("[-] DLL не найден.")
            input("Enter...")
            return

        print(f"[4/4] Регистрация...")
        result = os.system(f'regsvr32 /s "{real_dll_path}"')

        if result == 0:
            print("\n[+] Прозрачность установлена!")
            restart_explorer_generic()
        else:
            print("[-] Ошибка регистрации.")
            input("Enter...")

    except Exception as e:
        print(f"[-] Ошибка: {e}")
        input("Enter...")


def remove_transparency():
    install_dir = r"C:\Program Files\VizFix"
    dll_name = "ExplorerBlurMica.dll"
    print("\n--- УДАЛЕНИЕ ПРОЗРАЧНОСТИ ---")

    found_dll = ""
    if os.path.exists(install_dir):
        for root, dirs, files in os.walk(install_dir):
            if dll_name in files:
                found_dll = os.path.join(root, dll_name)
                break

    if found_dll:
        print("[1/3] Unregister DLL...")
        os.system(f'regsvr32 /u /s "{found_dll}"')

    print("[2/3] Остановка Explorer...")
    os.system("taskkill /f /im explorer.exe >nul 2>&1")
    time.sleep(1.5)

    print("[3/3] Удаление файлов...")
    try:
        if os.path.exists(install_dir):
            shutil.rmtree(install_dir)
            print("[+] Готово.")
        else:
            print("[i] Файлов нет.")
    except Exception as e:
        print(f"[-] Ошибка удаления: {e}")

    print("[*] Запуск Explorer...")
    os.system("start explorer.exe")
    input("\nНажмите Enter...")


# --- МЕНЮ И ПОДМЕНЮ ---

def menu_explorer():
    while True:
        clear_screen()
        print("=== VIZFIX > ПРОВОДНИК (EXPLORER) ===")
        print("\n[ Контекстное Меню ]")
        print("1. Включить CLASSIC (Win 10)")
        print("2. Включить MODERN (Win 11)")
        print("\n[ Расширения Файлов ]")
        print("3. ПОКАЗЫВАТЬ расширения (.py, .exe)")
        print("4. СКРЫВАТЬ расширения")
        print("\n[ Скрытые Элементы ]")
        print("5. ПОКАЗЫВАТЬ скрытые папки")
        print("6. СКРЫВАТЬ скрытые папки")
        print("\n[ Панель Задач ]")
        print("7. Добавить 'Завершить задачу' (Win 11 Only)")
        print("8. Убрать 'Завершить задачу' (Win 11 Only)")
        print("\n0. Назад")

        c = input("\nВыбор > ")
        if c == '1':
            enable_classic_menu()
        elif c == '2':
            enable_modern_menu()
        elif c == '3':
            toggle_file_extensions(True)
        elif c == '4':
            toggle_file_extensions(False)
        elif c == '5':
            toggle_hidden_items(True)
        elif c == '6':
            toggle_hidden_items(False)
        elif c == '7':
            toggle_taskbar_end_task(True)
        elif c == '8':
            toggle_taskbar_end_task(False)
        elif c == '0':
            return


def menu_start():
    while True:
        clear_screen()
        print("=== VIZFIX > МЕНЮ ПУСК (START MENU) ===")
        print("\n[ Поиск ]")
        print("1. Убрать Bing из поиска (Только локально)")
        print("2. Вернуть Bing в поиск (С интернетом)")
        print("\n0. Назад")

        c = input("\nВыбор > ")
        if c == '1':
            toggle_bing_search(False)
        elif c == '2':
            toggle_bing_search(True)
        elif c == '0':
            return


def menu_visuals():
    while True:
        clear_screen()
        print("=== VIZFIX > ВИЗУАЛЬНЫЕ ЭФФЕКТЫ ===")
        print("\n[ Прозрачность Окон ]")
        print("1. Установить прозрачность (Auto-Download)")
        print("2. Удалить прозрачность (Uninstall)")
        print("\n0. Назад")

        c = input("\nВыбор > ")
        if c == '1':
            download_and_install_mica()
        elif c == '2':
            remove_transparency()
        elif c == '0':
            return


def main_menu():
    while True:
        clear_screen()
        print("========================================")
        print("            VIZFIX CONSOLE              ")
        print("       Visual Fixer for Windows         ")
        print("========================================")
        print(f"Система: {platform.system()} {platform.release()}")

        print("\nГЛАВНОЕ МЕНЮ:")
        print("[1] Проводник (Меню, Файлы, Панель задач)")
        print("[2] Меню Пуск (Поиск Bing)")
        print("[3] Визуал (Прозрачность)")
        print("\n[0] ВЫХОД")

        choice = input("\nVizFix > ")

        if choice == '1':
            menu_explorer()
        elif choice == '2':
            menu_start()
        elif choice == '3':
            menu_visuals()
        elif choice == '0':
            print("Выход...")
            time.sleep(0.5)
            sys.exit()


if __name__ == "__main__":
    if not is_admin():
        print("[!] Запуск от имени администратора...")
        try:
            ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, " ".join(sys.argv), None, 1)
        except:
            print("[-] Ошибка прав доступа.")
            input()
    else:
        try:
            main_menu()
        except KeyboardInterrupt:
            sys.exit()
