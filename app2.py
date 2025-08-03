import pyautogui
import pyperclip
from googletrans import Translator
import time
import threading
import tkinter as tk

translator = Translator()
stop_flag = False

def translate_text(text, dest_language='ar'):
    try:
        translated = translator.translate(text, dest=dest_language)
        return translated.text
    except Exception as e:
        print(f"Translation failed: {e}")
        return ''

def main_loop(iterations=10, delay=1.5):
    global stop_flag
    print("Starting translation loop...")
    # time.sleep(5)  # Time to switch to Excel
    pyperclip.copy('')
    for i in range(iterations):
        if stop_flag:
            print("Translation stopped.")
            break

        print(f"\nProcessing row {i+1}")
        pyautogui.hotkey('ctrl', 'c')
        time.sleep(0.5)

        original_text = pyperclip.paste().strip()
        print(f"Original: {original_text}")

        if original_text:
            translated = translate_text(original_text)
            print(f"Translated: {translated}")

            pyautogui.press('tab')
            time.sleep(0.2)

            pyperclip.copy(translated)
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(0.2)

            pyautogui.press('enter')
            pyautogui.hotkey('shift', 'tab')

            time.sleep(delay)

# GUI Control Panel
def start_translation():
    global stop_flag
    stop_flag = False
    threading.Thread(target=main_loop, kwargs={'iterations': 100}, daemon=True).start()

def stop_translation():
    global stop_flag
    stop_flag = True

def create_floating_window():
    root = tk.Tk()
    root.title("Translator Control")
    root.geometry("160x100+100+100")  # Width x Height + X + Y
    root.attributes("-topmost", True)  # Always on top
    root.resizable(False, False)

    start_btn = tk.Button(root, text="▶ Start", width=15, command=start_translation, bg="lightgreen")
    start_btn.pack(pady=5)

    stop_btn = tk.Button(root, text="⏹ Stop", width=15, command=stop_translation, bg="tomato")
    stop_btn.pack(pady=5)

    root.mainloop()

if __name__ == "__main__":
    create_floating_window()
