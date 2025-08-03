import pyautogui
import pyperclip
import time
import threading
import tkinter as tk
import os


stop_flag = False

from transformers import MarianMTModel, MarianTokenizer

# Load Marian NMT model and tokenizer
model_name = "Helsinki-NLP/opus-mt-en-ar"
tokenizer = MarianTokenizer.from_pretrained(model_name)
model = MarianMTModel.from_pretrained(model_name)

def translate_text(text):
    # Tokenize input text
    inputs = tokenizer(text, return_tensors="pt", padding=True, truncation=True)
    
    # Run translation
    translated = model.generate(**inputs)
    
    # Decode result
    return tokenizer.decode(translated[0], skip_special_tokens=True)


def main_loop(iterations=10, delay=1.5):
    global stop_flag
    print("Starting translation loop...")
    time.sleep(5)  # Give you time to focus Excel

    for i in range(iterations):
        if stop_flag:
            print("Translation stopped.")
            break

        print(f"\nProcessing row {i+1}")

        pyperclip.copy('')  # Clear clipboard
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

# GUI control panel
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
    root.geometry("160x100+100+100")
    root.attributes("-topmost", True)
    root.resizable(False, False)

    tk.Button(root, text="▶ Start", width=15, command=start_translation, bg="lightgreen").pack(pady=5)
    tk.Button(root, text="⏹ Stop", width=15, command=stop_translation, bg="tomato").pack(pady=5)

    root.mainloop()

if __name__ == "__main__":
    create_floating_window()


