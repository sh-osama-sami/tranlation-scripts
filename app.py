import pyautogui
import pyperclip
from googletrans import Translator
import time

translator = Translator()

def translate_text(text, dest_language='ar'):
    try:
        translated = translator.translate(text, dest=dest_language)
        return translated.text
    except Exception as e:
        print(f"Translation failed: {e}")
        return ''

def main_loop(iterations=10, delay=1.5):
    print("Starting translation loop...")
    time.sleep(5)  # Give you time to switch to the desktop app

    for i in range(iterations):
        print(f"\nProcessing row {i+1}")

        # Copy the original text from the current cell
        pyautogui.hotkey('ctrl', 'c')
        time.sleep(0.5)

        original_text = pyperclip.paste().strip()
        print(f"Original: {original_text}")

        if original_text:
            translated = translate_text(original_text)
            print(f"Translated: {translated}")

            # Move to the adjacent (right) cell
            pyautogui.press('tab')
            time.sleep(0.2)

            # Paste translated text
            pyperclip.copy(translated)
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(0.2)

            # Move to the start of the next row
            pyautogui.press('enter')
            pyautogui.hotkey('shift', 'tab')

            time.sleep(delay)
    

if __name__ == "__main__":
    main_loop(iterations=20, delay=1.5)
