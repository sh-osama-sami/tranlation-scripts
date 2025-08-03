import pyautogui
import pyperclip
import time



def main_loop(iterations=10, delay=1.5):
    print("Starting translation loop...")
    time.sleep(5)  # Give you time to switch to the desktop app

    for i in range(iterations):
        print(f"\nProcessing row {i+1}")

        pyautogui.hotkey('ctrl', 'enter') # Confirm the current cell

        time.sleep(delay)
            

if __name__ == "__main__":
    main_loop(iterations=20, delay=1.5)
