import time
import tkinter as tk
from datetime import timedelta
import pygetwindow as gw
import reprlib
TRACKED_APP = "edge"
TRACKER_WINDOW_TITLE = "App Time Tracker"

def get_active_window_title():
    try:
        window = gw.getActiveWindow()
        if window is None:
            return ""
        title = window.title.strip().lower()
        if TRACKER_WINDOW_TITLE.lower() in title:
            return ""  # Ignore the tracker window
        return title
    except Exception:
        return ""

def format_time(seconds):
    return str(timedelta(seconds=int(seconds)))

class TimeTrackerApp:
    def __init__(self, root):
        self.root = root
        self.root.title(TRACKER_WINDOW_TITLE)
        self.root.geometry("300x150+50+50")
        self.root.resizable(False, False)
        self.root.attributes("-topmost", True)  # Optional: keep on top

        self.label_app = tk.Label(root, text=f"Tracking: {TRACKED_APP}", font=("Helvetica", 12))
        self.label_app.pack(pady=10)

        self.time_label = tk.Label(root, text="00:00:00", font=("Helvetica", 24), fg="green")
        self.time_label.pack(pady=10)

        self.status_label = tk.Label(root, text="Paused", font=("Helvetica", 10), fg="red")
        self.status_label.pack()

        self.total_seconds = 0
        self.last_checked = time.time()
        self.tracking = False

        self.update_timer()

    def update_timer(self):
        current_time = time.time()
        active_window = get_active_window_title()
        print("DEBUG ACTIVE WINDOW:", reprlib.repr(active_window))

        if TRACKED_APP in active_window:
            if not self.tracking:
                self.status_label.config(text="Tracking", fg="green")
                self.tracking = True
            self.total_seconds += current_time - self.last_checked
        else:
            if self.tracking:
                self.status_label.config(text="Paused", fg="red")
                self.tracking = False

        self.last_checked = current_time
        self.time_label.config(text=format_time(self.total_seconds))
        self.root.after(1000, self.update_timer)

def main():
    root = tk.Tk()
    app = TimeTrackerApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
