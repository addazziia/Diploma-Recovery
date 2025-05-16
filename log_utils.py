import os
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox

# Path to the log file
LOG_FILE = "activity.log"

def write_log(message):
    """
    Append a log entry with a timestamp to the log file.
    This is used to track user actions and system events.
    """
    timestamp = datetime.now().strftime("[%Y-%m-%d %H:%M:%S]")
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(f"{timestamp} {message}\n")


def show_log_window():
    """
    Open a separate GUI window to display the contents of the log file.
    Useful for auditing actions performed during a forensic session.
    """
    if not os.path.exists(LOG_FILE):
        messagebox.showerror("Error", "No logs found")
        return

    log_window = tk.Toplevel()
    log_window.title("Log Viewer")
    log_window.geometry("600x400")
    log_window.configure(bg="#2a2a3d")

    with open(LOG_FILE, "r", encoding="utf-8") as f:
        content = f.read()

    text_box = tk.Text(log_window, bg="#2a2a3d", fg="white", wrap="word")
    text_box.insert(tk.END, content)
    text_box.pack(expand=True, fill="both")


def export_log():
    """
    Allow the user to export the current log file to a selected location.
    The file is saved as a .txt document and can be shared for reporting.
    """
    if not os.path.exists(LOG_FILE):
        messagebox.showerror("Error", "No logs to export")
        return

    save_path = filedialog.asksaveasfilename(
        title="Export Logs As", defaultextension=".txt",
        filetypes=[("Text Files", "*.txt")]
    )
    if save_path:
        with open(LOG_FILE, "r", encoding="utf-8") as f:
            content = f.read()
        with open(save_path, "w", encoding="utf-8") as out:
            out.write(content)
        messagebox.showinfo("Success", f"Logs exported to: {save_path}")
