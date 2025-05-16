import tkinter as tk
from tkinter import filedialog, messagebox
import os
import json
import joblib
import matplotlib.pyplot as plt
from wordcloud import WordCloud
from datetime import datetime

# Import core modules
from slot_scanner import scan_and_extract_fragments
from slot_recovery import recover_docx_from_fragment
from analysis import analyze_files
from slot_image_recovery import extract_images_from_all_docx
from log_utils import write_log, export_log, show_log_window

log_box = None
selected_dump_path = ""
slots_stats = (0, 0)

#  UI Styling 
BG_COLOR = "#1e1e2f"
FG_COLOR = "#ffffff"
BUTTON_COLOR = "#4CAF50"
FONT = ("Segoe UI", 10)

#  Step 1: Select dump file 
def browse_file(entry):
    """Select disk dump file (.001/.bin) and update path."""
    global selected_dump_path
    selected_dump_path = filedialog.askopenfilename(filetypes=[("Dump Files", "*.001 *.bin *")])
    entry.delete(0, tk.END)
    entry.insert(0, selected_dump_path)
    msg = f"Selected dump file: {selected_dump_path}"
    log_box.insert(tk.END, f"[✓] {msg}\n")
    write_log(msg)

# Step 2: Scan and extract DOCX signature fragments 
def extract_fragments():
    """Scan dump and extract valid fragments that contain .docx headers."""
    global slots_stats
    if not selected_dump_path:
        messagebox.showerror("Error", "No dump file selected")
        return
    valid, total = scan_and_extract_fragments(selected_dump_path)
    slots_stats = (valid, total)
    msg = f"Extracted {valid} valid fragments from {total} scanned slots."
    log_box.insert(tk.END, f"[✓] {msg}\n")
    write_log(msg)
    messagebox.showinfo("Done", msg)

# Step 3: Recover files from selected fragment 
def browse_fragment():
    """Allow user to select a .bin fragment and extract .docx files from it."""
    fragment = filedialog.askopenfilename(title="Select Fragment (.bin)", filetypes=[("Fragment", "*.bin")])
    if not fragment:
        return
    log_box.insert(tk.END, f"[✓] Selected fragment: {fragment}\n")
    write_log(f"Selected fragment: {fragment}")
    recovered_files = recover_docx_from_fragment(fragment)
    msg = f"Recovered {len(recovered_files)} valid .docx files from fragment."
    log_box.insert(tk.END, f"[✓] {msg}\n")
    write_log(msg)
    messagebox.showinfo("Recovery", msg)

#  Step 4: Analyze structure, encryption, and text 
def run_analysis():
    """Analyze .doc/.docx files and extract metadata and content."""
    recovered_path = "recovered_docs_from_fragment"
    if not os.path.exists(recovered_path):
        messagebox.showerror("Missing", "Recover documents before analysis")
        return
    files = [os.path.join(recovered_path, f) for f in os.listdir(recovered_path)
             if f.endswith(".docx") or f.endswith(".doc")]
    report = analyze_files(files)
    with open("forensic_report.json", "w", encoding="utf-8") as f:
        json.dump(report, f, indent=2, ensure_ascii=False)

    # Automatically extract images from all recovered files
    extract_images_from_all_docx()

    msg = "Forensic analysis completed and report saved"
    log_box.insert(tk.END, f"[✓] {msg}\n")
    write_log(msg)

# Step 5: Classify files using ML model (Not add)
def run_ml():
    """Apply machine learning model to classify file contents."""
    if not os.path.exists("forensic_report.json"):
        messagebox.showerror("Error", "Run analysis first")
        return
    with open("forensic_report.json", "r", encoding="utf-8") as f:
        report = json.load(f)
    texts = [info.get("extracted_text", "") for info in report.values()]
    names = list(report.keys())

    model = joblib.load("nb_model.pkl")
    vectorizer = joblib.load("tfidf_vectorizer.pkl")
    X = vectorizer.transform(texts)
    preds = model.predict(X)
    probs = model.predict_proba(X)

    for name, pred, prob in zip(names, preds, probs):
        report[name]["ml_prediction"] = pred
        report[name]["ml_confidence"] = max(prob)

    with open("forensic_report_updated.json", "w", encoding="utf-8") as f:
        json.dump(report, f, indent=2, ensure_ascii=False)
    log_box.insert(tk.END, "[✓] ML classification saved to forensic_report_updated.json\n")
    write_log("ML analysis completed")

# Step 6: Show word cloud (Not add)
def show_keywords():
    """Display word cloud from recovered texts."""
    with open("forensic_report_updated.json", "r", encoding="utf-8") as f:
        report = json.load(f)
    text = " ".join(info.get("extracted_text", "") for info in report.values())
    wc = WordCloud(width=800, height=400, background_color=BG_COLOR, colormap="viridis").generate(text)
    plt.imshow(wc, interpolation='bilinear')
    plt.axis("off")
    plt.title("Keyword Cloud")
    plt.show()

#  Step 7: Show dump coverage chart 
def show_slot_chart():
    """Pie chart showing percentage of valid vs. empty slots in dump."""
    if not slots_stats or slots_stats[1] == 0:
        messagebox.showerror("Error", "No slot data available")
        return
    found, total = slots_stats
    empty = total - found
    labels = ['Valid .docx Slots', 'Empty Slots']
    sizes = [found, empty]
    colors = ['#4CAF50', '#F44336']
    plt.figure(figsize=(5, 5))
    plt.pie(sizes, labels=labels, autopct='%1.1f%%', colors=colors, startangle=90)
    plt.title("DOCX Signature Density in Dump")
    plt.axis("equal")
    plt.show()

# Step 8: Export logs
def export_logs_gui():
    """Save internal log to file."""
    export_log()
    messagebox.showinfo("Exported", "Log file saved to exported_log.txt")

# GUI Launch 
def launch_gui():
    global log_box
    root = tk.Tk()
    root.title("Document Recovery Forensics")
    root.configure(bg=BG_COLOR)

    # File path input + Browse
    top_frame = tk.Frame(root, bg=BG_COLOR)
    top_frame.pack(pady=10)

    tk.Label(top_frame, text="Disk dump path:", bg=BG_COLOR, fg=FG_COLOR, font=FONT).pack(side=tk.LEFT)
    path_entry = tk.Entry(top_frame, width=60)
    path_entry.pack(side=tk.LEFT, padx=5)
    tk.Button(top_frame, text="Browse", command=lambda: browse_file(path_entry),
              bg=BUTTON_COLOR, fg="white").pack(side=tk.LEFT)

    # Steps and buttons
    def add_step(step_number, step_text, command):
        row = tk.Frame(root, bg=BG_COLOR)
        row.pack(pady=3, padx=60, fill="x")
        circle = tk.Label(row, text=str(step_number), font=("Segoe UI", 10, "bold"),
                          bg="white", fg="black", width=2, height=1, relief="solid", bd=1)
        circle.pack(side=tk.LEFT, padx=5)
        btn = tk.Button(row, text=step_text, command=command, font=FONT,
                        bg=BUTTON_COLOR, fg="white", padx=10, pady=4)
        btn.pack(side=tk.LEFT, fill="x", expand=True)

    add_step(1, "Extract DOCX Signatures and Save", extract_fragments)
    add_step(2, "Recover from Selected Fragment", browse_fragment)
    add_step(3, "Start Recovery and Analysis", run_analysis)
    add_step(4, "Run ML Analysis", run_ml)
    add_step(5, "Show Keywords Cloud", show_keywords)
    add_step(6, "Show Dump Signature Chart", show_slot_chart)
    add_step(7, "View Logs", show_log_window)
    add_step(8, "Export Logs", export_logs_gui)

    # Log output box
    log_box_frame = tk.Frame(root, bg=BG_COLOR)
    log_box_frame.pack(pady=10)
    log_box = tk.Text(log_box_frame, height=12, width=100, bg="#2a2a3d", fg="#d0d0d0")
    log_box.pack()

    root.mainloop()

#Main entry
if __name__ == "__main__":
    launch_gui()
