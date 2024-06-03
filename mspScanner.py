import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, filedialog, simpledialog
import winreg
import openpyxl
from openpyxl.styles import Font
import csv
import logging
from concurrent.futures import ThreadPoolExecutor, as_completed
import json
import os
import threading
import time
from datetime import datetime, timedelta

# Constants
FAVORITES_FILE = "favorites.json"
LOG_FILE = "previous_searches.json"
LOG_EXPIRATION_DAYS = 7

# Initialize logger
logging.basicConfig(
    level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s"
)


# Function to convert a string to different variations (case-insensitive, underscores, etc.)
def get_string_variations(input_string):
    variations = set()
    variations.add(input_string)
    variations.add(input_string.lower())
    variations.add(input_string.upper())
    variations.add(input_string.replace("_", " "))
    variations.add(input_string.replace(" ", "_"))
    variations.add("".join([word.capitalize() for word in input_string.split("_")]))
    variations.add("".join([word.capitalize() for word in input_string.split()]))
    return variations


# Function to search the registry
def search_registry(root, search_terms, root_name):
    results = []
    try:
        for i in range(0, winreg.QueryInfoKey(root)[0]):
            sub_key = winreg.EnumKey(root, i)
            sub_key_path = f"{root_name}\\{sub_key}"
            try:
                sub_key_handle = winreg.OpenKey(root, sub_key)
                results.extend(
                    search_registry(sub_key_handle, search_terms, sub_key_path)
                )
                winreg.CloseKey(sub_key_handle)
            except WindowsError as e:
                logging.error(f"Failed to open sub_key {sub_key_path}: {e}")
    except WindowsError as e:
        logging.error(f"Failed to enumerate sub_keys for {root_name}: {e}")
    try:
        for i in range(0, winreg.QueryInfoKey(root)[1]):
            value = winreg.EnumValue(root, i)
            for term in search_terms:
                if term in value[0] or term in str(value[1]):
                    results.append((root_name, value[0], value[1], "Registry value"))
    except WindowsError as e:
        logging.error(f"Failed to enumerate values for {root_name}: {e}")
    return results


# Function to initiate search
def start_search():
    search_input = entry.get()
    if not search_input:
        messagebox.showwarning("Input Error", "Please provide a search input.")
        return

    search_button.config(state=tk.DISABLED)
    search_terms = get_string_variations(search_input)
    search_type = search_type_var.get()
    scan_scope = scan_scope_var.get()
    results = []

    def search():
        try:
            if search_type == "Registry":
                if scan_scope == "whole_device":
                    root_keys = [
                        winreg.HKEY_LOCAL_MACHINE,
                        winreg.HKEY_CURRENT_USER,
                        winreg.HKEY_CLASSES_ROOT,
                        winreg.HKEY_USERS,
                        winreg.HKEY_CURRENT_CONFIG,
                    ]
                else:
                    root_keys = [winreg.HKEY_CURRENT_USER]

                root_names = {
                    winreg.HKEY_LOCAL_MACHINE: "HKLM",
                    winreg.HKEY_CURRENT_USER: "HKCU",
                    winreg.HKEY_CLASSES_ROOT: "HKCR",
                    winreg.HKEY_USERS: "HKU",
                    winreg.HKEY_CURRENT_CONFIG: "HKCC",
                }

                with ThreadPoolExecutor(max_workers=5) as executor:
                    future_to_root = {
                        executor.submit(
                            search_registry,
                            winreg.OpenKey(root_key, ""),
                            search_terms,
                            root_names[root_key],
                        ): root_key
                        for root_key in root_keys
                    }
                    for future in as_completed(future_to_root):
                        try:
                            results.extend(future.result())
                        except Exception as e:
                            logging.error(f"Error occurred during registry search: {e}")

            display_results(results)
            global current_results
            current_results = results
            log_search(search_type, search_input, results)
        except Exception as e:
            logging.error(f"Search failed: {e}")
            messagebox.showerror(
                "Search Error", f"An error occurred during the search: {e}"
            )
        finally:
            search_button.config(state=tk.NORMAL)
            stop_animation.set()

    threading.Thread(target=search).start()
    threading.Thread(target=animate_search).start()


# Function to display results in the GUI
def display_results(results):
    results_text.config(state=tk.NORMAL)
    results_text.delete(1.0, tk.END)

    if not results:
        results_text.insert(tk.END, "No matching registry keys or values found.")
    else:
        for result in results:
            path, entry, value, description = result
            if isinstance(value, bytes):
                try:
                    value = value.decode("utf-8")
                except UnicodeDecodeError:
                    value = f"Binary data: {value}"
            elif (
                isinstance(value, str) and value.startswith("[") and value.endswith("]")
            ):
                try:
                    value = json.loads(value)
                    value = json.dumps(value, indent=4)
                except json.JSONDecodeError:
                    pass  # leave value as is if it can't be parsed as JSON

            results_text.insert(tk.END, f"Path: {path}\n")
            results_text.insert(tk.END, f"Entry: {entry}\n")
            results_text.insert(tk.END, f"Value: {value}\n")
            results_text.insert(tk.END, f"Description: {description}\n")
            results_text.insert(tk.END, "\n" + "=" * 80 + "\n\n")

    results_text.config(state=tk.DISABLED)
    export_button.config(state=tk.NORMAL)


# Function to save results to an Excel file
def save_to_excel(results, file_path):
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Scan Results"
    headers = ["Path", "Entry", "Value", "Description"]
    sheet.append(headers)
    for col in range(1, 5):
        sheet.cell(row=1, column=col).font = Font(bold=True)
    for result in results:
        sheet.append(result)
    workbook.save(file_path)
    messagebox.showinfo("Save Success", f"Results have been saved to {file_path}")


# Function to save results to a CSV file
def save_to_csv(results, file_path):
    with open(file_path, mode="w", newline="", encoding="utf-8") as file:
        writer = csv.writer(file)
        writer.writerow(["Path", "Entry", "Value", "Description"])
        writer.writerows(results)
    messagebox.showinfo("Save Success", f"Results have been saved to {file_path}")


# Function to export results
def export_results():
    if not current_results:
        messagebox.showwarning("No Results", "No results to export.")
        return
    filetypes = [("Excel files", "*.xlsx"), ("CSV files", "*.csv")]
    file_path = filedialog.asksaveasfilename(
        defaultextension=".xlsx", filetypes=filetypes
    )
    if file_path:
        if file_path.endswith(".xlsx"):
            save_to_excel(current_results, file_path)
        elif file_path.endswith(".csv"):
            save_to_csv(current_results, file_path)


# Function to save a favorite search
def save_favorite_search():
    search_input = entry.get()
    search_type = search_type_var.get()
    if not search_input:
        messagebox.showwarning("Input Error", "Please provide a search input.")
        return
    favorite_name = simpledialog.askstring(
        "Save Favorite Search", "Enter a name for this search:"
    )
    if favorite_name:
        favorite_search = {
            "name": favorite_name,
            "type": search_type,
            "input": search_input,
        }
        if os.path.exists(FAVORITES_FILE):
            with open(FAVORITES_FILE, "r") as file:
                favorites = json.load(file)
        else:
            favorites = []
        favorites.append(favorite_search)
        with open(FAVORITES_FILE, "w") as file:
            json.dump(favorites, file)
        messagebox.showinfo(
            "Save Success", f"Search '{favorite_name}' has been saved to favorites."
        )


# Function to load a favorite search
def load_favorite_search():
    if os.path.exists(FAVORITES_FILE):
        with open(FAVORITES_FILE, "r") as file:
            favorites = json.load(file)
        favorite_names = [fav["name"] for fav in favorites]
        selected_favorite = simpledialog.askstring(
            "Load Favorite Search",
            "Enter the name of the favorite search to load:",
            initialvalue=favorite_names[0] if favorite_names else "",
        )
        for favorite in favorites:
            if favorite["name"] == selected_favorite:
                search_type_var.set(favorite["type"])
                entry.delete(0, tk.END)
                entry.insert(0, favorite["input"])
                break
        else:
            messagebox.showwarning("Load Error", "Favorite search not found.")
    else:
        messagebox.showwarning("Load Error", "No favorite searches found.")


# Function to log a search
def log_search(search_type, search_input, results):
    if os.path.exists(LOG_FILE):
        with open(LOG_FILE, "r") as file:
            log = json.load(file)
    else:
        log = []

    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    log_entry = {
        "timestamp": timestamp,
        "type": search_type,
        "input": search_input,
        "results": results,
    }
    log.append(log_entry)

    # Remove expired log entries
    cutoff_time = datetime.now() - timedelta(days=LOG_EXPIRATION_DAYS)
    log = [
        entry
        for entry in log
        if datetime.strptime(entry["timestamp"], "%Y-%m-%d %H:%M:%S") > cutoff_time
    ]

    with open(LOG_FILE, "w") as file:
        json.dump(log, file)


# Function to view previous searches
def view_previous_searches():
    if os.path.exists(LOG_FILE):
        with open(LOG_FILE, "r") as file:
            log = json.load(file)
        log_text = "\n\n".join(
            f"Timestamp: {entry['timestamp']}\nType: {entry['type']}\nInput: {entry['input']}\nResults Count: {len(entry['results'])}"
            for entry in log
        )
        messagebox.showinfo(
            "Previous Searches", log_text if log_text else "No previous searches found."
        )
    else:
        messagebox.showinfo("Previous Searches", "No previous searches found.")


# Function to animate the search
def animate_search():
    animation = ["|", "/", "-", "\\"]
    idx = 0
    while not stop_animation.is_set():
        status_label.config(text=f"Searching... {animation[idx]}")
        idx = (idx + 1) % len(animation)
        time.sleep(0.1)
    status_label.config(text="Search Complete")


# Set up the main GUI window
root = tk.Tk()
root.title("Enhanced Scanner")
frame = ttk.Frame(root, padding="10")
frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

ttk.Label(frame, text="Search Type:").grid(row=0, column=0, sticky=tk.W)
search_type_var = tk.StringVar()
search_type_combobox = ttk.Combobox(
    frame, textvariable=search_type_var, values=["Registry"], state="readonly"
)
search_type_combobox.grid(row=0, column=1, sticky=(tk.W, tk.E))
search_type_combobox.current(0)

ttk.Label(frame, text="Search Input:").grid(row=1, column=0, sticky=tk.W)
entry = ttk.Entry(frame, width=50)
entry.grid(row=1, column=1, sticky=(tk.W, tk.E))

# Scan Scope
scan_scope_var = tk.StringVar(value="whole_device")
ttk.Label(frame, text="Scan Scope:").grid(row=0, column=2, sticky=tk.W)
ttk.Radiobutton(
    frame, text="Whole Device", variable=scan_scope_var, value="whole_device"
).grid(row=0, column=3, sticky=tk.W)
ttk.Radiobutton(
    frame, text="Current Profile", variable=scan_scope_var, value="current_profile"
).grid(row=0, column=4, sticky=tk.W)

# Add search progress bar
progress = ttk.Progressbar(frame, orient="horizontal", length=200, mode="determinate")
progress.grid(row=2, column=1, sticky=(tk.W, tk.E))

# Placeholder for advanced search options button
advanced_options_button = ttk.Button(
    frame,
    text="Advanced Options",
    command=lambda: messagebox.showinfo(
        "Advanced Options", "This feature is not implemented yet."
    ),
)
advanced_options_button.grid(row=2, column=0, sticky=tk.W)

search_button = ttk.Button(frame, text="Start Search", command=start_search)
search_button.grid(row=2, column=2, sticky=tk.E)

ttk.Button(frame, text="Save Favorite Search", command=save_favorite_search).grid(
    row=2, column=3, sticky=tk.W
)
ttk.Button(frame, text="Load Favorite Search", command=load_favorite_search).grid(
    row=2, column=4, sticky=tk.W
)
ttk.Button(frame, text="View Previous Searches", command=view_previous_searches).grid(
    row=2, column=5, sticky=tk.W
)

results_text = scrolledtext.ScrolledText(
    frame, width=100, height=30, wrap=tk.WORD, state=tk.DISABLED
)
results_text.grid(row=3, column=0, columnspan=6, sticky=(tk.W, tk.E, tk.N, tk.S))

export_button = ttk.Button(
    frame, text="Export Results", command=export_results, state=tk.DISABLED
)
export_button.grid(row=4, column=5, sticky=tk.E)

# Status label for animation
status_label = ttk.Label(frame, text="")
status_label.grid(row=4, column=0, columnspan=4, sticky=tk.W)

# Event to stop the animation
stop_animation = threading.Event()

root.columnconfigure(0, weight=1)
root.rowconfigure(0, weight=1)
frame.columnconfigure(1, weight=1)
frame.rowconfigure(3, weight=1)

current_results = []
root.mainloop()
