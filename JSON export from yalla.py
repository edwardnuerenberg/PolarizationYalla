# ----------------------------
# Database (JSON) Import/Export
# ----------------------------
def export_database():
    """Export the in-memory database (data_dict) as JSON (DataFrames -> dicts)."""
    if not data_dict:
        messagebox.showinfo("Info", "No data available to export.")
        return

    file_path = filedialog.asksaveasfilename(
        defaultextension=".json",
        filetypes=[("JSON files", "*.json"), ("All Files", "*.*")]
    )
    if not file_path:
        return

    try:
        with open(file_path, "w", encoding="utf-8") as f:
            json.dump({k: v.to_dict() for k, v in data_dict.items()}, f, ensure_ascii=False)
        messagebox.showinfo("Success", f"Database exported successfully to:\n{file_path}")
        set_status(f"Database exported: {os.path.basename(file_path)}")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to export database:\n{e}")

def load_database():
    """Load a previously exported JSON database into data_dict."""
    global data_dict
    file_path = filedialog.askopenfilename(
        filetypes=[("JSON files", "*.json"), ("All Files", "*.*")]
    )
    if not file_path:
        return

    try:
        with open(file_path, "r", encoding="utf-8") as f:
            data = json.load(f)
        data_dict = {k: pd.DataFrame(v) for k, v in data.items()}
        key_dropdown['values'] = list(data_dict.keys())
        messagebox.showinfo("Success", "Database loaded successfully.")
        set_status(f"Database loaded: {os.path.basename(file_path)} | keys={len(data_dict)}")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to load database:\n{e}")
