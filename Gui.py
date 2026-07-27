import tkinter as tk
from tkinter import filedialog, messagebox
from final_project import collect_data, export_to_excel

file_path = ""


def browse_file():
    global file_path
    file_path = filedialog.askopenfilename(
        title="Select Excel File",
        filetypes=[("Excel files", "*.xlsx")]
    )
    if file_path:
        file_label.config(text=file_path)


def run_script():
    if not file_path:
        messagebox.showerror("Error", "Please select a file first")
        return

    new_version = version_entry.get().strip()

    if not new_version:
        messagebox.showerror("Error", "Please enter the required version")
        return

    try:
        results = collect_data(file_path, new_version)

        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")]
        )

        if save_path:
            export_to_excel(results, save_path)
            messagebox.showinfo("Success", "File saved successfully!")

    except Exception as e:
        messagebox.showerror("Error", str(e))


# GUI Setup
root = tk.Tk()
root.title("Router Version Checker")
root.geometry("400x250")

# Browse File
browse_btn = tk.Button(root, text="Browse Excel File", command=browse_file)
browse_btn.pack(pady=10)

file_label = tk.Label(root, text="No file selected")
file_label.pack()

# Version Entry
version_label = tk.Label(root, text="Enter Required Version:")
version_label.pack(pady=10)

version_entry = tk.Entry(root, width=30)
version_entry.pack()

# Run Button
run_btn = tk.Button(root, text="Run", command=run_script, bg="green", fg="white")
run_btn.pack(pady=20)

root.mainloop()