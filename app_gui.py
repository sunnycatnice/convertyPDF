import tkinter as tk
from tkinter import filedialog, messagebox
import subprocess
import os
import platform
from workers.converter import convert_word_to_pdf

# Function to get the default desktop path based on the OS
def get_desktop_path():
    if platform.system() == "Windows":
        return os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
    elif platform.system() == "Darwin":
        return os.path.join(os.path.join(os.path.expanduser('~')), 'Desktop')
    else:
        return os.path.join(os.path.expanduser("~"), "Desktop")

# Function to convert files
def convert_files(print_after_conversion=False, formatted=False):
    files = filedialog.askopenfilenames(filetypes=[("Word files", "*.docx")])
    if not files:
        return

    save_path = filedialog.askdirectory(initialdir=get_desktop_path())
    if not save_path:
        save_path = get_desktop_path()
    
    success_files = []
    failed_files = []
    status_text.config(state=tk.NORMAL)  # Allow changes to status_text
    status_text.delete("1.0", tk.END)  # Clear previous status messages

    for file in files:
        try:
            convert_word_to_pdf([file], save_path, print_after_conversion, formatted)
            pdf_file = os.path.join(save_path, os.path.basename(file).replace('.docx', '.pdf'))
            if os.path.exists(pdf_file):
                success_files.append(pdf_file)
                status_text.insert(tk.END, f"✔ Successfully converted: {pdf_file}\n", "success")
            else:
                raise FileNotFoundError(f"Converted file {pdf_file} not found")
        except Exception as e:
            failed_files.append(file)
            status_text.insert(tk.END, f"✘ Conversion failed: {file}\nError: {str(e)}\n", "error")
    
    status_text.config(state=tk.DISABLED)  # Make status_text read-only

    if success_files:
        messagebox.showinfo("Conversion Successful", f"Successfully converted files:\n" + "\n".join(success_files))
    if failed_files:
        messagebox.showerror("Conversion Failed", f"Failed to convert files:\n" + "\n".join(failed_files))

app = tk.Tk()
app.title("ConvertifyPDF")
app.geometry("600x400")

label = tk.Label(app, text="Select Word Files to Convert to PDF")
label.pack(pady=10)

btn_convert = tk.Button(app, text="Convert", command=lambda: convert_files())
btn_convert.pack(pady=5)

btn_convert_and_print = tk.Button(app, text="Convert and Print", command=lambda: convert_files(print_after_conversion=True))
btn_convert_and_print.pack(pady=5)

btn_convert_formatted = tk.Button(app, text="Convert with Formatting", command=lambda: convert_files(formatted=True))
btn_convert_formatted.pack(pady=5)

btn_convert_formatted_and_print = tk.Button(app, text="Convert with Formatting and Print", command=lambda: convert_files(print_after_conversion=True, formatted=True))
btn_convert_formatted_and_print.pack(pady=5)

status_text = tk.Text(app, height=15, width=70)
status_text.pack(pady=10)
status_text.tag_config("success", foreground="green")
status_text.tag_config("error", foreground="red")
status_text.config(state=tk.DISABLED)  # Make status_text read-only

app.mainloop()
