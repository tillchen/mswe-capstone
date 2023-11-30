import tkinter as tk
from tkinter import filedialog

from fileProcessor import file_process, save_file

formatted_doc = None


def import_file():
    global formatted_doc, status_label
    file_path = filedialog.askopenfilename()

    if file_path:  # Check if a file was selected
        name = filename_entry.get()
        status_label.config(text="Processing...")
        root.update()  # Update the GUI to reflect the new label before processing starts
        formatted_doc = file_process(file_path, name)
        status_label.config(text="Import complete. File ready for export.")
        print(f"Imported file: {file_path}")
    else:
        status_label.config(text="No file selected. Please select a file.")


def export_file():
    default_name = filename_entry.get() + ".docx"
    file_path = filedialog.asksaveasfilename(defaultextension=".docx", initialfile=default_name)

    if file_path:  # Check if a path was provided
        global formatted_doc
        save_file(formatted_doc, file_path)
        status_label.config(text="Export successful.")
        print(f"Exported file: {file_path}")
    else:
        status_label.config(text="Export cancelled.")


# Create the main window
root = tk.Tk()
root.title("File Import/Export")
root.geometry("400x200")  # Adjust the size as needed

# Create a label for status messages
status_label = tk.Label(root, text="Please enter the title before importing your file", height=2)
status_label.pack(pady=5)

# Create an Entry widget for the filename
filename_entry = tk.Entry(root)
filename_entry.pack(pady=5)

# Create a frame to hold the buttons
button_frame = tk.Frame(root)
button_frame.pack(pady=10, padx=10)

# Create buttons with specified size
import_button = tk.Button(button_frame, text="Import File", command=import_file, height=2, width=15)
export_button = tk.Button(button_frame, text="Export File", command=export_file, height=2, width=15)

# Pack buttons to the left and right
import_button.pack(side=tk.LEFT, padx=5)
export_button.pack(side=tk.RIGHT, padx=5)

# Run the application
root.mainloop()
