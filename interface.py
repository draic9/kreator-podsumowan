import tkinter as tk
from tkinterdnd2 import DND_FILES, TkinterDnD
from tkinter import filedialog
import openpyxl
import os

# Function to handle dropped files
def drop(event):
    files = root.tk.splitlist(event.data)
    if files:
        file_path = files[0]
        try:
            if os.path.isfile(file_path):
                label_file.config(text="Wybrano plik: " + file_path)
                global input_file_path
                input_file_path = file_path
            else:
                raise FileNotFoundError("Wybrany plik nie istnieje.")
        except Exception as e:
            label_file.config(text="Błąd: " + str(e))

# Function to handle file browsing
def browse_file():
    file_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")])
    if file_path:
        try:
            if os.path.isfile(file_path):
                label_file.config(text="Wybrano plik: " + file_path)
                global input_file_path
                input_file_path = file_path
            else:
                raise FileNotFoundError("Wybrany plik nie istnieje.")
        except Exception as e:
            label_file.config(text="Błąd: " + str(e))

def confirm():
    try:
        if input_file_path:
            save_file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])
            if save_file_path:
                try:
                    # For demonstration, create an empty workbook and save it as .xlsx
                    workbook = openpyxl.Workbook()
                    workbook.save(save_file_path)
                    label_file.config(text="Plik zapisano pomyślnie.")
                except Exception as e:
                    label_file.config(text="Błąd: " + str(e))
            else:
                label_file.config(text="Operacje zapisywania anulowano")
        else:
            label_file.config(text="Nie wybrano żadnego pliku.")
    except Exception as e:
        label_file.config(text="Błąd: " + str(e))

def cancel():
    label_file.config(text="Nie wybrano jeszcze pliku.")
    global input_file_path
    input_file_path = ""

# Create the main window
root = TkinterDnD.Tk()

# Set the title of the window
root.title("Kreator podsumowań")

# Set the background color of the window to white
root.configure(bg="white")

# Set the window size
root.geometry("400x350")

# Disable window resizing
root.resizable(False, False)
    
# Create a frame for the drag-and-drop area
frame = tk.Frame(root, bg="white smoke", bd=2, relief=tk.SUNKEN, width=350, height=200)
frame.pack(pady=(20, 10), padx=20)
frame.pack_propagate(False)  # Prevent the frame from resizing with its content

# Create a label inside the frame for instructions
label_drag = tk.Label(frame, text="Przeciągnij plik tutaj", font=("Helvetica", 14), fg="black", bg="white smoke")
label_drag.pack(pady=(40, 10))

# Create a label for "OR" with background
label_or = tk.Label(frame, text="LUB", font=("Helvetica", 12), fg="dark gray", bg="white smoke")
label_or.pack(pady=5)

# Create a horizontal line below the "OR" label
canvas_line = tk.Canvas(frame, width=300, height=2, bd=0, highlightthickness=0)
canvas_line.create_line(10, 1, 290, 1, fill="dark gray")
canvas_line.pack()

# Place the line on top of the "OR" label
canvas_line.place(in_=label_or, relx=0.5, rely=0.5, anchor="s")

# Bring the "OR" label to the front
label_or.lift()

# Create a label for "Browse file"
label_browse = tk.Label(frame, text="Wybierz z folderu", font=("Helvetica", 12), fg="blue", bg="white smoke", cursor="hand2")
label_browse.pack(pady=(10, 40))

# Create a label to display the selected file
label_file = tk.Label(root, text="Nie wybrano jeszcze pliku.", font=("Helvetica", 10), wraplength=320, fg="black", bg="white")
label_file.pack(pady=10)

# Create a frame for the buttons
button_frame = tk.Frame(root, bg="white")
button_frame.pack(fill=tk.X, padx=25, pady=10)

# Create buttons
confirm_button = tk.Button(button_frame, text="Zatwierdź", command=confirm, bg="pale green", fg="black", padx=10)
confirm_button.pack(side=tk.RIGHT)

cancel_button = tk.Button(button_frame, text="Anuluj", command=cancel, bg="lightgray", fg="black", padx=10)
cancel_button.pack(side=tk.RIGHT, padx=5)

# Change label color on hover
def on_enter(event):
    label_browse.config(fg="red", font=("Helvetica", 12, "underline"))

def on_leave(event):
    label_browse.config(fg="blue", font=("Helvetica", 12))

# Enable drag and drop functionality
root.drop_target_register(DND_FILES)
root.dnd_bind('<<Drop>>', drop)

# Bind the browse label to the file dialog
label_browse.bind("<Button-1>", lambda e: browse_file())

# Bind hover events to change the color and underline
label_browse.bind("<Enter>", on_enter)
label_browse.bind("<Leave>", on_leave)

# Run the main loop
root.mainloop()
