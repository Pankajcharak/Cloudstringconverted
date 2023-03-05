import tkinter as tk
import openpyxl
import os
from tkinter import ttk

def slice_string():
    my_string = entry.get()
    fixed_width = 198

    if save_as.get() == "Excel":
        wb = openpyxl.Workbook()
        ws = wb.active

        num_chunks = (len(my_string) + fixed_width - 1) // fixed_width
        progress_bar.config(maximum=num_chunks, value=0)
        progress_bar_title.config(text="Slicing String into Excel")

        for i in range(0, len(my_string), fixed_width):
            chunk = my_string[i:i+fixed_width]
            new_chunk=chunk[:-1]
            ws.append([new_chunk])

            progress_bar.step(1)
            root.update_idletasks()

        wb.save("output.xlsx")
        label.config(text="Output saved to output.xlsx")

    elif save_as.get() == "Text":
        with open("output.txt", "w") as f:
            num_chunks = (len(my_string) + fixed_width - 1) // fixed_width
            progress_bar.config(maximum=num_chunks, value=0)
            progress_bar_title.config(text="Slicing String into Text")

            for i in range(0, len(my_string), fixed_width):
                chunk = my_string[i:i+fixed_width]
                new_chunk=chunk[:-1]
                f.write(new_chunk + "\n")

                progress_bar.step(1)
                root.update_idletasks()

        label.config(text="Output saved to output.txt")

    if open_in_notepad.get():
        os.system("\"C:\\Program Files\\Notepad++\\notepad++.exe\" output.txt")

root = tk.Tk()
root.title("CONVERT DOC ID TO GROUP OF 6")

label = tk.Label(root, text="Enter DOC ID from Notepad++")
label.pack()

entry_frame = tk.Frame(root)
entry_frame.pack()

entry_label = tk.Label(entry_frame, text="DOC ID:")
entry_label.pack(side=tk.LEFT)

entry = tk.Entry(entry_frame)
entry.pack(side=tk.LEFT)

save_as = tk.StringVar(value="Excel")

excel_button = tk.Radiobutton(root, text="Save as Excel", variable=save_as, value="Excel")
excel_button.pack()

text_button = tk.Radiobutton(root, text="Save as Text", variable=save_as, value="Text")
text_button.pack()

open_in_notepad = tk.BooleanVar(value=False)

notepad_button = tk.Checkbutton(root, text="Open in Notepad++", variable=open_in_notepad)
notepad_button.pack()

button_frame = tk.Frame(root)
button_frame.pack()

button = tk.Button(button_frame, text="Slice String", command=slice_string)
button.pack(side=tk.LEFT)

progress_bar_title = tk.Label(root, text="")
progress_bar_title.pack()

progress_bar = ttk.Progressbar(root, orient="horizontal", length=200, mode="determinate")
progress_bar.pack()

label = tk.Label(root, text="")
label.pack()

root.mainloop()
