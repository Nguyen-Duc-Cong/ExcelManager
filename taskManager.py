import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import pandas as pd

window = tk.Tk()
window.geometry("1000x500")
window.pack_propagate(False)
window.resizable(0, 0)

# ////////////////////////////////
frame_data = tk.LabelFrame(window, text="Excel Data")
frame_data.place(height=250, width=1000)

file_frame = tk.LabelFrame(window, text="Open File")
file_frame.place(height=100, width=1000, rely=0.65, relx=0)

button_browse = tk.Button(file_frame, text="Browse A File", command=lambda: File_dialog())
button_browse.place(rely=0.65, relx=0.50)

button_load = tk.Button(file_frame, text="Load File", command=lambda: Load_excel_data())
button_load.place(rely=0.65, relx=0.30)

label_file = ttk.Label(file_frame, text="No File Selected")
label_file.place(rely=0, relx=0)

tv1 = ttk.Treeview(frame_data)
tv1.place(relheight=1, relwidth=1)

treescrolly = tk.Scrollbar(frame_data, orient="vertical", command=tv1.yview)
treescrollx = tk.Scrollbar(frame_data, orient="horizontal", command=tv1.xview)
tv1.configure(xscrollcommand=treescrollx.set,
              yscrollcommand=treescrolly.set)
treescrollx.pack(side="bottom", fill="x")
treescrolly.pack(side="right", fill="y")


def File_dialog():
    filename = filedialog.askopenfilename(initialdir="/",
                                          title="Select A File",
                                          filetypes=(("xlsx files", "*.xlsx"), ("All Files", "*.*")))
    label_file["text"] = filename
    return None


def Load_excel_data():
    file_path = label_file["text"]
    try:
        excel_filename = r"{}".format(file_path)
        if excel_filename[-4:] == ".csv":
            df = pd.read_csv(excel_filename)
        else:
            df = pd.read_excel(excel_filename)

    except ValueError:
        tk.messagebox.showerror("Information", "The file you have chosen is invalid")
        return None
    except FileNotFoundError:
        tk.messagebox.showerror("Information", f"No such file as {file_path}")
        return None

    clear_data()
    tv1["column"] = list(df.columns)
    tv1["show"] = "headings"
    tv1["height"] = "6"
    for column in tv1["columns"]:
        tv1.heading(column, text=column)
    df_rows = df.to_numpy().tolist()
    for row in df_rows:
        tv1.insert("", "end", values=row)
    return None


def clear_data():
    tv1.delete(*tv1.get_children())
    return None

# ////////////////////////////////
window.mainloop()
