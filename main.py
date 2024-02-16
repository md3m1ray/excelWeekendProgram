import tkinter as tk
from tkinter import ttk
import openpyxl


def load_data():
    path = "people.xlsx"
    workbook = openpyxl.load_workbook(path)
    sheet = workbook.active

    list_values = list(sheet.values)

    for col_name in list_values[0]:
        treeview.heading(col_name, text=col_name)

    for value_tuple in list_values[1:]:
        treeview.insert('', tk.END, values=value_tuple)


def delete_row():
    path = "people.xlsx"
    workbook = openpyxl.load_workbook(path)
    sheet = workbook.active
    selected_item2 = treeview.focus()

    for row in sheet.iter_rows(min_row=2, min_col=1, max_row=100, max_col=1, values_only=False):
        for cell in row:
            if cell.value == f"{selected_item2}":
                sheet.delete_rows(cell.row)
                treeview.delete(selected_item2)

    workbook.save(path)


def insert_row():
    id_ = ""
    name = name_entry.get()
    day = status_combobox.get()

    path = "people.xlsx"
    workbook = openpyxl.load_workbook(path)
    sheet = workbook.active
    row_values = [id_, name, day]
    sheet.append(row_values)
    workbook.save(path)

    treeview.insert('', tk.END, values=row_values)

    name_entry.delete(0, "end")
    name_entry.insert(0, "İsim Soyisim")


# def on_double_click(event):
#     print(treeview.set(treeview.identify_row(event.y)))

def get_id(event):
    focused = treeview.focus()
    path = "people.xlsx"
    workbook = openpyxl.load_workbook(path)
    sheet = workbook.active

    cur_item = treeview.item(treeview.focus())
    for row in sheet.iter_rows(min_row=2, min_col=1, max_row=100, max_col=2, values_only=False):
        for cell in row:
            if cell.value == f"{cur_item['values'][1]}":
                sheet.cell(row=cell.row, column=1).value = f"{focused}"

    treeview.set(focused, column=0, value=focused)

    workbook.save(path)


def add_day():
    focused = treeview.focus()
    day = status_combobox.get()

    path = "people.xlsx"
    workbook = openpyxl.load_workbook(path)
    sheet = workbook.active

    for row in sheet.iter_rows(min_row=2, min_col=1, max_row=100, max_col=1, values_only=False):
        for cell in row:
            if cell.value == f"{focused}":
                sheet.cell(row=cell.row, column=3).value = day

    treeview.set(focused, column=2, value=day)

    workbook.save(path)


def select_all():
    for item in treeview.get_children():
        select_children(item)
        get_all()


def select_children(item):
    treeview.selection_add(item)


def get_all():
    path = "people.xlsx"
    workbook = openpyxl.load_workbook(path)
    sheet = workbook.active

    for item in treeview.get_children():
        for i in item:
            ids = "I00" + str(i)
            r = 2
            treeview.set(item, column=0, value=ids)
            sheet.cell(row=r, column=1).value = ids
            r += 1

    workbook.save(path)


def toggle_mode():
    if mode_switch.instate(["selected"]):
        style.theme_use("forest-light")
    else:
        style.theme_use("forest-dark")


root = tk.Tk()

style = ttk.Style(root)
root.tk.call("source", "forest-light.tcl")
root.tk.call("source", "forest-dark.tcl")
style.theme_use("forest-dark")
root.title("Hafta Sonu Çalışma Çizelgesi")
combo_list = ["Cumartesi", "Pazar"]

frame = ttk.Frame(root)
frame.pack()

widgets_frame = ttk.LabelFrame(frame, text="Personel Ekle/Sil")
widgets_frame.grid(row=0, column=0, padx=20, pady=10)

name_entry = ttk.Entry(widgets_frame)
name_entry.insert(0, "İsim Soyisim")
name_entry.bind("<FocusIn>", lambda e: name_entry.delete('0', 'end'))
name_entry.grid(row=2, column=0, padx=5, pady=(0, 5), sticky="ew")

button = ttk.Button(widgets_frame, text="Personel Ekle", command=insert_row)
button.grid(row=3, column=0, padx=5, pady=5, sticky="nsew")

separator = ttk.Separator(widgets_frame)
separator.grid(row=4, column=0, padx=(20, 10), pady=10, sticky="ew")

button = ttk.Button(widgets_frame, text="Seç ve Sil", command=delete_row)
button.grid(row=6, column=0, padx=5, pady=5, sticky="nsew")

separator = ttk.Separator(widgets_frame)
separator.grid(row=7, column=0, padx=(20, 10), pady=10, sticky="ew")

button = ttk.Button(widgets_frame, text="ID Güncelle", command=select_all)
button.grid(row=8, column=0, padx=5, pady=5, sticky="nsew")

separator = ttk.Separator(widgets_frame)
separator.grid(row=9, column=0, padx=(20, 10), pady=10, sticky="ew")

status_combobox = ttk.Combobox(widgets_frame, values=combo_list)
status_combobox.current(0)
status_combobox.grid(row=10, column=0, sticky="ew")

button = ttk.Button(widgets_frame, text="Seç ve Değiştir", command=add_day)
button.grid(row=11, column=0, padx=5, pady=5, sticky="nsew")

separator = ttk.Separator(widgets_frame)
separator.grid(row=12, column=0, padx=(20, 10), pady=10, sticky="ew")

mode_switch = ttk.Checkbutton(
    widgets_frame, text="Tema", style="Switch", command=toggle_mode)
mode_switch.grid(row=13, column=0, padx=5, pady=10, sticky="nsew")

treeFrame = ttk.Frame(frame)
treeFrame.grid(row=0, column=1, pady=10)
treeScroll = ttk.Scrollbar(treeFrame)
treeScroll.pack(side="right", fill="y")

cols = ("ID", "İsim Soyisim", "Gün")
treeview = ttk.Treeview(treeFrame, show="headings",
                        yscrollcommand=treeScroll.set, columns=cols, height=35)

treeview.column("ID", width=60)
treeview.column("İsim Soyisim", width=200)
treeview.column("Gün", width=70)

treeview.bind("<Double-1>", get_id)
# treeview.bind("<Return>", lambda e: get_id)

treeview.pack()
treeScroll.config(command=treeview.yview)

load_data()

root.mainloop()
