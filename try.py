import tkinter as tk
from tkinter import ttk


def adjust_treeview_columns(tree, fullscreen_mode):
    # Define widths for fullscreen and normal modes
    fullscreen_widths = {
        "Artikel": 150,
        "Spannung": 100,
        "Stk": 80,
        "Rohling": 120,
        "Status": 100,
        "Wkz": 100,
        "Bemerkung": 150,
        "Auftragsstatus": 120
    }

    normal_widths = {
        "Artikel": 120,
        "Spannung": 80,
        "Stk": 60,
        "Rohling": 100,
        "Status": 80,
        "Wkz": 80,
        "Bemerkung": 120,
        "Auftragsstatus": 0
    }

    # Choose widths based on the window state
    widths = fullscreen_widths if fullscreen_mode else normal_widths

    # Set column widths
    for col in tree["columns"]:
        tree.column(col, width=widths.get(col, 100), anchor='w')


def toggle_fullscreen():
    global fullscreen
    fullscreen = not fullscreen
    root.attributes('-fullscreen', fullscreen)
    adjust_treeview_columns(tree, fullscreen)


root = tk.Tk()
root.title("Treeview Example")
root.geometry('1366x768')  # Initial size

fullscreen = False  # Track fullscreen state

# Create Treeview
tree = ttk.Treeview(root, style="Custom.Treeview", show='headings', selectmode="browse")
tree["columns"] = ("Artikel", "Spannung", "Stk", "Rohling", "Status", "Wkz", "Bemerkung", "Auftragsstatus")
for col in tree["columns"]:
    tree.heading(col, text=col)

# Set initial column widths
adjust_treeview_columns(tree, fullscreen)

tree.grid(row=0, column=0, sticky="nsew", padx=(10, 0), pady=(100, 10))

# Toggle fullscreen mode with a button
toggle_button = tk.Button(root, text="Toggle Fullscreen", command=toggle_fullscreen)
toggle_button.grid(row=1, column=0, pady=10)

root.mainloop()
