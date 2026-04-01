import os
import re
import tkinter as tk
from tkinter import ttk, messagebox
import sqlite3
import sys


APP_NAME = "Asset Manager"
APP_VERSION = "v1.0.1"

#Colour theme 
BG_COLOR = "#f4f6f8"
PANEL_COLOR = "#e9ecef"
BUTTON_COLOR = "#008950"
BUTTON_TEXT = "#ffffff"
TEXT_COLOR = "#2c2c2c"
HEADER_COLOR = "#d6d9dc"
ROW_ALT = "#f2f2f2"
FONT_MAIN = ("Segoe UI", 10)

#Database location
BASE_DIR = r"\\Servername\path\Asset Manager\Database\Assets"
DB_NAME = os.path.join(BASE_DIR, "assets.db")

#Openning and returning connection to sqlite database
def connect_db():
    os.makedirs(BASE_DIR, exist_ok=True)
    return sqlite3.connect(DB_NAME, timeout=10, check_same_thread=False)

#Table and attributes 
def create_table():
    conn = connect_db()
    cur = conn.cursor()
    #asset tag is unique to prevent duplicated entries 
    cur.execute("""
        CREATE TABLE IF NOT EXISTS asset (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            asset_tag TEXT UNIQUE NOT NULL,
            category TEXT,
            make TEXT,
            model TEXT,
            serial_num TEXT,
            assigned_user TEXT,
            status TEXT,
            warranty TEXT,
            comments TEXT
        )
    """)

    conn.commit()
    conn.close()


"""
Fetches all asset records from the database.
Returns a list of tuples, each representing one asset row.
The 'id' column is included last so it can be hidden in the UI
but still accessed for update/delete operations.
"""
def fetch_assets():
    conn = connect_db()
    cur = conn.cursor()

    cur.execute("""
        SELECT asset_tag, category, make, model,
               serial_num, assigned_user, status,
               warranty, comments, id
        FROM asset
    """)

    rows = cur.fetchall()
    conn.close()
    return rows

#Searches across all fields rows that match the given keyword.
def search_assets(keyword):
    conn = connect_db()
    cur = conn.cursor()

    value = f"%{keyword}%" #SQL LIKE wildcard for matching keyword anywhere within a field value

    cur.execute("""
        SELECT asset_tag, category, make, model,
               serial_num, assigned_user, status,
               warranty, comments, id
        FROM asset
        WHERE asset_tag LIKE ? OR category LIKE ?
           OR make LIKE ? OR model LIKE ?
           OR serial_num LIKE ? OR assigned_user LIKE ?
           OR status LIKE ? OR warranty LIKE ?
    """, (value,) * 8) #Passing the searched keyword through all the placeholders 

    rows = cur.fetchall()
    conn.close()

    return rows

#Validating assigned user should be name.surename
def valid_assigned_user(name):
    return re.fullmatch(r"[A-Za-z]+\.[A-Za-z]+", name) is not None


"""
Reads values from all input fields and inserts a new asset record into the database.
Validates that Asset Tag is not empty and Assigned User follows name.surname format.
Shows an error dialog if the Asset Tag already exists (UNIQUE constraint violation).
Refreshes the table and clears input fields on success.
"""
def insert_asset():

    if not tag_var.get().strip(): 
        messagebox.showerror("Error", "Asset Tag is required")
        return

    if not valid_assigned_user(assigned_var.get().strip()):
        messagebox.showerror("Error", "Assigned User must be name.surname")
        return

    conn = connect_db()
    cur = conn.cursor()

    try:
        cur.execute("""
            INSERT INTO asset (
                asset_tag, category, make, model,
                serial_num, assigned_user, status,
                warranty, comments
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            tag_var.get().strip(),
            category_var.get(),
            make_var.get().strip(),
            model_var.get().strip(),
            serial_var.get().strip(),
            assigned_var.get().strip(),
            status_var.get(),
            warranty_var.get().strip(),
            comments_text.get("1.0", tk.END).strip()
        ))

        conn.commit()
        
    #Executes when constraint is violated 
    except sqlite3.IntegrityError:
        messagebox.showerror("Database Error", "Asset Tag already exists")

    finally:
        conn.close()

    refresh_table()
    clear_fields()

    """
    Updates the database record for the selected Treeview row.
    Reads the hidden 'id' column (last value in the row tuple) to identify the record.
    Shows an error if no row is selected before the button is clicked.
    """
def update_asset():

    selected = asset_table.focus() #Focus on the selected row in the Treeview

    if not selected:
        messagebox.showerror("Error", "Select an asset first")
        return

    asset_id = asset_table.item(selected, "values")[-1] #Selects and return the last tuple of the row (id)

    conn = connect_db()
    cur = conn.cursor()

    cur.execute("""
        UPDATE asset SET
            asset_tag=?, category=?, make=?, model=?,
            serial_num=?, assigned_user=?, status=?,
            warranty=?, comments=?
        WHERE id=?
    """, (
        tag_var.get().strip(),
        category_var.get(),
        make_var.get().strip(),
        model_var.get().strip(),
        serial_var.get().strip(),
        assigned_var.get().strip(),
        status_var.get(),
        warranty_var.get().strip(),
        comments_text.get("1.0", tk.END).strip(),
        asset_id
    ))

    conn.commit()
    conn.close()

    refresh_table()

"""
Deletes selected asset record
User confirmation before deletion
"""
def delete_asset():

    selected = asset_table.focus()

    if not selected:
        messagebox.showerror("Error", "Select an asset first")
        return

    asset_id = asset_table.item(selected, "values")[-1]

    if not messagebox.askyesno("Confirm", "Delete selected asset?"):
        return

    conn = connect_db()
    cur = conn.cursor()

    cur.execute("DELETE FROM asset WHERE id=?", (asset_id,))
    conn.commit()
    conn.close()

    refresh_table()
    clear_fields()



"""
Sorts all rows by the values in the specified column.
Updates column headings showing direction arrow
"""
def sort_column(tree, col, reverse):

    data = [(tree.set(k, col), k) for k in tree.get_children("")]#Enables to sort rows based on the column value

    #Numeric sort for columns containing numbers 
    try:
        data.sort(key=lambda t: float(t[0]), reverse=reverse)
    except:
        #Alphabetic fallback for text columns
        data.sort(reverse=reverse)

    #Reodering the rows in Treeview as per sorted order
    for index, (_, k) in enumerate(data):
        tree.move(k, "", index)

    #Resets the column headings and removing sort arrows.
    for c in columns[:-1]:
        tree.heading(c, text=c.replace("_", " ").title(),
                     command=lambda _c=c: sort_column(tree, _c, False))

    #Adding the arrow indicator.
    arrow = " ▲" if not reverse else " ▼"

    tree.heading(
        col,
        text=col.replace("_", " ").title() + arrow,
        command=lambda: sort_column(tree, col, not reverse)
    )



    """
    Clears all rows from the Treeview and reloads them fresh from the database.
    Adds background colors using even/odd tags for neater look.
    Updates asset count when added or removed.
    """
def refresh_table():

    asset_table.delete(*asset_table.get_children())

    rows = fetch_assets()

    for i, row in enumerate(rows):

        if i % 2 == 0:
            asset_table.insert("", tk.END, values=row, tags=("even",))
        else:
            asset_table.insert("", tk.END, values=row, tags=("odd",))

    count_var.set(f"Total Assets: {len(rows)}")


def clear_fields():
    #Resets all input fields 
    tag_var.set("")
    category_var.set("")
    make_var.set("")
    model_var.set("")
    serial_var.set("")
    assigned_var.set("")
    status_var.set("")
    warranty_var.set("")

    comments_text.delete("1.0", tk.END)


#Loads input field with selected data from Treeview, useful for updating.
def load_selected(event):

    selected = asset_table.focus()

    if not selected:
        return

    data = asset_table.item(selected, "values")

    tag_var.set(data[0])
    category_var.set(data[1])
    make_var.set(data[2])
    model_var.set(data[3])
    serial_var.set(data[4])
    assigned_var.set(data[5])
    status_var.set(data[6])
    warranty_var.set(data[7])

    comments_text.delete("1.0", tk.END)
    comments_text.insert(tk.END, data[8])


def search_action():

    keyword = search_var.get().strip()

    asset_table.delete(*asset_table.get_children()) #Clear existing rows before inserting updated data into the Treeview.

    rows = search_assets(keyword) if keyword else fetch_assets()#Using filtered search if keyword exists, otherwise show all records

    for row in rows:
        asset_table.insert("", tk.END, values=row)

    count_var.set(f"Total Assets: {len(rows)}")

#Icon path
if getattr(sys, 'frozen', False):
    ICON_PATH = os.path.join(sys._MEIPASS, "asset.ico")
else:
    ICON_PATH = os.path.join(os.path.dirname(__file__), "asset.ico")
#Main window setup
root = tk.Tk()
root.title(f"{APP_NAME} {APP_VERSION}")
root.geometry("1200x560")
root.configure(bg=BG_COLOR)

#Setting icon
try:
    root.iconbitmap(ICON_PATH)
except Exception as e:
    print("Icon failed to load:", e)


#Treeview Styling 
style = ttk.Style()
style.theme_use("default")

style.configure(
    "Treeview",
    background="white",
    foreground=TEXT_COLOR,
    rowheight=25,
    fieldbackground="white",
    font=FONT_MAIN
)

style.map("Treeview", background=[("selected", BUTTON_COLOR)])

style.configure(
    "Treeview.Heading",
    background=HEADER_COLOR,
    foreground=TEXT_COLOR,
    font=("Segoe UI", 10, "bold")
)

"""
Tkninter varibales.
StringVar instances providing two way binding betweens widgets and python values
"""
tag_var = tk.StringVar()
category_var = tk.StringVar()
make_var = tk.StringVar()
model_var = tk.StringVar()
serial_var = tk.StringVar()
assigned_var = tk.StringVar()
status_var = tk.StringVar()
warranty_var = tk.StringVar()
search_var = tk.StringVar()
count_var = tk.StringVar()


left = tk.Frame(root, padx=10, pady=10, bg=PANEL_COLOR)
left.pack(side=tk.LEFT, fill=tk.Y)


def label(parent, text):
    tk.Label(parent, text=text, bg=PANEL_COLOR,
             fg=TEXT_COLOR, font=FONT_MAIN).pack(anchor="w")


label(left, "Asset Tag")
tk.Entry(left, textvariable=tag_var, width=30).pack()

label(left, "Category")
ttk.Combobox(left, textvariable=category_var,
             values=["Laptop", "Desktop", "All-in-One"], 
             state="readonly", width=28).pack()

label(left, "Make")
ttk.Combobox(left, textvariable=make_var,
             values=["Lenovo ThinkPad", "Lenovo ThinkBook",
                     "Dell Latitude", "Asus", "Lenovo", "HP"],
             state="readonly", width=28).pack()

label(left, "Model")
tk.Entry(left, textvariable=model_var, width=30).pack()

label(left, "Serial Number")
tk.Entry(left, textvariable=serial_var, width=30).pack()

label(left, "Assigned User")
tk.Entry(left, textvariable=assigned_var, width=30).pack()

label(left, "Status")
ttk.Combobox(left, textvariable=status_var,
             values=["New", "In Use", "Needs Repair",
                     "Unrepairable", "Disposed",
                     "Out for Repair", "New & Not distributed",
                     "Not distributed"],
             state="readonly", width=28).pack()

label(left, "Warranty")
tk.Entry(left, textvariable=warranty_var, width=30).pack()

label(left, "Comments")
comments_text = tk.Text(left, width=30, height=5)
comments_text.pack()

#Action buttons which triggers its corresponding database
for text, cmd in [
    ("Add Asset", insert_asset),
    ("Update Asset", update_asset),
    ("Delete Asset", delete_asset),
    ("Clear Fields", clear_fields)
]:
    tk.Button(left, text=text, command=cmd,
              bg=BUTTON_COLOR, fg=BUTTON_TEXT,
              activebackground="#2c4763",
              relief="flat",
              font=("Segoe UI", 10, "bold")).pack(fill=tk.X, pady=4)



right = tk.Frame(root, padx=10, pady=10, bg=BG_COLOR)
right.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

search_frame = tk.Frame(right, bg=BG_COLOR)
search_frame.pack(fill=tk.X)

tk.Label(search_frame, text="Search",
         bg=BG_COLOR, fg=TEXT_COLOR,
         font=FONT_MAIN).pack(side=tk.LEFT)

tk.Entry(search_frame, textvariable=search_var,
         width=30).pack(side=tk.LEFT, padx=5)

tk.Button(search_frame, text="Search",
          command=search_action).pack(side=tk.LEFT)

#Relaoding all records.
tk.Button(search_frame, text="Refresh",
          command=refresh_table).pack(side=tk.LEFT, padx=5)

tk.Label(right, textvariable=count_var,
         font=("Segoe UI", 10, "bold"),
         bg=BG_COLOR).pack(anchor="e")

#Treeview table 
table_frame = tk.Frame(right)
table_frame.pack(fill=tk.BOTH, expand=True)

columns = ("asset_tag", "category", "make", "model",
           "serial_num", "assigned_user",
           "status", "warranty", "comments", "id")


scroll_y = ttk.Scrollbar(table_frame, orient="vertical")
scroll_x = ttk.Scrollbar(table_frame, orient="horizontal")

asset_table = ttk.Treeview(
    table_frame,
    columns=columns,
    show="headings", #hides the default empty column
    yscrollcommand=scroll_y.set,
    xscrollcommand=scroll_x.set
)

#Linking scrollbars to the Treeview scroll methods
scroll_y.config(command=asset_table.yview)
scroll_x.config(command=asset_table.xview)


scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
scroll_x.pack(side=tk.BOTTOM, fill=tk.X)
asset_table.pack(fill=tk.BOTH, expand=True)

for col in columns[:-1]:#Skipping 'id' it has no visible heading
    asset_table.heading(
        col,
        text=col.replace("_", " ").title(),
        command=lambda c=col: sort_column(asset_table, c, False)
    )
    asset_table.column(col, width=140, anchor="w")

#Hides the 'id' column — width=0 and stretch=False makes it invisible
asset_table.column("id", width=0, stretch=False)

#Populates input fields automatically whenever a row is selected
asset_table.bind("<<TreeviewSelect>>", load_selected)

#Configuring row tags for background 
asset_table.tag_configure("even", background="white")
asset_table.tag_configure("odd", background=ROW_ALT)

create_table()
refresh_table()

root.mainloop()