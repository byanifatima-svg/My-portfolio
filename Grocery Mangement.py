
import customtkinter as ctk
from customtkinter import *
import sqlite3
import pandas as pd
from tkinter import messagebox
from tkinter import ttk
import math
TOP =ctk.CTk()
TOP.title("Gracery Mangement System")
TOP.geometry("970x650")
TOP.resizable(width=False,height=False)
TOP_Main =ctk.CTkFrame(TOP,corner_radius=25)
TOP_Main.pack(expand =True,fill=BOTH,padx =5,pady =5)

conn = sqlite3.connect("grocery.db")
cursor = conn.cursor()

cursor.execute(""" CREATE TABLE IF NOT EXISTS grocery(
             id INTEGER PRIMARY KEY AUTOINCREMENT ,
             name VARCHAR (50),
             qantity INT,
             price INT
             )         
""" 
               )
conn.commit()

label_Item_Name =ctk.CTkLabel(TOP_Main,text="Item Name:",font=("Arail",15))
label_Item_Name.grid(row=0,column =0,padx =(55,5),pady=(30,5))
entry_Item_Name = ctk.CTkEntry(TOP_Main,width=350,height=30)
entry_Item_Name.grid(row=0,column =1,padx =5,pady=(30,5))

label_Qantity =ctk.CTkLabel(TOP_Main,text="Qantity:",font=("Arial",15))
label_Qantity.grid(row =1,column =0,padx=(55,5),pady=(25,5))
entry_Qantity =ctk.CTkEntry(TOP_Main,width=350,height=30)
entry_Qantity.grid(row =1,column =1,padx =5,pady =(25,5))

label_Price =ctk.CTkLabel(TOP_Main,text="Price:",font=("Arail",15))
label_Price.grid(row=2,column =0,padx =(55,5),pady=(25,5))
entry_Price =ctk.CTkEntry(TOP_Main,width=350,height=30)
entry_Price.grid(row =2,column =1,padx =5,pady =(25,5))

TOP_Main.grid_columnconfigure(0,weight=1)
TOP_Main.grid_columnconfigure(1,weight=1)
TOP_Main.grid_columnconfigure(2,weight=1)

tree = ttk.Treeview(TOP_Main,height=10,
columns=("Name","Quantity","Price"),show="headings")
tree.heading("Name",text=("Name"))
tree.heading("Quantity",text="Quantity")
tree.heading("Price",text="Price")

tree.column("Name",width=150)
tree.column("Quantity",width=100)
tree.column("Price",width=100)
tree.grid(row=6,rowspan=4,column=0,columnspan=4,sticky="nsew")

scrollbar = ctk.CTkScrollbar(TOP_Main,orientation="vertical")
scrollbar.grid(row =6,column =4,rowspan =4,sticky ="ns")
scrollbar.configure(command=tree.yview)
tree.configure(yscrollcommand=scrollbar.set)
TOP_Main.grid_rowconfigure(6,weight=1)
TOP_Main.grid_columnconfigure(0,weight=1)

def select_item(event):
    selected = tree.selection()
    if not selected:
        return
    values = tree.item(selected,"values")
    entry_Item_Name.delete(0,"end")
    entry_Qantity.delete(0,"end")
    entry_Price.delete(0,"end")

    entry_Item_Name.insert(0,values[0])
    entry_Qantity.insert(0,values[1])
    entry_Price.insert(0,values[2])
tree.bind("<<TreeviewSelect>>",select_item) 
def add_item():
    name =entry_Item_Name.get()
    qty =entry_Qantity.get()
    price =entry_Price.get()

    if name=="" or qty =="" or price == "":
        messagebox.showerror("Error","Please fill all fields!")
        return
    qty =int(qty)
    price =float(price)
  
    tree.insert("","end",values=(name,qty,price))
    entry_Item_Name.delete(0,"end")
    entry_Price.delete(0,"end")
    entry_Qantity.delete(0,"end")

   
btn_Add_Item = ctk.CTkButton(TOP_Main,
fg_color="#267D49",
text_color="white",
hover_color="#18D463",
corner_radius=8,border_color="#18D463",border_width=2,
text="Add Item",width=160,height=36,command=add_item)
btn_Add_Item.grid(row =3,column=0,padx =(65,5),pady =(40,5))

def delete_item():
    selected_item =tree.selection()
    if not selected_item:
        messagebox.showerror("Error","Pleace select an item")
        return
    tree.delete(selected_item)
  

    
btn_Delete = ctk.CTkButton(TOP_Main,
fg_color="#A82B2B",
text_color="white",
hover_color="#FF0000",
corner_radius=8,border_color="#FF0000",border_width=2,
text="Delete Item",width=160,height=36,command=delete_item
)
btn_Delete.grid(row =3,column=1,padx =(5,5),pady =(40,5))


def update_item():
    selected_item2 = tree.selection()
    if not selected_item2:
        messagebox.showerror("Error","Select an item")
        return
    name =entry_Item_Name.get()
    qantity=entry_Qantity.get()

    
    price =entry_Price.get()

    tree.item(selected_item2,values=(name,qantity,price))
    entry_Item_Name.delete(0,"end")
    entry_Price.delete(0,"end")
    entry_Qantity.delete(0,"end")

btn_Update_Item = ctk.CTkButton(TOP_Main,
fg_color="#249BCF",
text_color="white",
hover_color="#4DC4F8",
corner_radius=8,border_color="#4DC4F8",border_width=2,
text="Update Item",width=160,height=36,command=update_item)
btn_Update_Item.grid(row =3,column=2,padx =(5,5),pady =(40,5))

def show_item():
    try:
        df = pd.read_excel("grocery.xlsx")
        tree.delete(*tree.get_children())
        for index,row in df.iterrows():
            tree.insert("","end", values =(row["name"],row["qantity"],row["price"]))
            
    except FileNotFoundError:
        messagebox.showerror("Error","File not found!")

btn_Show_Item = ctk.CTkButton(TOP_Main,
fg_color="#C16910",
text_color="white",
hover_color="#CD9657",
corner_radius=8,border_color="#CD9657",border_width=2,
text="Show Item",width=160,height=36,command=show_item)
btn_Show_Item.grid(row =3,column=3,padx =(5,20),pady =(40,5))

label_Search = ctk.CTkLabel(TOP_Main,text="Search:",font=("Arial",15))
label_Search.grid(row =4,column =0,padx =(65,0),pady =(20,5))
entry_Search = ctk.CTkEntry(TOP_Main,width=350,height=30)
entry_Search.grid(row =4,column =1,padx = 0,pady=(20,5))

def search_item():
    search_text =entry_Search.get().lower()
    if search_text == "":
        messagebox.showerror("error","Please enter item name")  #scroll
        return
    try:
        df =pd.read_excel("grocery.xlsx")
        tree.delete(*tree.get_children())
        found =False

        for index,row in df.iterrows():
            if search_text in str(row["name"]).lower():
                tree.insert("","end",values=(row["name"],row["qantity"],row["price"]))
                found =True
        if not found:
                    messagebox.showinfo("Not found","item not found!")

    except Exception as e:
        messagebox.showerror("Error",str(e))

    
btn_Search =ctk.CTkButton(TOP_Main,text="Search",fg_color="#F4C55E",text_color="white",
hover_color="#D1C79A",
corner_radius=8,border_color="#D1C79A",border_width=2,
width=160,height=36,command=search_item)
btn_Search.grid(row =4,column =2,padx =5,pady=(20,5))

def save_btn():
    name_item = entry_Item_Name.get()
    qantity_item = entry_Qantity.get()
    price_item = entry_Price.get()

    cursor.execute("INSERT INTO grocery(name,qantity,price)VALUES(?,?,?)",
                (name_item,qantity_item ,price_item ))
    conn.commit()
    query = '''SELECT name,qantity,price FROM grocery
'''
    df = pd.read_sql_query(query,conn)
    df.to_excel("grocery.xlsx",index=False)

    if name_item=="" or qantity_item =="" or price_item =="":
        messagebox.showinfo("Info","Please fill all entries!")

    else:
        messagebox.showinfo("Info","Data successfully saved!")
    entry_Item_Name.delete(0,"end")
    entry_Price.delete(0,"end")
    entry_Qantity.delete(0,"end")

btn_Save =ctk.CTkButton(TOP_Main,text="Save",fg_color="#5EBE31",
text_color="white",
hover_color="#A5D781",
corner_radius=8,border_color="#A5D781",border_width=2,
width=160,height=36,command=save_btn)
btn_Save.grid(row =5,column =0,padx=(65,5),pady =10)
label_total = ctk.CTkLabel(TOP_Main,text="Total: 0")
label_total.grid(row =5,column =3)
def calculate_total():
    total =0
    for item in tree.get_children():
        values = tree.item(item,"values")
        try:
            qty = float(values[1])
            pri =float(values[2])
            if math.isnan(qty) or math.isnan(pri):
                continue

            total += qty * pri
        except:
            continue
    label_total.configure(text =f"Total:{total}")
btn_total =ctk.CTkButton(TOP_Main,text="Total Price",fg_color="#294BBC",
text_color="white",
hover_color="#4275B4",
corner_radius=8,border_color="#4275B4",border_width=2,
width=160,height=36,command=calculate_total)
btn_total.grid(row =5,column = 1,padx =5,pady =10)





TOP.mainloop()