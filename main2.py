import openpyxl
import csv
import tkinter as tk
from tkinter import ttk

# Load the invoice template Excel file
invoice_template = openpyxl.load_workbook('invoice_template.xlsx')

# Select the sheet where you want to insert data
worksheet = invoice_template['Sheet1']

# Define the product list
with open('product_list.csv') as csvfile:
    reader = csv.DictReader(csvfile)
    product_list = [row for row in reader]

output_list = []
for product in product_list:
    output_dict = {
        "Product Name:": product['name'],
        "Product Number:": product['number'],
        "Product HS-Code:": product['hs'],
        "Product Packing:": product['packing']
    }
    output_list.append(output_dict)
row = 23
serial_number = 1

selected_products = []

# GUI code
def add_to_invoice():
    global selected_products
    choice = int(products_combobox.get().split(".")[0]) - 1
    selected_product = product_list[choice]
    quantity = int(quantity_entry.get())
    price = int(price_entry.get())
    selected_product['price'] = price
    selected_product['quantity'] = quantity
    selected_product['total_price'] = quantity * price
    selected_products.append(selected_product)
    selected_products_listbox.insert(tk.END, f'{selected_product["name"]} || {selected_product["packing"]} (x{quantity})')
    quantity_entry.delete(0, tk.END)
    price_entry.delete(0, tk.END)
    update_total()

def remove_from_invoice():
    global selected_products
    selected = selected_products_listbox.curselection()
    if len(selected) > 0:
        selected_index = selected[0]
        selected_products.pop(selected_index)
        selected_products_listbox.delete(selected_index)
        update_total()

def update_total():
    global selected_products
    total_sum = sum(product['total_price'] for product in selected_products)
    total_label.config(text=f'Total: ${total_sum}')

def generate_invoice():
    global selected_products
    invoice_number = invoice_entry.get()
    invoice_date = date_entry.get()
    lpo_no = lpo_number.get()
    lpo_d = lpo_date.get()
    company_n = company_name.get()
    customer_n = customer_name.get()
    cr_no = cr_number.get()
    vatreg_no = vatreg_number.get()
    saleman = salesman.get()
    contact_per = contact_person.get()
    desig = designation.get()
    contact_no = contact_number.get()
    omanID = oman.get()
    worksheet['B13'] = invoice_number
    worksheet['J11'] = invoice_number
    worksheet['B14'] = invoice_date
    worksheet['B15'] = lpo_no
    worksheet['E15'] = lpo_d
    worksheet['B17'] = company_n
    worksheet['B18'] = customer_n
    worksheet['B19'] = cr_no
    worksheet['B20'] = vatreg_no
    worksheet['J14'] = saleman
    worksheet['J17'] = contact_per
    worksheet['J18'] = desig
    worksheet['J19'] = contact_no
    worksheet['J20'] = omanID

    for i, product in enumerate(selected_products):
        serial_number = 1
        row = i + 23
        vat = product['total_price'] * 0.05
        sum_price = int(product['total_price'] + vat)

        worksheet[f'A{row}'] = serial_number
        worksheet[f'B{row}'] = product['number']
        worksheet[f'C{row}'] = product['name']
        worksheet[f'D{row}'] = product['hs']
        worksheet[f'E{row}'] = product['packing']
        worksheet[f'F{row}'] = product['quantity']
        worksheet[f'G{row}'] = product['price']
        worksheet[f'H{row}'] = product['total_price']
        worksheet[f'I{row}'] = vat
        worksheet[f'J{row}'] = sum_price
        worksheet[f'I42{row}'] = product['total_price']
        worksheet[f'J42{row}'] = vat

        serial_number += 1

        # add the sum_price to the selected product dictionary
        product['sum_price'] = sum_price

    total_sum = sum(product['sum_price'] for product in selected_products)
    worksheet['J36'] = total_sum

    invoice_template.save(f'invoice-{invoice_number}.xlsx')
    root.destroy()

root = tk.Tk()
root.title("Invoice Generator")


canvas = tk.Canvas(root)
canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
scrollbar = tk.Scrollbar(root, orient=tk.VERTICAL, command=canvas.yview)
scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
canvas.configure(yscrollcommand=scrollbar.set)

# Create a frame inside the canvas for the widgets
frame = tk.Frame(canvas)
canvas.create_window((0, 0), window=frame, anchor=tk.NW)

# Product combobox
products_label = tk.Label(frame, text="Products:")
products_label.grid(row=1, column=0, padx=10, pady=10, sticky="W")
products_combobox = ttk.Combobox(frame, state="readonly", width=80, values=[f"{i+1}. {product['name']} || {product['packing']}" for i, product in enumerate(product_list)])
products_combobox.grid(row=1, column=1, padx=10, pady=10)


quantity_label = tk.Label(frame, text="Quantity:")
quantity_label.grid(row=2, column=0, padx=10, pady=10, sticky="W")

quantity_entry = tk.Entry(frame)
quantity_entry.grid(row=2, column=1, padx=10, pady=10)

price_label = tk.Label(frame, text="Price:")
price_label.grid(row=3, column=0, padx=10, pady=10, sticky="W")

price_entry = tk.Entry(frame)
price_entry.grid(row=3, column=1, padx=10, pady=10)

invoice_label = tk.Label(frame, text="Invoice Number:")
invoice_label.grid(row=8, column=0, padx=10, pady=10, sticky="W")

invoice_entry = tk.Entry(frame)
invoice_entry.grid(row=8, column=1, padx=10, pady=10)

date_label = tk.Label(frame, text="Invoice Date:")
date_label.grid(row=9, column=0, padx=10, pady=10, sticky="W")

date_entry = tk.Entry(frame)
date_entry.grid(row=9, column=1, padx=10, pady=10)

lpo_label = tk.Label(frame, text="LPO Number:")
lpo_label.grid(row=10, column=0, padx=10, pady=10, sticky="W")

lpo_number = tk.Entry(frame)
lpo_number.grid(row=10, column=1, padx=10, pady=10)


lpo_date = tk.Label(frame, text="LPO Date:")
lpo_date.grid(row=11, column=0, padx=10, pady=10, sticky="W")

lpo_date = tk.Entry(frame)
lpo_date.grid(row=11, column=1, padx=10, pady=10)


company_name = tk.Label(frame, text="Company Name:")
company_name.grid(row=12, column=0, padx=10, pady=10, sticky="W")

company_name = tk.Entry(frame)
company_name.grid(row=12, column=1, padx=10, pady=10)


customer_name = tk.Label(frame, text="Customer Name:")
customer_name.grid(row=13, column=0, padx=10, pady=10, sticky="W")

customer_name = tk.Entry(frame)
customer_name.grid(row=13, column=1, padx=10, pady=10)


cr_number = tk.Label(frame, text="Credit Number:")
cr_number.grid(row=14, column=0, padx=10, pady=10, sticky="W")

cr_number = tk.Entry(frame)
cr_number.grid(row=14, column=1, padx=10, pady=10)


vatreg_number = tk.Label(frame, text="Vat Registration Number:")
vatreg_number.grid(row=15, column=0, padx=10, pady=10, sticky="W")

vatreg_number = tk.Entry(frame)
vatreg_number.grid(row=15, column=1, padx=10, pady=10)


salesman = tk.Label(frame, text="Salesman:")
salesman.grid(row=16, column=0, padx=10, pady=10, sticky="W")

salesman = tk.Entry(frame)
salesman.grid(row=16, column=1, padx=10, pady=10)


contact_person = tk.Label(frame, text="Contact Person:")
contact_person.grid(row=17, column=0, padx=10, pady=10, sticky="W")

contact_person= tk.Entry(frame)
contact_person.grid(row=17, column=1, padx=10, pady=10)


designation = tk.Label(frame, text="Designation:")
designation.grid(row=18, column=0, padx=10, pady=10, sticky="W")

designation = tk.Entry(frame)
designation.grid(row=18, column=1, padx=10, pady=10)


contact_number = tk.Label(frame, text="Contact Number:")
contact_number.grid(row=19, column=0, padx=10, pady=10, sticky="W")

contact_number = tk.Entry(frame)
contact_number.grid(row=19, column=1, padx=10, pady=10)


oman = tk.Label(frame, text="Oman ID:")
oman.grid(row=20, column=0, padx=10, pady=10, sticky="W")

oman = tk.Entry(frame)
oman.grid(row=20, column=1, padx=10, pady=10)

add_button = tk.Button(frame, text="Add to Invoice", command=add_to_invoice)
add_button.grid(row=4, column=0, padx=10, pady=10)

remove_button = tk.Button(frame, text="Remove from Invoice", command=remove_from_invoice)
remove_button.grid(row=7, column=1, padx=10, pady=10)

selected_products_listbox = tk.Listbox(frame, width=100)
selected_products_listbox.grid(row=5, column=0, columnspan=2, padx=10, pady=10)

total_label = tk.Label(frame, text="Total: $0")
total_label.grid(row=6, column=1, padx=10, pady=10, sticky="E")

generate_button = tk.Button(frame, text="Generate Invoice", command=generate_invoice)
generate_button.grid(row=21, column=0, columnspan=2, padx=10, pady=10)


frame.update_idletasks()
canvas.configure(scrollregion=canvas.bbox(tk.ALL))

root.mainloop()
