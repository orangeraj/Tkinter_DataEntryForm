import tkinter
from tkinter import ttk
from tkinter import messagebox
import openpyxl
import os

def get_info():

    custname = cust_name_entry.get()
    itemdetails = item_details_entry.get()
    price = price_entry.get()
    ordertype = ordertype_combobox.get()
    deliveryperson = deliveryperson_combobox.get()
    paidstatus = paid_status_var.get()
    
    if custname:
        if itemdetails:
            if price:
                if ordertype:

                    filepath = "C:\\Users\\mhatr\\OneDrive\\Tkinter\\data.xlsx"
                    
                    if not os.path.exists(filepath):
                        workbook = openpyxl.Workbook()
                        sheet = workbook.active
                        heading = ["Customer Name", "Item Details", "Price", "Order Type", 
                                   "Payment Status", "Delivery Person"]
                        sheet.append(heading)
                        workbook.save(filepath)
                    
                    workbook = openpyxl.load_workbook(filepath)
                    sheet = workbook.active
                    sheet.append([custname, itemdetails, price, ordertype, paidstatus, deliveryperson])

                    workbook.save(filepath)

                else: 
                    tkinter.messagebox.showwarning(title="ERROR !", message="Order Type missing")
            else:
                tkinter.messagebox.showwarning(title="ERROR !", message="Price missing")
        else:
            tkinter.messagebox.showwarning(title="ERROR !", message=" Item Details missing")
    else:
        tkinter.messagebox.showwarning(title="ERROR !", message="Customer Name missing")

window = tkinter.Tk()
window.title("Order Entry Form")
frame = tkinter.Frame(window)
frame.pack()


#frame 01: order details
order_detail_frame = tkinter.LabelFrame(frame, text="Order Details")
order_detail_frame.grid(row=0, column=0, padx=20, pady=10)

#save customer details
cust_name_label = tkinter.Label(order_detail_frame, text='Customer Name')
cust_name_label.grid(row=0,column=0)
cust_name_entry = tkinter.Entry(order_detail_frame)
cust_name_entry.grid(row=1,column=0)

#save item details
item_details_label = tkinter.Label(order_detail_frame, text="Item Details")
item_details_label.grid(row=0, column=1)
item_details_entry = tkinter.Entry(order_detail_frame)
item_details_entry.grid(row=1,column=1)

#save total order price
price_label = tkinter.Label(order_detail_frame, text="Price")
price_label.grid(row=0, column=2)
price_entry = tkinter.Entry(order_detail_frame)
price_entry.grid(row=1,column=2)

#check if delivery or takeaway
ordertype_label = tkinter.Label(order_detail_frame, text="Order Type")
ordertype_label.grid(row=2,column=0)
ordertype_combobox = ttk.Combobox(order_detail_frame, values=["Take Away","Delivery"])
ordertype_combobox.grid(row=3,column=0)

for widget in order_detail_frame.winfo_children():
    widget.grid_configure(padx=10, pady=5)



#frame 02: delivery details
delivey_detail_frame = tkinter.LabelFrame(frame, text="Delivery Details")
delivey_detail_frame.grid(row=1, column=0, sticky="news", padx=20, pady=10)

#select delivery person
deliveryperson_label = tkinter.Label(delivey_detail_frame, text="Delivery Person")
deliveryperson_label.grid(row=0,column=0)
deliveryperson_combobox = ttk.Combobox(delivey_detail_frame, values=["Lomesh","Rajas"])
deliveryperson_combobox.grid(row=1,column=0)

#check if amount is paid or not
amountpaid_label = tkinter.Label(delivey_detail_frame, text="Payment Status")
amountpaid_label.grid(row=0, column=1)

paid_status_var = tkinter.StringVar(value="Not Paid")
amountpaid_check = tkinter.Checkbutton(delivey_detail_frame, text="Paid", 
                                       variable=paid_status_var, onvalue="Paid", offvalue="Not Paid")
amountpaid_check.grid(row=1, column=1)


for widget in delivey_detail_frame.winfo_children():
    widget.grid_configure(padx=10, pady=5)



#frame 03: Submit it
submit_detail_frame = tkinter.LabelFrame(frame, text="Submit")
submit_detail_frame.grid(row=2, column=0, sticky="news", padx=20, pady=10)

button = tkinter.Button(frame, text="Submit", command= get_info)
button.grid(row=3, column=0, sticky="news", padx=20, pady=20)


window.mainloop()