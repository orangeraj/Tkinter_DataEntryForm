import tkinter
from tkinter import ttk
from tkinter import messagebox
import openpyxl
import os
from openpyxl.utils.exceptions import InvalidFileException

def get_info():

    custname = cust_name_entry.get()
    deliverytime = deliverytime_entry.get()
    #price = price_entry.get()
    #price = float(price)
    #price_f = "{:.2f}".format(price)
    ordertype = ordertype_combobox.get()
    
    deliveryperson = deliveryperson_combobox.get()
    paidstatus = paid_status_var.get()

    rice_quant = rice_spinbox.get()
    rice_quant = int(rice_quant)

    bhaji_quant = bhaji_spinbox.get()
    bhaji_quant = int(bhaji_quant)

    bhakri_quant = bhakari_spinbox.get()
    bhakri_quant = int(bhakri_quant)

    varan_quant = varan_spinbox.get()
    varan_quant = int(varan_quant)

    bhaji_type = bhaji_entry.get()
    bhakri_type = bhakri_entry.get()
    
    order_details = ("Rice: "+ str(rice_quant) + " | \n" + 
                    "Bhaji: "+ str(bhaji_quant) + " " + str(bhaji_type) + " " + " | \n" + 
                    "Bhakari: "+ str(bhakri_quant) + " " + str(bhakri_type) + " " + " | \n" +
                    "Varan: "+ str(varan_quant) + " | \n" )
    
    #calculate total price
    Rice_PerPlate = 10
    Bhaji_PerPlate = 15
    Varan_PerPlate = 20
    Bhakri_PerPlate = 15

    calc_price = rice_quant * float(Rice_PerPlate) + bhaji_quant * float(Bhaji_PerPlate) + bhakri_quant * float(Bhakri_PerPlate) + varan_quant * float(Varan_PerPlate)  
    
    print(calc_price)
    print(order_details)
    
    if custname:
        if ordertype:
            filepath = "C:\\Users\\mhatr\\OneDrive\\Tkinter\\data.xlsx"      
            if not os.path.exists(filepath):
                workbook = openpyxl.Workbook()
                sheet = workbook.active
                heading = ["Customer Name", "Delivery Time", "Price", "Order Type", 
                            "Payment Status", "Delivery Person", "Order Details"]
                sheet.append(heading)
                workbook.save(filepath)
               
            try: 
                workbook = openpyxl.load_workbook(filepath)
                sheet = workbook.active
                sheet.append([custname, deliverytime, calc_price, ordertype, paidstatus, deliveryperson, order_details])
                workbook.save(filepath)

            except:
                #print("Please Close Excel File !")
                tkinter.messagebox.showwarning(title="ERROR !", message="Please Close Excel File")

        else: 
            tkinter.messagebox.showwarning(title="ERROR !", message="Order Type missing")
    else:
        tkinter.messagebox.showwarning(title="ERROR !", message="Customer Name missing")

#function ends-----------

def get_calc():
    pass


#function ends-----------

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
deliverytime_label = tkinter.Label(order_detail_frame, text="Delivery Time")
deliverytime_label.grid(row=0, column=1)
deliverytime_entry = tkinter.Entry(order_detail_frame)
deliverytime_entry.grid(row=1,column=1)

#save total order price
price_label = tkinter.Label(order_detail_frame, text="Total Price")
price_label.grid(row=0, column=2)
calc_price = tkinter.StringVar(value = 0.0)
price_entry = tkinter.Label(order_detail_frame, textvariable=calc_price)
price_entry.grid(row=1,column=2)

#check if delivery or takeaway
ordertype_label = tkinter.Label(order_detail_frame, text="Order Type")
ordertype_label.grid(row=2,column=0)
ordertype_combobox = ttk.Combobox(order_detail_frame, values=["Take Away","Delivery"])
ordertype_combobox.grid(row=3,column=0)

#calculate button
calc_button = tkinter.Button(order_detail_frame, text="Calculate Price", command= get_calc)
calc_button.grid(row=3, column=1, padx=20, pady=20)

for widget in order_detail_frame.winfo_children():
    widget.grid_configure(padx=10, pady=5)



#frame 02: delivery details
delivey_detail_frame = tkinter.LabelFrame(frame, text="Delivery Details")
delivey_detail_frame.grid(row=1, column=0, sticky="news", padx=20, pady=20)

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
menu_detail_frame = tkinter.LabelFrame(frame, text="Menu")
menu_detail_frame.grid(row=2, column=0, sticky="news", padx=20, pady=10)

#menu items grid
quantity_label = tkinter.Label(menu_detail_frame, text="Quantity")
quantity_label.grid(row=0, column=1)

rice_label = tkinter.Label(menu_detail_frame, text="Rice")
rice_label.grid(row=1, column=0)
rice_spinbox = tkinter.Spinbox(menu_detail_frame, from_=0, to=99)
rice_spinbox.grid(row=1,column=1)

bhaji_label = tkinter.Label(menu_detail_frame, text="Bhaji")
bhaji_label.grid(row=2, column=0)
bhaji_spinbox = tkinter.Spinbox(menu_detail_frame, from_=0, to=99)
bhaji_spinbox.grid(row=2,column=1)

bhakari_label = tkinter.Label(menu_detail_frame, text="Bhakari")
bhakari_label.grid(row=3, column=0)
bhakari_spinbox = tkinter.Spinbox(menu_detail_frame, from_=0, to=99)
bhakari_spinbox.grid(row=3,column=1)

varan_label = tkinter.Label(menu_detail_frame, text="Varan")
varan_label.grid(row=4, column=0)
varan_spinbox = tkinter.Spinbox(menu_detail_frame, from_=0, to=99)
varan_spinbox.grid(row=4,column=1)

#types of menu
type_label = tkinter.Label(menu_detail_frame, text="Type")
type_label.grid(row=0, column=2)

bhaji_entry = tkinter.Entry(menu_detail_frame)
bhaji_entry.grid(row=2,column=2)

bhakri_entry = tkinter.Entry(menu_detail_frame)
bhakri_entry.grid(row=3,column=2)


for widget in menu_detail_frame.winfo_children():
    widget.grid_configure(padx=10, pady=5)


#submit button
button = tkinter.Button(frame, text="Submit", command= get_info)
button.grid(row=3, column=0, sticky="news", padx=20, pady=20)


window.mainloop()

