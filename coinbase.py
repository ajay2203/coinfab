import pandas as pd
import os
import numpy as np
import math


final_column= ["Order ID", "AWB Number" , "Total weight as per X (KG)", "Weight slab as per X (KG)", "Total weight as per Courier Company (KG)", "Weight slab charged by Courier Company (KG)", "Delivery Zone as per X", "Delivery Zone charged by Courier Company","Expected Charge as per X (Rs.)","Charges Billed by Courier Company (Rs.)" ,"Difference Between Expected Charges and Billed Charges (Rs.)"]
os.chdir(r"C:\Users\YADAV\OneDrive\Desktop\project")

df_coo= pd.DataFrame(pd.read_excel("Company X - Order Report.xlsx"))
df_csm = pd.DataFrame(pd.read_excel("Company X - SKU Master.xlsx"))
df_cpz= pd.DataFrame(pd.read_excel("Company X - Pincode Zones.xlsx"))


df_cur_inv= pd.DataFrame(pd.read_excel("Courier Company - Invoice.xlsx"))
df_cou_rates = pd.DataFrame(pd.read_excel("Courier Company - Rates.xlsx"))


def add_rate(df1, df2):
    # df1["Fixed_rate", "Additional_rate", "Rto_Fixedrate", "Rto_additional_rate"]= "","","",""
    df= []
    for grpname, elemnt in df1.groupby("Zone"):
        filtered_rates = df2[[cols for cols in df2.columns if f"_{grpname}_" in cols]]
        elemnt["Forward_Fixed"] = filtered_rates[f"fwd_{grpname}_fixed"][0]
        elemnt["Forward_Additional"] = filtered_rates[f"fwd_{grpname}_additional"][0]
        elemnt["Rto_Fixed"] = filtered_rates[f"rto_{grpname}_fixed"][0]
        elemnt["Rto_Additional"] = filtered_rates[f"rto_{grpname}_additional"][0]
        df.append(elemnt)

    return pd.concat(df)

def calculate_payments(df1):
    df = []
    for grpname, elemnt in df1.groupby(["Type of Shipment", "Weight_slabs"]):
        print(grpname)
        if ("RTO" not in grpname[0]) :
            elemnt["Expected Charge as per X (Rs.)"] = elemnt["Forward_Fixed"] + (elemnt["Weight_multiple"] * elemnt["Forward_Additional"])
        else:
            elemnt["Expected Charge as per X (Rs.)"] = (elemnt["Forward_Fixed"] + (elemnt["Weight_multiple"] * elemnt["Forward_Additional"])) + \
                                    (elemnt["Rto_Fixed"] + (elemnt["Weight_multiple"] * elemnt["Rto_Additional"]))
        df.append(elemnt)
    return pd.concat(df)

# df_cpz = df_cpz[~df_cpz.duplicated()]
df_cpz = df_cpz.merge(df_cur_inv[["AWB Code", "Order ID","Type of Shipment", "Warehouse Pincode","Customer Pincode"]], how= "left" , on=[ "Warehouse Pincode","Customer Pincode"])
df_coo = pd.merge(df_coo, df_csm, how= "left", on="SKU")
# df_coo= df_coo[~df_coo.duplicated()]
df_coo = df_coo.merge(df_cpz.rename(columns={"Order ID":"ExternOrderNo"}), how="left", on= "ExternOrderNo")

df_coo= df_coo.groupby("ExternOrderNo").agg({"Order Qty":"sum", "Weight (g)":"sum", "Zone": "unique"}).reset_index()
df_coo["Zone"] = df_coo["Zone"].str[0]
df_coo = add_rate(df_coo, df_cou_rates)
print(df_coo)
df_coo["Weight_multiple"] = df_coo["Weight (g)"].apply(lambda x : math.ceil(x/500) -1 )
# df_coo["Weight_slabs"] =  df_coo["Weight (g)"].apply(lambda x:round(x/500,1) + round((round(x/500) + 0.5) - round(x/500,1),1) )

bins= [1,500,1000,1500, 2000, 2500, 3000]
labels= ["0.5", "1.0", "1.5", "2.0", "2.5", "3.0"]
df_coo["Weight_slabs"] = pd.cut(df_coo["Weight (g)"], bins= bins , labels= labels, include_lowest=True)



df_coo = df_coo.merge(df_cpz[["Order ID", "Type of Shipment", "AWB Code", "Zone"]].rename(columns={"Order ID":"ExternOrderNo"}), how="left", on= "ExternOrderNo")

df_coo = calculate_payments(df_coo)
df_coo.rename(columns= {"ExternOrderNo": "Order ID"}, inplace=True)
# df_coo.groupby("Order ID").agg({"Order Qty":"sum", "Weight (g)":"sum"})
df_coo= df_coo.merge(df_cur_inv[["AWB Code", "Order ID", "Charged Weight", "Billing Amount (Rs.)"]], how="left", on=["AWB Code", "Order ID"])
df_coo = df_coo[~df_coo.duplicated()]

bins= [0, 0.5,1.0,1.5,2.0, 2.5, 3.0, 3.5, 4.0, 4.5]
labels= ["0.5", "1.0", "1.5", "2.0", "2.5", "3.0", "3.5", "4.0", "4.5"]
df_coo["Weight_slabs_by_cc"] = pd.cut(df_coo["Charged Weight"], bins= bins , labels= labels, include_lowest=True)


df_coo = df_coo[['Order ID','AWB Code','Weight (g)', 'Weight_slabs', 'Charged Weight', 'Weight_slabs_by_cc', 'Zone_x', 'Zone_y', 'Expected Charge as per X (Rs.)', 'Billing Amount (Rs.)']]
df_coo["Diffrence_amount"] = df_coo['Expected Charge as per X (Rs.)']-  df_coo['Billing Amount (Rs.)']
df_coo["Charged_amount_status"]=np.where(df_coo['Expected Charge as per X (Rs.)'] == df_coo['Billing Amount (Rs.)'], "correctly_charged", np.where(df_coo['Expected Charge as per X (Rs.)']> df_coo['Billing Amount (Rs.)'], "undercharged", "overcharged"))

df2 =df_coo.groupby(["Charged_amount_status"]).agg({"Charged_amount_status":"count", "Diffrence_amount":"sum"})
df_coo= df_coo.iloc[:,:-1]
df_coo.columns= final_column

with pd.ExcelWriter("final_report.xlsx") as writer:
    df2.to_excel(writer, sheet_name="Summary", index=False)
    df_coo.to_excel(writer, sheet_name="Calculations", index=False)

