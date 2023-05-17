import mysql.connector as mysql
import tkinter as tk
from tkinter import ttk, messagebox
import decimal
db = mysql.connect(host="localhost",user="root",password="spidy",database="mpdre")
cursor = db.cursor()
style = ttk.Style()
style.theme_use("alt")
style.configure("Custom.TFrame", background="#CCCCFF")
style.configure("Custom.TLabelframe", bordercolor="black", background="#CCCCFF")
style.configure("Custom.TLabelframe.Label", background="#CCCCFF")
style.configure("Custom.TLabel", background="#CCCCFF")
style.configure("Custom.TEntry", background="lightblue")
window=tk.Tk()
window.title("Stocks")
window.iconbitmap("mines.ico")
window.configure(background="#CCCCFF")
frame=ttk.Frame(window,padding=20,style="Custom.TFrame")
frame.pack()
# --- Stocks ---
data2_frame=ttk.LabelFrame(frame, text="Stocks")
data2_frame.grid(row=4, column=0, padx=40, pady=10)
# --- Stocks Labels ---
_3m_excel_label_ = ttk.Label(data2_frame, text="3m Excel")
_3m_excel_label_.grid(row=0, column=0)
_4m_excel_label_ = ttk.Label(data2_frame, text="4m Excel")
_4m_excel_label_.grid(row=1, column=0)
_5m_excel_label_ = ttk.Label(data2_frame, text="5m Excel")
_5m_excel_label_.grid(row=2, column=0)
_6m_excel_label_ = ttk.Label(data2_frame, text="6m Excel")
_6m_excel_label_.grid(row=3, column=0)
_7m_excel_label_ = ttk.Label(data2_frame, text="7m Excel")
_7m_excel_label_.grid(row=4, column=0)
_8m_excel_label_ = ttk.Label(data2_frame, text="8m Excel")
_8m_excel_label_.grid(row=5, column=0)
_9m_excel_label_ = ttk.Label(data2_frame, text="9m Excel")
_9m_excel_label_.grid(row=6, column=0)
_10m_excel_label_ = ttk.Label(data2_frame, text="10m Excel")
_10m_excel_label_.grid(row=7, column=0)
_12m_excel_label_ = ttk.Label(data2_frame, text="12m Excel")
_12m_excel_label_.grid(row=8, column=0)
_15m_excel_label_ = ttk.Label(data2_frame, text="15m Excel")
_15m_excel_label_.grid(row=9, column=0)
_slurry_label_ = ttk.Label(data2_frame, text="Slurry")
_slurry_label_.grid(row=10, column=0)
_slurry_big_label_ = ttk.Label(data2_frame, text="Slurry Big")
_slurry_big_label_.grid(row=11, column=0)
_ed_label_ = ttk.Label(data2_frame, text="E.D")
_ed_label_.grid(row=12, column=0)
_d_det_label_ = ttk.Label(data2_frame, text="D.Det")
_d_det_label_.grid(row=13, column=0)
_df_5gms_label_ = ttk.Label(data2_frame, text="DF 5gms")
_df_5gms_label_.grid(row=14, column=0)
_df_10gms_label_ = ttk.Label(data2_frame, text="DF 10gms")
_df_10gms_label_.grid(row=15, column=0)
# --- Stocks Labels Completed ---
# --- Stocks Queries ---
_3m_excel_query = "select sum(_3m_total)-sum(_3m_used) from explosives"
_4m_excel_query="select sum(_4m_total)-sum(_4m_used) from explosives"
_5m_excel_query="select sum(_5m_total)-sum(_5m_used) from explosives"
_6m_excel_query="select sum(_6m_total)-sum(_6m_used) from explosives"
_7m_excel_query="select sum(_7m_total)-sum(_7m_used) from explosives"
_8m_excel_query="select sum(_8m_total)-sum(_8m_used) from explosives"
_9m_excel_query="select sum(_9m_total)-sum(_9m_used) from explosives"
_10m_excel_query="select sum(_10m_total)-sum(_10m_used) from explosives"
_12m_excel_query="select sum(_12m_total)-sum(_12m_used) from explosives"
_15m_excel_query="select sum(_15m_total)-sum(_15m_used) from explosives"
_slurry_query="select sum(_slurry_total)-sum(_slurry_used) from explosives"
_slurry_big_query="select sum(_slurry_big_total)-sum(_slurry_big_used) from explosives"
_ed_query="select sum(_ed_total)-sum(_ed_used) from explosives"
_d_det_query="select sum(_ddet_total)-sum(_ddet_used) from explosives"
_df_5gms_query="select sum(_df_5_gms_total)-sum(_df_5_gms_used) from explosives"
_df_10gms_query="select sum(_df_10_gms_total)-sum(_df_10_gms_used) from explosives"

# --- Stocks Queries Completed ---
# --- Stocks Data ---
cursor.execute("USE mpdre")
_3m_excel_data = cursor.execute(_3m_excel_query)
_3m_excel_query_data = cursor.fetchone()
_4m_excel_data = cursor.execute(_4m_excel_query)
_4m_excel_query_data = cursor.fetchone()
_5m_excel_data = cursor.execute(_5m_excel_query)
_5m_excel_query_data = cursor.fetchone()
_6m_excel_data = cursor.execute(_6m_excel_query)
_6m_excel_query_data = cursor.fetchone()
_7m_excel_data = cursor.execute(_7m_excel_query)
_7m_excel_query_data = cursor.fetchone()
_8m_excel_data = cursor.execute(_8m_excel_query)
_8m_excel_query_data = cursor.fetchone()
_9m_excel_data = cursor.execute(_9m_excel_query)
_9m_excel_query_data = cursor.fetchone()
_10m_excel_data = cursor.execute(_10m_excel_query)
_10m_excel_query_data = cursor.fetchone()
_12m_excel_data = cursor.execute(_12m_excel_query)
_12m_excel_query_data = cursor.fetchone()
_15m_excel_data = cursor.execute(_15m_excel_query)
_15m_excel_query_data = cursor.fetchone()
_slurry_data = cursor.execute(_slurry_query)
_slurry_query_data = cursor.fetchone()
_slurry_big_data = cursor.execute(_slurry_big_query)
_slurry_big_query_data = cursor.fetchone()
_ed_data = cursor.execute(_ed_query)
_ed_query_data = cursor.fetchone()
_d_det_data = cursor.execute(_d_det_query)
_d_det_query_data = cursor.fetchone()
_df_5gms_data = cursor.execute(_df_5gms_query)
_df_5gms_query_data = cursor.fetchone()
_df_10gms_data = cursor.execute(_df_10gms_query)
_df_10gms_query_data = cursor.fetchone()
# --- Stocks Data Completed ---
# --- Stocks Data Entry ---
_3m_excel_entry1_ = ttk.Label(data2_frame, text="None")
_3m_excel_entry1_.grid(row=0, column=1)
_4m_excel_entry1_ = ttk.Label(data2_frame, text="None")
_4m_excel_entry1_.grid(row=1, column=1)
_5m_excel_entry1_ = ttk.Label(data2_frame, text="None")
_5m_excel_entry1_.grid(row=2, column=1)
_6m_excel_entry1_ = ttk.Label(data2_frame, text="None")
_6m_excel_entry1_.grid(row=3, column=1)
_7m_excel_entry1_ = ttk.Label(data2_frame, text="None")
_7m_excel_entry1_.grid(row=4, column=1)
_8m_excel_entry1_ = ttk.Label(data2_frame, text="None")
_8m_excel_entry1_.grid(row=5, column=1)
_9m_excel_entry1_ = ttk.Label(data2_frame, text="None")
_9m_excel_entry1_.grid(row=6, column=1)
_10m_excel_entry1_ = ttk.Label(data2_frame, text="None")
_10m_excel_entry1_.grid(row=7, column=1)
_12m_excel_entry1_ = ttk.Label(data2_frame, text="None")
_12m_excel_entry1_.grid(row=8, column=1)
_15m_excel_entry1_ = ttk.Label(data2_frame, text="None")
_15m_excel_entry1_.grid(row=9, column=1)
_slurry_entry1_ = ttk.Label(data2_frame, text="None")
_slurry_entry1_.grid(row=10, column=1)
_slurry_big_entry1_ = ttk.Label(data2_frame, text="None")
_slurry_big_entry1_.grid(row=11, column=1)
_ed_entry1_ = ttk.Label(data2_frame, text="None")
_ed_entry1_.grid(row=12, column=1)
_d_det_entry1_ = ttk.Label(data2_frame, text="None")
_d_det_entry1_.grid(row=13, column=1)
_df_5gms_entry1_ = ttk.Label(data2_frame, text="None")
_df_5gms_entry1_.grid(row=14, column=1)
_df_10gms_entry1_ = ttk.Label(data2_frame, text="None")
_df_10gms_entry1_.grid(row=15, column=1)
# --- Stocks Data Entry Completed ---
# --- Refresh Button ---
def refresh_stocks():
    _3m_excel_query = "select sum(_3m_total)-sum(_3m_used) from explosives;"
    _4m_excel_query="select sum(_4m_total)-sum(_4m_used) from explosives;"
    _5m_excel_query="select sum(_5m_total)-sum(_5m_used) from explosives;"
    _6m_excel_query="select sum(_6m_total)-sum(_6m_used) from explosives;"
    _7m_excel_query="select sum(_7m_total)-sum(_7m_used) from explosives;"
    _8m_excel_query="select sum(_8m_total)-sum(_8m_used) from explosives;"
    _9m_excel_query="select sum(_9m_total)-sum(_9m_used) from explosives;"
    _10m_excel_query="select sum(_10m_total)-sum(_10m_used) from explosives;"
    _12m_excel_query="select sum(_12m_total)-sum(_12m_used) from explosives;"
    _15m_excel_query="select sum(_15m_total)-sum(_15m_used) from explosives;"
    _slurry_query="select sum(_slurry_total)-sum(_slurry_used) from explosives;"
    _slurry_big_query="select sum(_slurry_big_total)-sum(_slurry_big_used) from explosives;"
    _ed_query="select sum(_ed_total)-sum(_ed_used) from explosives;"
    _d_det_query="select sum(_ddet_total)-sum(_ddet_used) from explosives;"
    _df_5gms_query="select sum(_df_5_gms_total)-sum(_df_5_gms_used) from explosives;"
    _df_10gms_query="select sum(_df_10_gms_total)-sum(_df_10_gms_used) from explosives;"
    cursor.execute("USE mpdre")
    cursor.execute(_3m_excel_query)
    r=cursor.fetchone()
    v1=r[0]
    _3m_excel_query_data = float(decimal.Decimal(v1))
    cursor.execute(_4m_excel_query)
    r=cursor.fetchone()
    v2=r[0]
    _4m_excel_query_data = float(decimal.Decimal(v2))
    cursor.execute(_5m_excel_query)
    r=cursor.fetchone()
    v3=r[0]
    _5m_excel_query_data = float(decimal.Decimal(v3))
    cursor.execute(_6m_excel_query)
    r=cursor.fetchone()
    v4=r[0]
    _6m_excel_query_data = float(decimal.Decimal(v4))
    cursor.execute(_7m_excel_query)
    r=cursor.fetchone()
    v5=r[0]
    _7m_excel_query_data = float(decimal.Decimal(v5))
    cursor.execute(_8m_excel_query)
    r=cursor.fetchone()
    v6=r[0]
    _8m_excel_query_data = float(decimal.Decimal(v6))
    cursor.execute(_9m_excel_query)
    r=cursor.fetchone()
    v7=r[0]
    _9m_excel_query_data = float(decimal.Decimal(v7))
    cursor.execute(_10m_excel_query)
    r=cursor.fetchone()
    v8=r[0]
    _10m_excel_query_data = float(decimal.Decimal(v8))  
    cursor.execute(_12m_excel_query)
    r=cursor.fetchone()
    v9=r[0]
    _12m_excel_query_data = float(decimal.Decimal(v9))
    cursor.execute(_15m_excel_query)
    r=cursor.fetchone()
    v10=r[0]
    _15m_excel_query_data = float(decimal.Decimal(v10))
    cursor.execute(_slurry_query)
    r=cursor.fetchone()
    v11=r[0]
    _slurry_query_data = float(decimal.Decimal(v11))
    cursor.execute(_slurry_big_query)
    r=cursor.fetchone()
    v12=r[0]
    _slurry_big_query_data = float(decimal.Decimal(v12))
    cursor.execute(_ed_query)
    r=cursor.fetchone()
    v13=r[0]
    _ed_query_data = float(decimal.Decimal(v13))
    cursor.execute(_d_det_query)
    r=cursor.fetchone()
    v14=r[0]
    _d_det_query_data = float(decimal.Decimal(v14))
    cursor.execute(_df_5gms_query)
    r=cursor.fetchone()
    v15=r[0]
    _df_5gms_query_data = float(decimal.Decimal(v15))
    cursor.execute(_df_10gms_query)
    r=cursor.fetchone()
    v16=r[0]
    _df_10gms_query_data = float(decimal.Decimal(v16))
    _3m_excel_entry1_.configure(text=_3m_excel_query_data)
    _4m_excel_entry1_.configure(text=_4m_excel_query_data)
    _5m_excel_entry1_.configure(text=_5m_excel_query_data)
    _6m_excel_entry1_.configure(text=_6m_excel_query_data)
    _7m_excel_entry1_.configure(text=_7m_excel_query_data)
    _8m_excel_entry1_.configure(text=_8m_excel_query_data)
    _9m_excel_entry1_.configure(text=_9m_excel_query_data)
    _10m_excel_entry1_.configure(text=_10m_excel_query_data)
    _12m_excel_entry1_.configure(text=_12m_excel_query_data)
    _15m_excel_entry1_.configure(text=_15m_excel_query_data)
    _slurry_entry1_.configure(text=_slurry_query_data)
    _slurry_big_entry1_.configure(text=_slurry_big_query_data)
    _ed_entry1_.configure(text=_ed_query_data)
    _d_det_entry1_.configure(text=_d_det_query_data)
    _df_5gms_entry1_.configure(text=_df_5gms_query_data)
    _df_10gms_entry1_.configure(text=_df_10gms_query_data)  

refresh_button=ttk.Button(data2_frame,text="Refresh",command=refresh_stocks)
refresh_button.grid(row=16,column=0)
# --- Refresh Button Completed ---
window.mainloop()