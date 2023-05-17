import mysql.connector
from mysql.connector import Error
import tkinter as tk
from tkinter import ttk, messagebox
from program import main
import datetime
from tempfile import mktemp
from os import system, startfile
import tkinter as tk
from tkinter import ttk, messagebox
import mysql.connector
from fpdf import FPDF
title = "Lakshmi Blue Metals"

window = tk.Tk()
window.title("Report")
window.iconbitmap("906310.ico")
window.configure(background="#CCCCFF")
frame = ttk.Frame(window, padding=20)
frame.pack()
dataframe = ttk.LabelFrame(frame, text="Report", borderwidth=2, relief="groove")
dataframe.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")
date_label = ttk.Label(dataframe, text="Enter the no of the month:")
date_label.grid(row=0, column=0, padx=10, pady=10)
date_entry = ttk.Entry(dataframe, width=20)
date_entry.grid(row=0, column=1, padx=10, pady=10)
date_entry.focus()
def generate_report_purchased():
    class PDF(FPDF):
        def header(self):
            self.set_font('Arial', 'BIU', 18)
            title_w = self.get_string_width(title) + 6
            doc_w = self.w
            self.set_x((doc_w - title_w) / 2)
            self.set_draw_color(0, 80, 180)
            self.set_fill_color(230, 230, 0)
            self.set_text_color(220, 50, 50)
            self.set_line_width(1)
            self.cell(title_w, 10, title, border=1, ln=1, align='C', fill=1)
            self.set_font('Arial', 'BI', 14)
            self.cell(0, 10, txt="Mines Monthly Purchase Report", ln=1, align='C')
            self.ln(15)

        def footer(self):
            self.set_y(-15)
            self.set_font('Arial', 'I', 8)
            self.set_text_color(169, 169, 169)
            self.cell(0, 10, txt="Page " + str(self.page_no()) + "/{nb}", ln=1, align='C')

    date = date_entry.get()
    if not date:
        messagebox.showerror("Report", "Please enter date.")
        return
    try:
        connection = mysql.connector.connect(host='localhost',
                                             database='mpdre',
                                             user='root',
                                             password='spidy')

        sql_select_Query1 = "SELECT sum(_3m_total),sum(_4m_total),sum(_5m_total),sum(_6m_total),sum(_7m_total),sum(_8m_total),sum(_9m_total),sum(_10m_total),sum(_12m_total),sum(_15m_total),sum(_slurry_total),sum(_slurry_big_total),sum(_ed_total),sum(_ddet_total),sum(_df_5_gms_total),sum(_df_10_gms_total),sum(noofholes_2_5)*2.5,sum(noofholes_5)*5,sum(noofholes_6)*6,sum(noofholes_8)*8,sum(wagon_drill_4_5) FROM explosives WHERE month(date)=%s"

        with connection.cursor() as cursor:
            cursor.execute(sql_select_Query1, (date,))
            records = cursor.fetchall()

        if not records:
            messagebox.showerror("Report", "No data found for the given date.")
            return

        rec = []
        for row in records:
            rec.extend(row)
        rec = rec[:16] + [sum(rec[16:20])] + rec[20:]

        curre = int(date) - 1
        month = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]
        items = ["3m Excel", "4m  Excel", "5m Excel", "6m Excel", "7m Excel", "8m Excel", "9m Excel", "10m Excel", "12m Excel", "15m Excel", "Slurry", "Slurry Big", "ED", "DDet", "DF 5gms", "DF 10gms", "Total Compressor Ft", "Wagon Drill 4.5"]

        pdf = PDF('P', 'mm', 'A4')
        pdf.set_auto_page_break(auto=True, margin=10)
        pdf.alias_nb_pages()
        pdf.add_page()
        pdf.set_font("times", 'BU', 16)
        pdf.cell(0, 10, txt=f"Monthly Report of {month[curre]}", ln=1, align="L")

        # Generate table headers
        pdf.set_fill_color(230, 230, 230)
        pdf.set_text_color(0, 0, 0)
        pdf.set_font("Arial", "B", 12)
        headers = ["S.No", "Items", "Quantity Purchased"]
        col_widths = [30, 80, 80]
        row_height = pdf.font_size * 2
        for i,header in enumerate(headers):
            pdf.cell(col_widths[i], row_height, txt=header, border=1, fill=1, align="C")
        pdf.ln(row_height)

        # Generate table data
        pdf.set_font("Arial", "", 12)
        for i, item in enumerate(items):
            pdf.cell(col_widths[0], row_height, txt=str(i+1), border=1, align="C")
            pdf.cell(col_widths[1], row_height, txt=item, border=1, align="C")
            pdf.cell(col_widths[2], row_height, txt=str(rec[i]), border=1, align="C")
            pdf.ln(row_height)

        pdf.output("report_temp.pdf")
        window.destroy()
        connection.close()

    except mysql.connector.Error as e:
        messagebox.showerror("Error", "Error reading data from MySQL table: {}".format(e))
        window.destroy()
generate_button = ttk.Button(frame, text="Generate Report Purchase", command=generate_report_purchased)
generate_button.grid(row=1, column=0, padx=10, pady=10, sticky="nsew")
def generate_report_used():
    class PDF(FPDF):
        def header(self):
            self.set_font('Arial', 'BIU', 18)
            title_w = self.get_string_width(title) + 6
            doc_w = self.w
            self.set_x((doc_w - title_w) / 2)
            self.set_draw_color(0, 80, 180)
            self.set_fill_color(230, 230, 0)
            self.set_text_color(220, 50, 50)
            self.set_line_width(1)
            self.cell(title_w, 10, title, border=1, ln=1, align='C', fill=1)
            self.set_font('Arial', 'BI', 14)
            self.cell(0, 10, txt="Mines Monthly Report", ln=1, align='C')
            self.ln(15)

        def footer(self):
            self.set_y(-15)
            self.set_font('Arial', 'I', 8)
            self.set_text_color(169, 169, 169)
            self.cell(0, 10, txt="Page " + str(self.page_no()) + "/{nb}", ln=1, align='C')
    date = date_entry.get()
    if not date:
        messagebox.showerror("Report", "Please enter date.")
        return
    try:
        connection = mysql.connector.connect(host='localhost',
                                             database='mpdre',
                                             user='root',
                                             password='spidy')
        sql_select_Query1 = "SELECT sum(_3m_used),sum(_4m_used),sum(_5m_used),sum(_6m_used),sum(_7m_used),sum(_8m_used),sum(_9m_used),sum(_10m_used),sum(_12m_used),sum(_15m_used),sum(_slurry_used),sum(_slurry_big_used),sum(_ed_used),sum(_ddet_used),sum(_df_5_gms_used),sum(_df_10_gms_used) FROM explosives WHERE month(date)=%s"
        sql_select_Query2 = "SELECT sum(_3m_total),sum(_4m_total),sum(_5m_total),sum(_6m_total),sum(_7m_total),sum(_8m_total),sum(_9m_total),sum(_10m_total),sum(_12m_total),sum(_15m_total),sum(_slurry_total),sum(_slurry_big_total),sum(_ed_total),sum(_ddet_total),sum(_df_5_gms_total),sum(_df_10_gms_total) FROM explosives WHERE month(date)=%s"
        cursor = connection.cursor()
        cursor.execute(sql_select_Query1, (date,))
        records1 = cursor.fetchall()
        cursor.execute(sql_select_Query2, (date,))
        records2 = cursor.fetchall()
        rec2 = []
        rec1 = []
        for row in records1:
            for i in range(len(row)):
                rec2.append(row[i])
        for row in records2:
            for i in range(len(row)):
                rec1.append(row[i])
        s_no = []
        for i in range(1, len(rec2) + 1):
            s_no.append(i)
        if not records1 and not records2:
            messagebox.showerror("Report", "No data found for the given date.")
            return
        curre = int(date) - 1
        month = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]
        items=["3m Excel", "4m Excel", "5m Excel", "6m Excel", "7m Excel", "8m Excel", "9m Excel", "10m Excel", "12m Excel", "15m Excel", "Slurry", "Slurry Big", "ED", "DDet", "DF 5gms", "DF 10gms", "Total Compressor Ft", "Wagon Drill 4.5"]
        pdf = PDF('P', 'mm', 'A4')
        pdf.set_auto_page_break(auto=True, margin=10)
        pdf.alias_nb_pages()
        pdf.add_page()
        pdf.set_font("times", 'BU', 16)
        pdf.cell(0, 10, txt=f"Monthly Report of {month[curre]}", ln=1, align="L")
        
        # Generate table headers
        pdf.set_fill_color(230, 230, 230)
        pdf.set_text_color(0, 0, 0)
        pdf.set_font("Arial", "B", 12)
        headers = ["S.No", "Items", "Quantity Purchased", "Quantity Used"]
        col_widths = [40, 50, 50, 50]
        row_height = pdf.font_size * 2
        for i in range(len(headers)):
            pdf.cell(col_widths[i], row_height, txt=headers[i], border=1, fill=1, align="C")
        pdf.ln(row_height)

        # Generate table data
        pdf.set_font("Arial", "", 12)
        for i in range(len(s_no)):
            pdf.cell(col_widths[0], row_height, txt=str(s_no[i]), border=1, align="C")
            pdf.cell(col_widths[1], row_height, txt=str(items[i]), border=1, align="C")
            pdf.cell(col_widths[2], row_height, txt=str(rec1[i]), border=1, align="C")
            pdf.cell(col_widths[3], row_height, txt=str(rec2[i]), border=1, align="C")
            pdf.ln(row_height)

        pdf.output("report_temp_1.pdf")
        window.destroy()

    except mysql.connector.Error as e:
        messagebox.showerror("Error", "Error reading data from MySQL table: {}".format(e))
        window.destroy()
generate_button_used = ttk.Button(frame, text="Generate Report Used", command=generate_report_used)
generate_button_used.grid(row=2, column=0, padx=10, pady=10, sticky="nsew")
window.mainloop()


