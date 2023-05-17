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

title = "Sri Lakshmi Blue Metals"


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
        self.cell(0, 10, txt="Expl Monthly Purchase Report", ln=1, align='C')
        self.ln(15)

    def footer(self):
        self.set_y(-15)
        self.set_font('Arial', 'I', 8)
        self.set_text_color(169, 169, 169)
        self.cell(0, 10, txt="Page " + str(self.page_no()) + "/{nb}", ln=1, align='C')


window = tk.Tk()
window.title("Report")
window.iconbitmap("906310.ico")
window.geometry("325x200")
window.configure(background="#CCCCFF")
frame = ttk.Frame(window, padding=20)
frame.pack()
dataframe = ttk.LabelFrame(frame, text="Report", borderwidth=2, relief="groove")
dataframe.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")
date_label = ttk.Label(dataframe, text="Enter the no of the month:")
date_label.grid(row=0, column=0, padx=10, pady=10)
date_entry = ttk.Entry(dataframe, width=20)
date_entry.grid(row=0, column=1, padx=10, pady=10)


def generate_report():
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
        cursor = connection.cursor()
        cursor.execute(sql_select_Query1, (date,))
        records = cursor.fetchall()
        rec = []
        for row in records:
            for i in range(len(row)):
                rec.append(row[i])

        rec[16] = rec[16] + rec[17] + rec[18] + rec[19]
        rec.pop(17)
        rec.pop(18)
        rec.pop(19)
        s_no = []
        for i in range(1, len(rec) + 1):
            s_no.append(i)
        if not records:
            messagebox.showerror("Report", "No data found for the given date.")
            return
        curre = int(date) - 1
        month = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]
        items=["3m", "4m", "5m", "6m", "7m", "8m", "9m", "10m", "12m", "15m", "slurry", "slurry big", "ed", "ddet", "df 5 gms", "df 10 gms", "compressor_ft", "wagon_drill_4_5"]
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
        for i in range(len(headers)):
            pdf.cell(col_widths[i], row_height, txt=headers[i], border=1, fill=1, align="C")
        pdf.ln(row_height)

        # Generate table data
        pdf.set_font("Arial", "", 12)
        for i in range(len(s_no)):
            pdf.cell(col_widths[0], row_height, txt=str(s_no[i]), border=1, align="C")
            pdf.cell(col_widths[1], row_height, txt=str(items[i]), border=1, align="C")
            pdf.cell(col_widths[2], row_height, txt=str(rec[i]), border=1, align="C")
            pdf.ln(row_height)

        pdf.output("report_temp.pdf")
        window.destroy()

    except mysql.connector.Error as e:
        messagebox.showerror("Error", "Error reading data from MySQL table: {}".format(e))
        window.destroy()
generate_button = ttk.Button(frame, text="Generate Report", command=generate_report)
generate_button.grid(row=1, column=0, padx=10, pady=10, sticky="nsew")

window.mainloop()


