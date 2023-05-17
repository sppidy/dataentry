from fpdf import FPDF
title="Sri Lakshmi Blue Metals"
class PDF(FPDF):
    def header(self):
        self.set_font('Arial', 'BIU', 18)
        title_w=self.get_string_width(title)+6
        doc_w=self.w
        self.set_x((doc_w-title_w)/2)
        self.set_draw_color(0,80,180)
        self.set_fill_color(230,230,0)
        self.set_text_color(220,50,50)
        self.set_line_width(1)
        self.cell(title_w,10,title,border=1,ln=1,align='C',fill=1)
        self.set_font('Arial', 'BU', 14)
        self.cell(0,10,txt="Monthly Report",ln=1,align='C')
        self.ln(15)
    def footer(self):
        self.set_y(-15)
        self.set_font('Arial', 'I', 8)
        self.set_text_color(169,169,169)
        self.cell(0,10,txt="Page "+str(self.page_no())+"/{nb}",ln=1,align='C')

pdf = PDF('L', 'mm', 'A4')
pdf.set_auto_page_break(auto=True, margin=10)
pdf.alias_nb_pages()
pdf.add_page()
pdf.set_font("times", 'BU',16)
pdf.cell(0, 10, txt="Monthly Report", ln=1, align="L")

pdf.output("report_temp.pdf")