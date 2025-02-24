import xlsxwriter
import random

class ChartGenerator:
    def __init__(self, xls_data, output_path="static_chart.xlsx"):
        self.xls_data = xls_data
        self.output_path = output_path
        self.workbook = xlsxwriter.Workbook(self.output_path)
    
    def generate_random_colors(self, count):
        colors = []
        while len(colors) < count:
            color = self.random_color()
            if color not in colors:
                colors.append(color)
        return colors

    def random_color(self):
        golden_ratio_conjugate = 0.618033988749895
        h = random.random()
        h += golden_ratio_conjugate
        h %= 1
        r, g, b = self.hsv_to_rgb(h, 0.5, 0.95)
        return f"{self.to_hex(r)}{self.to_hex(g)}{self.to_hex(b)}"

    def to_hex(self, n):
        return f"{n:02X}"

    def hsv_to_rgb(self, h, s, v):
        h_i = int(h * 6)
        f = h * 6 - h_i
        p = v * (1 - s)
        q = v * (1 - f * s)
        t = v * (1 - (1 - f) * s)
        
        if h_i == 0:
            r, g, b = v, t, p
        elif h_i == 1:
            r, g, b = q, v, p
        elif h_i == 2:
            r, g, b = p, v, t
        elif h_i == 3:
            r, g, b = p, q, v
        elif h_i == 4:
            r, g, b = t, p, v
        elif h_i == 5:
            r, g, b = v, p, q
        else:
            r, g, b = 0, 0, 0  
        
        return int(r * 256), int(g * 256), int(b * 256)
    
    def create_chart(self, sheet_name, data, chart_title, insert_position):
        custom_colors = self.generate_random_colors(len(data))
        mrr_worksheet = self.workbook.add_worksheet(sheet_name)
        mrr_worksheet.hide()
        
        for row_num, (name, value) in enumerate(data, start=1):
            mrr_worksheet.write(row_num, 0, name)
            mrr_worksheet.write(row_num, 1, value)
        
        chart = self.workbook.add_chart({"type": "column"})
        self.summary_worksheet.select()
        mrr_worksheet.hide()
        points = [{"fill": {"color": f"#{color}"}} for color in custom_colors]
        chart.add_series({
            "categories": f"={sheet_name}!$A$2:$A${len(data) + 1}",
            "values": f"={sheet_name}!$B$2:$B${len(data) + 1}",
            "data_labels": {"value": True},
            "points": points,  
        })
        
        chart.set_title({"name": chart_title})
        chart.set_x_axis({
            "name": "Recipient",
            "name_font": {"size": 12, "bold": True},
            "num_font": {"rotation": -45},  
        })
        chart.set_y_axis({
            "name": "Count",
            "name_font": {"size": 12, "bold": True},
        })
        chart.set_legend({"position": "right"})
        
        self.summary_worksheet.insert_chart(insert_position, chart)
    
    def generate_excel_summary(self, sheet):
        self.summary_worksheet = self.workbook.add_worksheet(sheet)
        bold_format = self.workbook.add_format({"bold": True})
        self.summary_worksheet.write("D2", self.xls_data['first_header'], bold_format)
        self.summary_worksheet.write("D3", self.xls_data['second_header'])
        self.create_chart("MRR", self.xls_data['chart_data'], "Most Recognized Recipients", "D6")
        self.create_chart("MRR_I", self.xls_data['chart_data_issuers'], "Top Issuing Users", "D25")
        print(f"Excel file created successfully: {self.output_path}")

    def generate_excel_data(self, sheet, headers, recognition_hash):
        self.data_worksheet = self.workbook.add_worksheet(sheet)
        self.data_worksheet.select()
        
        header_format = self.workbook.add_format({
            "bold": True,
            "bg_color": "#2f75b5",
            "font_color": "white",
            "align": "center",
            "border": 1
        })

        self.data_worksheet.write_row(0, 0, headers, header_format)
        self.data_worksheet.autofilter(0, 0, 0, len(headers) - 1)

        for col_num, header in enumerate(headers):
            self.data_worksheet.set_column(col_num, col_num, len(header) + 5)

        date_format = self.workbook.add_format({'num_format': 'mm/dd/yyyy'})
        row_num = 1
        for entry in recognition_hash.values():
            self.data_worksheet.write(row_num, 0, entry.get("award_recognition_name", ""))  
            self.data_worksheet.write(row_num, 1, entry.get("award_recognition_category", ""))  
            self.data_worksheet.write(row_num, 2, entry.get("message", ""))  
            self.data_worksheet.write(row_num, 3, entry.get("message_by", ""))  
            self.data_worksheet.write(row_num, 4, entry.get("given_by_emp_id", ""))  
            self.data_worksheet.write(row_num, 5, "")  
            self.data_worksheet.write(row_num, 6, entry.get("message_to", ""))  
            self.data_worksheet.write(row_num, 7, entry.get("message_to_emp_id", ""))  
            self.data_worksheet.write(row_num, 8, "")  
            self.data_worksheet.write(row_num, 9, "")  
        
            date_value = entry.get("message_given_on", "")
            if date_value:
                self.data_worksheet.write_datetime(row_num, 10, date_value, date_format)  
        
            self.data_worksheet.write(row_num, 11, int(entry.get("award_points", 0)))  
            self.data_worksheet.write(row_num, 12, entry.get("award_reward_points", ""))  
            self.data_worksheet.write(row_num, 13, entry.get("award_total_reward_points", ""))  
            self.data_worksheet.write(row_num, 14, entry.get("team_name", "")) 
            self.data_worksheet.write(row_num, 15, "")  

            row_num += 1  

        self.workbook.close()
        

