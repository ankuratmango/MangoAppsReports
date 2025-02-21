import xlsxwriter
import random

class ChartGenerator:
    def __init__(self, chart_data, chart_data_issuers, output_path="static_chart.xlsx"):
        self.chart_data = chart_data
        self.chart_data_issuers = chart_data_issuers
        self.output_path = output_path
        self.workbook = xlsxwriter.Workbook(self.output_path)
        self.summary_worksheet = self.workbook.add_worksheet("Summary")
    
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
            r, g, b = 0, 0, 0  # Fallback
        
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
    
    def generate_excel(self):
        self.create_chart("MRR", self.chart_data, "Most Recognized Recipients", "D2")
        self.create_chart("MRR_I", self.chart_data_issuers, "Top Issuing Users", "D20")
        self.workbook.close()
        print(f"Excel file created successfully: {self.output_path}")

