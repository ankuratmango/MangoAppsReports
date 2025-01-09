import xlsxwriter
from collections import Counter
import random


def generate_random_colors(count):
    colors = []
    while len(colors) < count:
        color = random_color()
        if color not in colors:
            colors.append(color)
    return colors

def random_color():
    golden_ratio_conjugate = 0.618033988749895
    h = random.random()
    h += golden_ratio_conjugate
    h %= 1
    r, g, b = hsv_to_rgb(h, 0.5, 0.95)
    return f"{to_hex(r)}{to_hex(g)}{to_hex(b)}"

def to_hex(n):
    return f"{n:02X}"

def hsv_to_rgb(h, s, v):
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


# Sample data
chart_data = [
    ("Namrata Puranik Puranik", 51),
    ("Gauri Puranik", 16),
    ("Aalkhimovich aalk", 16),
    ("Alumni User", 10),
    ("Ankur Tripathi", 7),
]

custom_colors = generate_random_colors(len(chart_data))

output_path = "static_chart_no_data.xlsx"
workbook = xlsxwriter.Workbook(output_path)
summary_worksheet = workbook.add_worksheet("Summary")
mrr_worksheet = workbook.add_worksheet("MRR")

mrr_worksheet.hide()

for row_num, (name, value) in enumerate(chart_data, start=1):
    mrr_worksheet.write(row_num, 0, name)
    mrr_worksheet.write(row_num, 1, value)

chart = workbook.add_chart({"type": "column"})
summary_worksheet.select()
mrr_worksheet.hide()

points = [{"fill": {"color": f"#{color}"}} for color in custom_colors]
chart.add_series({
    "categories": f"=MRR!$A$2:$A${len(chart_data) + 1}",
    "values": f"=MRR!$B$2:$B${len(chart_data) + 1}",
    "data_labels": {"value": True},
    "points": points,  
})

chart.set_title({"name": "Most Recognized Recipients"})
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

summary_worksheet.insert_chart("D2", chart)
workbook.close()

print(f"Excel file created successfully: {output_path}")
