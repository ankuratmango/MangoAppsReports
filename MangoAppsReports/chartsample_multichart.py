import xlsxwriter
import random
from typing import List, Tuple

class ColorGenerator:
    @staticmethod
    def generate_random_colors(count: int) -> List[str]:
        """Generate a list of unique random colors."""
        colors = set()
        while len(colors) < count:
            colors.add(ColorGenerator._random_color())
        return list(colors)

    @staticmethod
    def _random_color() -> str:
        """Generate a single random color using HSV to RGB conversion."""
        golden_ratio = 0.618033988749895
        h = (random.random() + golden_ratio) % 1
        r, g, b = ColorGenerator._hsv_to_rgb(h, 0.5, 0.95)
        return f"{r:02X}{g:02X}{b:02X}"

    @staticmethod
    def _hsv_to_rgb(h: float, s: float, v: float) -> Tuple[int, int, int]:
        """Convert HSV color values to RGB."""
        h_i = int(h * 6)
        f = h * 6 - h_i
        p, q, t = v * (1 - s), v * (1 - f * s), v * (1 - (1 - f) * s)
        
        rgb_map = {
            0: (v, t, p),
            1: (q, v, p),
            2: (p, v, t),
            3: (p, q, v),
            4: (t, p, v),
            5: (v, p, q)
        }
        r, g, b = rgb_map.get(h_i, (0, 0, 0))
        return int(r * 256), int(g * 256), int(b * 256)

class ChartCreator:
    def __init__(self, workbook: xlsxwriter.Workbook):
        self.workbook = workbook
        self.summary_ws = workbook.add_worksheet("Summary")

    def _write_data_to_worksheet(self, worksheet: xlsxwriter.Worksheet, 
                               data: List[Tuple[str, int]]) -> None:
        """Write data to worksheet starting from row 1."""
        for row, (name, value) in enumerate(data, 1):
            worksheet.write(row, 0, name)
            worksheet.write(row, 1, value)

    def _create_chart(self, data: List[Tuple[str, int]], sheet_name: str, 
                     title: str, x_axis_name: str) -> xlsxwriter.chart.Chart:
        """Create and configure a column chart."""
        worksheet = self.workbook.add_worksheet(sheet_name)
        worksheet.hide()
        self._write_data_to_worksheet(worksheet, data)
        
        chart = self.workbook.add_chart({"type": "column"})
        colors = ColorGenerator.generate_random_colors(len(data))
        points = [{"fill": {"color": f"#{color}"}} for color in colors]
        
        chart.add_series({
            "categories": f"={sheet_name}!$A$2:$A${len(data) + 1}",
            "values": f"={sheet_name}!$B$2:$B${len(data) + 1}",
            "data_labels": {"value": True},
            "points": points,
        })
        
        chart.set_title({"name": title})
        chart.set_x_axis({
            "name": x_axis_name,
            "name_font": {"size": 12, "bold": True},
            "num_font": {"rotation": -45},
        })
        chart.set_y_axis({
            "name": "Count",
            "name_font": {"size": 12, "bold": True},
        })
        chart.set_legend({"position": "right"})
        return chart

    def create_recipients_chart(self, data: List[Tuple[str, int]]) -> None:
        """Create chart for most recognized recipients."""
        chart = self._create_chart(
            data, "MRR", "Most Recognized Recipients", "Recipient"
        )
        self.summary_ws.insert_chart("D2", chart)

    def create_issuers_chart(self, data: List[Tuple[str, int]]) -> None:
        """Create chart for top issuing users."""
        chart = self._create_chart(
            data, "MRR_I", "Top Issuing Users", "Recipient"
        )
        self.summary_ws.insert_chart("D20", chart)

def main():
    recipients_data = [
        ("Namrata Puranik Puranik", 51),
        ("Gauri Puranik", 16),
        ("Aalkhimovich aalk", 16),
        ("Alumni User", 10),
        ("Ankur Tripathi", 7),
    ]
    
    issuers_data = [
        ('Gauri Puranik', 84),
        ('Namrata Puranik Puranik', 21)
    ]

    output_path = "static_chart_no_data.xlsx"
    with xlsxwriter.Workbook(output_path) as workbook:
        chart_creator = ChartCreator(workbook)
        chart_creator.create_recipients_chart(recipients_data)
        chart_creator.create_issuers_chart(issuers_data)

    print(f"Excel file created successfully: {output_path}")

if __name__ == "__main__":
    main()