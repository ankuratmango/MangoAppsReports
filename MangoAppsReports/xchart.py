import xlsxwriter

workbook = xlsxwriter.Workbook("chart_data_labels.xlsx")
worksheet = workbook.add_worksheet()
bold = workbook.add_format({"bold": 1})

# Add the worksheet data that the charts will refer to.
headings = ["Number", "Data", "Text"]

data = [
    [2, 3, 4, 5, 6, 7],
    [20, 10, 20, 30, 40, 30],
]

worksheet.write_row("A1", headings, bold)
worksheet.write_column("A2", data[0])
worksheet.write_column("B2", data[1])

#######################################################################
#
# Example with customized bar colors.
#

# Create a Column chart.
chart1 = workbook.add_chart({"type": "column"})

# Configure the data series and add the data labels, with different colors for each bar.
chart1.add_series(
    {
        "categories": "=Sheet1!$A$2:$A$7",
        "values": "=Sheet1!$B$2:$B$7",
        "data_labels": {"value": True},
        "points": [
            {"fill": {"color": "red"}},
            {"fill": {"color": "blue"}},
            {"fill": {"color": "green"}},
            {"fill": {"color": "yellow"}},
            {"fill": {"color": "purple"}},
            {"fill": {"color": "orange"}},
        ],
    }
)

# Add a chart title.
chart1.set_title({"name": "Chart with standard data labels"})

chart1.set_title({"name": "Most Recognized Recipients"})
chart1.set_x_axis({
    "name": "Recipient",
    "name_font": {"size": 12, "bold": True},
    "num_font": {"rotation": -45},  # Rotate X-axis labels to fit long names
})
chart1.set_y_axis({
    "name": "Count",
    "name_font": {"size": 12, "bold": True},
})
chart1.set_legend({"position": "right"})

# Turn off the chart legend.
chart1.set_legend({"none": True})

# Insert the chart into the worksheet (with an offset).
worksheet.insert_chart("D2", chart1, {"x_offset": 25, "y_offset": 10})

workbook.close()
