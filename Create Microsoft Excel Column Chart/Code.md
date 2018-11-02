# C#
```js
// Directory path of input and output files.
string dirPath = "D:/Download/";

// Load source Excel file containing the chart data.
Workbook wb = new Workbook(dirPath + "sampleCreateMicrosoftExcelColumnChart.xlsx");

// Access first worksheet.
Worksheet ws = wb.Worksheets[0];

// Specify dimensions of the chart.
int upperLeftRow = 7;
int upperLeftColumn = 4;
int lowerRightRow = 24;
int lowerRightColumn = 13;

// Create Line chart with specified dimensions.
int idx = ws.Charts.Add(ChartType.Column, upperLeftRow, upperLeftColumn, lowerRightRow, lowerRightColumn);

// Access the Column chart.
Chart ch = ws.Charts[idx];

// Set the outline of chart area.
ch.ChartArea.Border.Color = Color.Black;
ch.ChartArea.Border.Weight = WeightType.SingleLine;

// Set the chart title, make it non-bold and set its font size.
ch.Title.Text = "Classification of Languages";
ch.Title.Font.IsBold = false;
ch.Title.Font.Size = 15;

// Add three vertical series in chart covering the range B2:D5.
ch.NSeries.Add("B2:D5", true);

// Set the category data covering the range A2:A5.
ch.NSeries.CategoryData = "A2:A5";

// Set the names of the chart series taken from cells.
ch.NSeries[0].Name = "=B1";
ch.NSeries[1].Name = "=C1";
ch.NSeries[2].Name = "=D1";

// Set the 1st series fill color.
ch.NSeries[0].Area.ForegroundColor = Color.FromArgb(74, 127, 176);
ch.NSeries[0].Area.Formatting = FormattingType.Custom;

// Set the 2nd series fill color.
ch.NSeries[1].Area.ForegroundColor = Color.FromArgb(91, 155, 213);
ch.NSeries[1].Area.Formatting = FormattingType.Custom;

// Set the 3rd series fill color.
ch.NSeries[2].Area.ForegroundColor = Color.FromArgb(173, 198, 229);
ch.NSeries[2].Area.Formatting = FormattingType.Custom;

// Set plot area formatting as none and hide its border.
ch.PlotArea.Area.FillFormat.FillType = FillType.None;
ch.PlotArea.Border.IsVisible = false;

// Set value axis major tick mark as none and hide axis line. 
// Also set the color of value axis major grid lines.
ch.ValueAxis.MajorTickMark = TickMarkType.None;
ch.ValueAxis.AxisLine.IsVisible = false;
ch.ValueAxis.MajorGridLines.Color = Color.FromArgb(217, 217, 217);

// Save the output Excel file in XLSX format.
wb.Save(dirPath + "outputCreateMicrosoftExcelColumnChart.xlsx", SaveFormat.Xlsx);
```

# Java
```js
// Directory path of input and output files.
String dirPath = "D:/Download/";

// Load source Excel file containing the chart data.
Workbook wb = new Workbook(dirPath + "sampleCreateMicrosoftExcelColumnChart.xlsx");

// Access first worksheet.
Worksheet ws = wb.getWorksheets().get(0);

// Specify dimensions of the chart.
int upperLeftRow = 7;
int upperLeftColumn = 4;
int lowerRightRow = 24;
int lowerRightColumn = 13;

// Create Column chart with specified dimensions.
int idx = ws.getCharts().add(ChartType.COLUMN, upperLeftRow, upperLeftColumn, lowerRightRow, lowerRightColumn);

// Access the Line chart.
Chart ch = ws.getCharts().get(idx);

// Set the outline of chart area.
ch.getChartArea().getBorder().setColor(Color.getBlack());
ch.getChartArea().getBorder().setWeight(WeightType.SINGLE_LINE);

// Set the chart title, make it non-bold and set its font size.
ch.getTitle().setText("Classification of Languages");
ch.getTitle().getFont().setBold(false);
ch.getTitle().getFont().setSize(15);

// Add three vertical series in chart covering the range B2:D5.
ch.getNSeries().add("B2:D5", true);

// Set the category data covering the range A2:A5.
ch.getNSeries().setCategoryData("A2:A5");

// Set the names of the chart series taken from cells.
ch.getNSeries().get(0).setName("=B1");
ch.getNSeries().get(1).setName("=C1");
ch.getNSeries().get(2).setName("=D1");

// Set the 1st series fill color.
ch.getNSeries().get(0).getArea().setForegroundColor(Color.fromArgb(74, 127, 176));
ch.getNSeries().get(0).getArea().setFormatting(FormattingType.CUSTOM);

// Set the 2nd series fill color.
ch.getNSeries().get(1).getArea().setForegroundColor(Color.fromArgb(91, 155, 213));
ch.getNSeries().get(1).getArea().setFormatting(FormattingType.CUSTOM);

// Set the 3rd series fill color.
ch.getNSeries().get(2).getArea().setForegroundColor(Color.fromArgb(173, 198, 229));
ch.getNSeries().get(2).getArea().setFormatting(FormattingType.CUSTOM);

// Set plot area formatting as none and hide its border.
ch.getPlotArea().getArea().getFillFormat().setFillType(FillType.NONE);
ch.getPlotArea().getBorder().setVisible(false);

// Set value axis major tick mark as none and hide axis line. 
// Also set the color of value axis major grid lines.
ch.getValueAxis().setMajorTickMark(TickMarkType.NONE);
ch.getValueAxis().getAxisLine().setVisible(false);
ch.getValueAxis().getMajorGridLines().setColor(Color.fromArgb(217, 217, 217));

// Save the output Excel file in XLSX format.
wb.save(dirPath + "outputCreateMicrosoftExcelColumnChart.xlsx", SaveFormat.XLSX);
```

# C++
```js
// Directory path of input and output files.
string dirPath = "D:/Download/";

// Load source Excel file containing the chart data.
Workbook wb = new Workbook(dirPath + "sampleCreateMicrosoftExcelColumnChart.xlsx");

// Access first worksheet.
Worksheet ws = wb.Worksheets[0];

// Specify dimensions of the chart.
int upperLeftRow = 7;
int upperLeftColumn = 4;
int lowerRightRow = 24;
int lowerRightColumn = 13;

// Create Line chart with specified dimensions.
int idx = ws.Charts.Add(ChartType.Column, upperLeftRow, upperLeftColumn, lowerRightRow, lowerRightColumn);

// Access the Line chart.
Chart ch = ws.Charts[idx];

ch.ChartArea.Border.Color = Color.Black;
ch.ChartArea.Border.Weight = WeightType.SingleLine;

// Set the chart title, make it non-bold and set its font size.
ch.Title.Text = "Classification of Languages";
ch.Title.Font.IsBold = false;
ch.Title.Font.Size = 15;

// Add three vertical series in chart covering the range B2:D5.
ch.NSeries.Add("B2:D5", true);

// Set the category data covering the range A2:A5.
ch.NSeries.CategoryData = "A2:A5";

// Set the names of the chart series taken from cells.
ch.NSeries[0].Name = "=B1";
ch.NSeries[1].Name = "=C1";
ch.NSeries[2].Name = "=D1";

// Set the 1st series fill color.
ch.NSeries[0].Area.ForegroundColor = Color.FromArgb(74, 127, 176);
ch.NSeries[0].Area.Formatting = FormattingType.Custom;

// Set the 2nd series fill color.
ch.NSeries[1].Area.ForegroundColor = Color.FromArgb(91, 155, 213);
ch.NSeries[1].Area.Formatting = FormattingType.Custom;

// Set the 3rd series fill color.
ch.NSeries[2].Area.ForegroundColor = Color.FromArgb(173, 198, 229);
ch.NSeries[2].Area.Formatting = FormattingType.Custom;

// Set plot area formatting as none and hide its border.
ch.PlotArea.Area.FillFormat.FillType = FillType.None;
ch.PlotArea.Border.IsVisible = false;

// Set value axis major tick mark as none and hide axis line. 
// Also set the color of value axis major grid lines.
ch.ValueAxis.MajorTickMark = TickMarkType.None;
ch.ValueAxis.AxisLine.IsVisible = false;
ch.ValueAxis.MajorGridLines.Color = Color.FromArgb(217, 217, 217);

// Save the output Excel file in XLSX format.
wb.Save(dirPath + "outputCreateMicrosoftExcelColumnChart.xlsx", SaveFormat.Xlsx);
```
