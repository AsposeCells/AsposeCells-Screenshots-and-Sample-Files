# C#
```js
// The following sample code uses Aspose.Cells API to ... 
// Generate Excel Cell Data with Formatting and Save to XLSX, PDF and HTML formats.

//Create Empty Workbook.
Workbook wb = new Workbook();

//Access First Worksheet.
Worksheet ws = wb.Worksheets[0];

//Add Month Data by Rows and Column Indices.
ws.Cells[0, 0].PutValue("Month");
ws.Cells[1, 0].PutValue("Jan");
ws.Cells[2, 0].PutValue("Feb");
ws.Cells[3, 0].PutValue("Mar");
ws.Cells[4, 0].PutValue("Apr");

//Add Sales Data by Cell Names.
ws.Cells["B1"].PutValue("Sales");
ws.Cells["B2"].PutValue(600);
ws.Cells["B3"].PutValue(700);
ws.Cells["B4"].PutValue(900);
ws.Cells["B5"].PutValue(100);

//Add Expenses Data by Cell Names.
ws.Cells["C1"].PutValue("Expenses");
ws.Cells["C2"].PutValue(800);
ws.Cells["C3"].PutValue(400);
ws.Cells["C4"].PutValue(600);
ws.Cells["C5"].PutValue(200);

//Set the Fill Color and Bold Font of Cells A1, B1 and C1.
Style st = ws.Cells["A1"].GetStyle();
st.Font.IsBold = true;
st.Pattern = BackgroundType.Solid;
st.ForegroundColor = Color.Yellow;
ws.Cells["A1"].SetStyle(st, true);
ws.Cells["B1"].SetStyle(st, true);
ws.Cells["C1"].SetStyle(st, true);

//Center Align Range - Vertically and Horizontally and Set Borders.
Range rng1 = ws.Cells.CreateRange("A1:C5");
StyleFlag flag1 = new StyleFlag();
flag1.HorizontalAlignment = true;
flag1.VerticalAlignment = true;
flag1.Borders = true;
Style st1 = wb.CreateStyle();
st1.HorizontalAlignment = TextAlignmentType.Center;
st1.VerticalAlignment = TextAlignmentType.Center;
st1.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
st1.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
st1.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
st1.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;
rng1.ApplyStyle(st1, flag1);

//Set the Currency format of Range
Range rng2 = ws.Cells.CreateRange("B2:C5");
StyleFlag flag2 = new StyleFlag();
flag2.NumberFormat = true;
Style st2 = wb.CreateStyle();
st2.Number = 5;
rng2.ApplyStyle(st2, flag2);

//Add Column Chart.
int idx = ws.Charts.Add(ChartType.Column, 10, 1, 30, 8);

//Access Chart.
Chart ch = ws.Charts[idx];

//Add Two Vertical Series
ch.NSeries.Add("B2:B5", true);
ch.NSeries.Add("C2:C5", true);

//Set the Category Data.
ch.NSeries.CategoryData = "A2:A5";

//Set the Series Names.
ch.NSeries[0].Name = "=B1";
ch.NSeries[1].Name = "=C1";

//Set the Chart Title.
ch.Title.Text = "Sales and Expenses by Months";

//Calculate Chart Items.
ch.Calculate();

//Save Workbook in Xlsx format.
String dirPath = "D:\\Download\\"; 
wb.Save(dirPath + "output.xlsx", SaveFormat.Xlsx);

//Save Workbook in PDF format.
wb.Save(dirPath + "output.pdf", SaveFormat.Pdf);
                
//Save Workbook in HTML format.                
wb.Save(dirPath + "output.html", SaveFormat.Html);
```