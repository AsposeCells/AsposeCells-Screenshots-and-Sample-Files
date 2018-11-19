# C#
```js
// Directory path for input Excel file.
string dirPath = "D:/Download/";

// Load the input Excel file inside workbook object.
Aspose.Cells.Workbook wb = new Workbook(dirPath + "SampleExcelColor.xlsx");

// Access first worksheet.
Worksheet ws = wb.Worksheets[0];

// Access cell C4 by name.
Cell cell = ws.Cells["C4"];

// Access cell style.
Style st = cell.GetStyle();

// Print fill color of the cell i.e. Yellow.
// Please note, Yellow is (R=255, G=255, B=0)
Console.WriteLine(st.ForegroundColor);

// Print font color of the cell i.e. Red.
// Please note, Red is (R=255, G=0, B=0)
Console.WriteLine(st.Font.Color);
```

# Java
```js
// Directory path for input Excel file.
String dirPath = "D:/Download/";

// Load the input Excel file inside workbook object.
com.aspose.cells.Workbook wb = new Workbook(dirPath + "SampleExcelColor.xlsx");

// Access first worksheet.
Worksheet ws = wb.getWorksheets().get(0);

// Access cell C4 by name.
Cell cell = ws.getCells().get("C4");

// Access cell style.
Style st = cell.getStyle();

// Print fill color of the cell i.e. Yellow.
// Please note, Yellow is (R=255, G=255, B=0)
System.out.println(st.getForegroundColor());

// Print font color of the cell i.e. Red.
// Please note, Red is (R=255, G=0, B=0)
System.out.println(st.getFont().getColor());
```

**Console Output - C#**

Color [A=255, R=255, G=255, B=0]
Color [A=255, R=255, G=0, B=0]

**Console Output - Java**

com.aspose.cells.Color@ffffff00
com.aspose.cells.Color@ffff0000