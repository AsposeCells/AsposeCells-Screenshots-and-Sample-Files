# C#
```js
// Directory path for input and output Excel files.
string dirPath = "D:/Download/";

// Load the input Excel file inside workbook object.
Aspose.Cells.Workbook wb = new Workbook(dirPath + "SampleFormatExcelCellNumbersToCurrency.xlsx");
            
// Access first worksheet.
Worksheet ws = wb.Worksheets[0];

// Format Cell G3 with Curreny > Dollar.
Cell cell = ws.Cells["G3"];
Style st = cell.GetStyle();
st.Custom = "\"$\"#,##0.00";
cell.SetStyle(st);

// Format Cell G4 with Curreny > Yaun.
cell = ws.Cells["G4"];
st = cell.GetStyle();
st.Custom = "[$¥-804]#,##0.00";
cell.SetStyle(st);

// Format Cell G5 with Curreny > Pound.
cell = ws.Cells["G5"];
st = cell.GetStyle();
st.Custom = "[$£-809]#,##0.00";
cell.SetStyle(st);

// Format Cell G6 with Curreny > Euro.
cell = ws.Cells["G6"];
st = cell.GetStyle();
st.Custom = "#,##0.00[$€-40B]";
cell.SetStyle(st);

// Save the workbook in XLSX format. 
// You can also save it to XLS or other formats.
wb.Save(dirPath + "OutputFormatExcelCellNumbersToCurrency.xlsx");
```