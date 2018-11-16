# C#
```js
// Directory path of input and output files.
string dirPath = "D:/Download/";

// Specify load options Excel97To2003 i.e. XLS format. 
LoadOptions opts = new LoadOptions(LoadFormat.Excel97To2003);
            
// Load the input XLS file inside the Aspose.Cells workbook object.
Workbook wb = new Workbook(dirPath + "SampleConvertMicrosoftExcelXLSToXLSX.xls", opts);

// Save the workbook as output XLSX file.
wb.Save(dirPath + "OutputConvertMicrosoftExcelXLSToXLSX.xlsx", SaveFormat.Xlsx);
```

# Java
```js
// Directory path of input and output files.
String dirPath = "D:/Download/";

// Specify load options Excel97To2003 i.e. XLS format. 
LoadOptions opts = new LoadOptions(LoadFormat.EXCEL_97_TO_2003);

// Load the input XLS file inside the Aspose.Cells workbook object.
Workbook wb = new Workbook(dirPath + "SampleConvertMicrosoftExcelXLSToXLSX.xls", opts);

// Save the workbook as output XLSX file.
wb.save(dirPath + "OutputConvertMicrosoftExcelXLSToXLSX.xlsx", SaveFormat.XLSX);
```
