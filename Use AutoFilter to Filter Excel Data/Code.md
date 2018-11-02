# C#
```js
// Directory path for input and output Excel files.
String dirPath = "D:/Download/";

// Load the input Excel file containing the sample data.
Workbook wb = new Workbook(dirPath + "sampleUseAutoFilterToFilterExcelData.xlsx");

// Access first worksheet.
Worksheet ws = wb.Worksheets[0];

// Apply auto filter to the range.
ws.AutoFilter.Range = "D3:G3";

// Add filter to first column (i.e. Vehicle) inside the range - Criteria --> Bike
ws.AutoFilter.AddFilter(0, "Bike");

// Add filter to first column (i.e. Vehicle) inside the range - Criteria --> Car
ws.AutoFilter.AddFilter(0, "Car");

// Refresh the auto filter.
ws.AutoFilter.Refresh();

// Add filter to second column (i.e. Color) inside the range - Criteria --> Green
ws.AutoFilter.AddFilter(1, "Green");

// Add filter to second column (i.e. Color) inside the range - Criteria --> Blue
ws.AutoFilter.AddFilter(1, "Blue");

// Refresh the auto filter.
ws.AutoFilter.Refresh();

// Save the workbook in XLSX format. 
// You can also save it to XLS or other formats.
wb.Save(dirPath + "outputUseAutoFilterToFilterExcelData.xlsx", SaveFormat.Xlsx);
```

# Java
```js
// Directory path for input and output Excel files.
String dirPath = "D:/Download/";

// Load the input Excel file containing the sample data.
Workbook wb = new Workbook(dirPath + "sampleUseAutoFilterToFilterExcelData.xlsx");

// Access first worksheet.
Worksheet ws = wb.getWorksheets().get(0);

// Apply auto filter to the range.
ws.getAutoFilter().setRange("D3:G3");

// Add filter to first column (i.e. Vehicle) inside the range - Criteria --> Bike
ws.getAutoFilter().addFilter(0, "Bike");

// Add filter to first column (i.e. Vehicle) inside the range - Criteria --> Car
ws.getAutoFilter().addFilter(0, "Car");

// Refresh the auto filter.
ws.getAutoFilter().refresh();

// Add filter to second column (i.e. Color) inside the range - Criteria --> Green
ws.getAutoFilter().addFilter(1, "Green");

// Add filter to second column (i.e. Color) inside the range - Criteria --> Blue
ws.getAutoFilter().addFilter(1, "Blue");

// Refresh the auto filter.
ws.getAutoFilter().refresh();

// Save the workbook in XLSX format. 
// You can also save it to XLS or other formats.
wb.save(dirPath + "outputUseAutoFilterToFilterExcelData.xlsx", SaveFormat.XLSX);
```

# C++
```js
// Path of input Excel file.
intrusive_ptr<Aspose::Cells::System::String> inputExcelFile = new Aspose::Cells::System::String("D:/Download/sampleUseAutoFilterToFilterExcelData.xlsx");

// Path of output Excel file.
intrusive_ptr<Aspose::Cells::System::String> outputExcelFile = new Aspose::Cells::System::String("D:/Download/outputUseAutoFilterToFilterExcelData.xlsx");

// Declaration of some variables to be used later.
intrusive_ptr<Aspose::Cells::System::String> strRng;
intrusive_ptr<Aspose::Cells::System::String> strCriteria;

// Load the input Excel file containing the sample data.
intrusive_ptr<Aspose::Cells::IWorkbook> wb = Aspose::Cells::Factory::CreateIWorkbook(inputExcelFile);

// Access first worksheet.
intrusive_ptr<Aspose::Cells::IWorksheet> ws = wb->GetIWorksheets()->GetObjectByIndex(0);

// Apply auto filter to the range.
strRng = new Aspose::Cells::System::String("D3:G3");
ws->GetIAutoFilter()->SetRange(strRng);

// Add filter to first column (i.e. Vehicle) inside the range - Criteria --> Bike
strCriteria = new Aspose::Cells::System::String("Bike");
ws->GetIAutoFilter()->AddFilter(0, strCriteria);

// Add filter to first column (i.e. Vehicle) inside the range - Criteria --> Car
strCriteria = new Aspose::Cells::System::String("Car");
ws->GetIAutoFilter()->AddFilter(0, strCriteria);

// Refresh the auto filter.
ws->GetIAutoFilter()->Refresh();

// Add filter to second column (i.e. Color) inside the range - Criteria --> Green
strCriteria = new Aspose::Cells::System::String("Green");
ws->GetIAutoFilter()->AddFilter(1, strCriteria);

// Add filter to second column (i.e. Color) inside the range - Criteria --> Blue
strCriteria = new Aspose::Cells::System::String("Blue");
ws->GetIAutoFilter()->AddFilter(1, strCriteria);

// Refresh the auto filter.
ws->GetIAutoFilter()->Refresh();

// Save the workbook in XLSX format. 
// You can also save it to XLS or other formats.
wb->Save(outputExcelFile, Aspose::Cells::SaveFormat::SaveFormat_Xlsx);
```
