C#

// Directory path for input and output files.
String dirPath = "D:/Download/";
 
// Load the input Excel file containing the sample data.
Workbook wb = new Workbook(dirPath + "SampleExcelToTextWithTabs.xlsx");
 
// Initialize index. Index starts from 0 in Aspose.Cells. So...
// 1st sheet index will be 0.
// 2nd sheet index will be 1.
// 3rd sheet index will be 2.
int index = 0;
 
// Set the active sheet which you want to export to tab delimited format.
wb.Worksheets.ActiveSheetIndex = index;
 
// Save the workbook as tab delimited text file. Tabulation inside the cell values will be preserved.
TxtSaveOptions txtSaveOpts = new TxtSaveOptions(SaveFormat.TabDelimited);
wb.Save(dirPath + "OutputExcelToTextWithTabs.txt", txtSaveOpts);

// ------------------------------------------------------------------
// ******************************************************************
// ------------------------------------------------------------------

Java

// Directory path for input and output files.
String dirPath = "D:/Download/";
 
// Load the input Excel file containing the sample data.
Workbook wb = new Workbook(dirPath + "SampleExcelToTextWithTabs.xlsx");
 
// Initialize index. Index starts from 0 in Aspose.Cells. So...
// 1st sheet index will be 0.
// 2nd sheet index will be 1.
// 3rd sheet index will be 2.
int index = 0;
 
// Set the active sheet which you want to export to tab delimited format.
wb.getWorksheets().setActiveSheetIndex(index);
 
// Save the workbook as tab delimited text file. Tabulation inside the cell values will be preserved.
TxtSaveOptions txtSaveOpts = new TxtSaveOptions(SaveFormat.TAB_DELIMITED);
wb.save(dirPath + "OutputExcelToTextWithTabs.txt", txtSaveOpts);

