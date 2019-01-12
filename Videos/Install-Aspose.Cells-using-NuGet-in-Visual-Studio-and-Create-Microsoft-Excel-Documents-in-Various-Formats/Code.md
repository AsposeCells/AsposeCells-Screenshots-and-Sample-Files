# C#.NET

```js
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Aspose.Cells;

namespace SampleProject
{
    class Program
    {
        static void Main(string[] args)
        {
            // Directory path where output Excel files will be created.
            string dirPath = "E:/OutputDir/";

            // Create Aspose.Cells empty workbook object.
            Aspose.Cells.Workbook workbook = new Workbook();

            // Put some value in cell C4 of first worksheet.
            workbook.Worksheets[0].Cells["C4"].PutValue("Welcome to Aspose.Cells API to create and manipulate Excel files!");

            // Save the workbook in output XLS format.
            workbook.Save(dirPath + "1_OutputXLS.xls", SaveFormat.Excel97To2003);

            // Save the workbook in output XLSX format.
            workbook.Save(dirPath + "2_OutputXLSX.xlsx", SaveFormat.Xlsx);

            // Save the workbook in output XLSM format.
            workbook.Save(dirPath + "3_OutputXLSM.xlsm", SaveFormat.Xlsm);

            // Save the workbook in output XLSB format.
            workbook.Save(dirPath + "4_OutputXLSB.xlsb", SaveFormat.Xlsb);

            // Save the workbook in output PDF format.
            workbook.Save(dirPath + "5_OutputPDF.pdf", SaveFormat.Pdf);
        }
    }
}

```
