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
            // Directory path for input and output files.
            string dirPath = "E:/InputOutputDir/";

            // Load input XLS file inside the Aspose.Cells workbook object.
            Aspose.Cells.Workbook workbook = new Workbook(dirPath + "InputXLS.xls");

            // Save the workbook in output Tab Delimited format.
            workbook.Save(dirPath + "OutputTabDelimited.txt", SaveFormat.TabDelimited);
        }
    }
}
```
