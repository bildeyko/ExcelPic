using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Drawing;

namespace ExcelPic
{
    class Program
    {
        static void Main(string[] args)
        {
            if(args.Length == 0)
            {
                System.Console.WriteLine("Необходимо задать путь к файлу");
                return;
            }

            var path = args[0];
            Bitmap img = new Bitmap(path);
            DisplayInExcel(img);

        }

        static void DisplayInExcel(Bitmap img)
        {
            var excelApp = new Excel.Application();
            excelApp.Visible = true;
            excelApp.Workbooks.Add();
            Excel._Worksheet ws = (Excel.Worksheet)excelApp.ActiveSheet;

            int offset = 0;
            bool[] cols = new bool[img.Width * 3];
            bool[] rows = new bool[img.Height];
            for (var i = 1; i < img.Height; i++)
            {
                for (var j = 1; j < img.Width; j++)
                {
                    var pixel = img.GetPixel(j-1, i-1);

                    var colName = j + offset;
                    ws.Cells[i, colName].Interior.Color = ColorTranslator.ToOle(Color.FromArgb(pixel.R, 0, 0));                    
                    if (!cols[j+offset])
                    {
                        cols[j + offset] = true;
                        ws.Cells[i, colName].EntireColumn.ColumnWidth = 0.7;
                    }
                    offset++;

                    colName = j + offset;
                    ws.Cells[i, colName].Interior.Color = ColorTranslator.ToOle(Color.FromArgb(0, pixel.G, 0));
                    if (!cols[j + offset])
                    {
                        cols[j + offset] = true;
                        ws.Cells[i, colName].EntireColumn.ColumnWidth = 0.7;
                    }
                    offset++;

                    colName = j + offset;
                    ws.Cells[i, colName].Interior.Color = ColorTranslator.ToOle(Color.FromArgb(0, 0, pixel.B));
                    if (!cols[j + offset])
                    {
                        cols[j + offset] = true;
                        ws.Cells[i, colName].EntireColumn.ColumnWidth = 0.7;
                    }
                }
                offset = 0;
            }
        }
    }
}
