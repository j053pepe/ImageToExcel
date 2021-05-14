using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ImageToExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

                if (xlApp == null)
                {
                    Console.WriteLine("Excel is not properly installed!!");
                    return;
                }

                Stopwatch stopwatch = new Stopwatch();
                Console.WriteLine("Timer iniciado");
                stopwatch.Start();
                string pathImage = @"C:\Users\j053_\Documents\Visual Studio 2019\Projects\ImageToExcel\ImageToExcel\test.jpg";
                //string pathImage = @"E:\Imágenes\test2.png";
                Excel.Workbook xlWorkBook;
                Excel.Worksheet xlWorkSheet;
                object misValue = System.Reflection.Missing.Value;
                string Adress = "";

                xlWorkBook = xlApp.Workbooks.Add(misValue);
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);


                HashSet<Color> colors = new HashSet<Color>();

                Bitmap bmp = new Bitmap(pathImage);

                for (int x = 0; x < bmp.Size.Width; x++)
                {
                    for (int y = 0; y < bmp.Size.Height; y++)
                    {
                        try
                        {
                            //colors.Add(bmp.GetPixel(x, y));
                            Adress = "";
                            Adress = GetAdress(x) + (y + 1);

                            Excel.Range formatRange;
                            formatRange = xlWorkSheet.get_Range(Adress, Adress);
                            formatRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(bmp.GetPixel(x, y));
                            formatRange.ColumnWidth = 2.5;
                            //xlWorkSheet.Cells[1, 2] = "Red";
                        }
                        catch (Exception errorimagen)
                        {
                            Console.WriteLine(errorimagen.Message);
                        }
                    }
                }



                //xlWorkBook.SaveAs("E:\\csharp-Excel.xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                //Here saving the file in xlsx

                xlWorkBook.SaveAs("C:\\Users\\j053_\\Documents\\Visual Studio 2019\\Projects\\ImageToExcel\\ImageToExcel\\csharp-Excel.xlsx", Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook, misValue,
                misValue, misValue, misValue, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);

                xlWorkBook.Close(true, misValue, misValue);
                xlApp.Quit();

                Marshal.ReleaseComObject(xlWorkSheet);
                Marshal.ReleaseComObject(xlWorkBook);
                Marshal.ReleaseComObject(xlApp);

                Console.WriteLine("Excel file created , you can find the file E:\\csharp-Excel.xls");
                TimeSpan ts = stopwatch.Elapsed;
                Console.WriteLine($"Tiempo transcurrido: {ts.ToString("mm\\:ss\\.ff")}");
                Console.WriteLine("Presiona una tecla para salir.");
                Console.ReadLine();
            }
            catch (Exception errorgeneral)
            {
                Console.WriteLine(errorgeneral.Message);
                Console.WriteLine(errorgeneral.StackTrace);
            }

        }

        public static string GetAdress(int x)
        {
            var listAbc = new List<string>() { "a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", "y", "z" };

           if (x < 26)
            {
                return listAbc[x];
            }
            else
            {
                int divicion = x / 26;
                int residuo = x % 26;
                double decimales = x / 26;
                string decimString = decimales.ToString();

                divicion = decimString.Contains('.') ? divicion : divicion - 1;
                string adress = "";

                if (divicion > 26)
                {
                    adress += GetAdress(divicion);
                }

                adress += listAbc[divicion] + listAbc[residuo>0? residuo - 1: residuo];

                return adress;
            }
        }
    }
}
