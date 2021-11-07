using ImageToExcelEPPlus.Class;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ImageToExcelEPPlus.Metodos;


namespace ImageToExcelEPPlus
{
    class Program
    {
        static async Task Main(string[] args)
        {
            Console.WriteLine("Inicio de programa");
            Stopwatch stopwatch = new Stopwatch();
            Console.WriteLine("Timer iniciado");

            try
            {

                List<ColorPosicion> lstMatrix = new List<ColorPosicion>();
                string Adress = "";
                var file = new FileInfo(@"C:\Users\j053_\source\repos\j053pepe\ImageToExcel\ImageToExcelNPOI\myWorkbook3.xlsx");
                string pathImage = @"C:\Users\j053_\source\repos\j053pepe\ImageToExcel\ImageToExcelNPOI\HH&Tv2.jpg"; 
                if (File.Exists(@"C:\Users\j053_\source\repos\j053pepe\ImageToExcel\ImageToExcelNPOI\myWorkbook3.xlsx"))
                {
                    file.Delete();
                    Console.WriteLine("Archivo eliminado");
                }

                Bitmap bmp = new Bitmap(pathImage);
                stopwatch.Start();

                for (int x = 0; x < bmp.Size.Width; x++)
                {
                    for (int y = 0; y < bmp.Size.Height; y++)
                    //for (int y = 0; y < 20; y++)
                    {
                        try
                        {
                            //colors.Add(bmp.GetPixel(x, y));
                            Adress = "";
                            Adress = await Matriz.GetAdress(x) + (y + 1);

                            lstMatrix.Add(new ColorPosicion
                            {
                                ColorOriginal = bmp.GetPixel(x, y),
                                Posicion = Adress
                            });

                        }
                        catch (Exception errorimagen)
                        {
                            Console.WriteLine(errorimagen.Message);
                        }
                    }
                }


                using (var package = new ExcelPackage(file))
                {
                    var sheet = package.Workbook.Worksheets.Add("Imagen");
                    sheet.DefaultColWidth = 2.5;

                    lstMatrix.ForEach(item =>
                    {
                        sheet.Cells[item.Posicion].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        sheet.Cells[item.Posicion].Style.Fill.BackgroundColor.SetColor(item.ColorOriginal);
                        //using(var range = sheet.Cells[item.Posicion])
                        //{
                        //    range.Style.Fill.BackgroundColor.SetColor(item.ColorOriginal);
                        //}
                    });

                    // Save to file
                    package.Save();
                }

                Console.Beep();
                Console.WriteLine("Archivo creado");
                TimeSpan ts = stopwatch.Elapsed;
                Console.WriteLine($"Tiempo transcurrido: {ts.ToString("mm\\:ss\\.ff")}");
                Console.WriteLine("Presiona una tecla para salir.");
                Console.ReadLine();
            }
            catch (Exception errorgeneral)
            {
                Console.WriteLine(errorgeneral.Message);
                Console.WriteLine(errorgeneral.StackTrace);
                Console.Beep();
                Console.WriteLine("Ocurrio un error. Presiona una tecla para salir.");
                Console.ReadLine();
            }
        }

    }
}
