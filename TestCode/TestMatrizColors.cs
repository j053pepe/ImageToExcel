using ImageToExcelEPPlus.Class;
using ImageToExcelEPPlus.Metodos;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Threading.Tasks;

namespace TestCode
{
    [TestClass]
    public class TestMatrizColors
    {
        [TestMethod]
        public async Task FirstListColors()
        {
            string pathImage = @"C:\Users\j053_\Source\Repos\j053pepe\ImageToExcel\ImageToExcel\test.jpg";
            string Adress = "";
            Bitmap bmp = new Bitmap(pathImage);
            List<ColorAgrupado> lstMatrix = new List<ColorAgrupado>();

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

                        int index = lstMatrix.FindIndex(item => item.ColorOriginal == bmp.GetPixel(x, y));

                        if (index != -1)
                        {
                            lstMatrix[index].Posiciones.Add(Adress);
                        }
                        else
                        {
                            lstMatrix.Add(new ColorAgrupado
                            {
                                ColorOriginal = bmp.GetPixel(x, y),
                                Posiciones = new List<string> { Adress }
                            });
                        }
                    }
                    catch (Exception errorimagen)
                    {
                        Console.WriteLine(errorimagen.Message);
                    }
                }
            }

            Console.WriteLine($"Items en lista: {lstMatrix.Count}");
        }

        [TestMethod]
        public async Task CloneMatriz()
        {
            string pathImage = @"C:\Users\j053_\Source\Repos\j053pepe\ImageToExcel\ImageToExcel\test.jpg";
            string Adress = "";
            Bitmap bmp = new Bitmap(pathImage);
            List<ColorPosicion> lstMatrix = new List<ColorPosicion>();

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

            Console.WriteLine($"Items en lista: {lstMatrix.Count}");
        }

    }
}
