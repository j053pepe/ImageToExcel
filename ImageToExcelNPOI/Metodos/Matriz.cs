using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ImageToExcelEPPlus.Metodos
{
    public class Matriz
    {

        public static async Task<string> GetAdress(int x)
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
                double decimales = x / 26.00;
                //string decimString = decimales.ToString();

                //divicion = decimString.Contains('.') ? divicion : divicion - 1;
                string adress = "";

                if (divicion > 26)
                {
                    adress += await GetAdress(divicion);
                }

                adress += listAbc[divicion > 0 ? (residuo == 0 && divicion >= 2 ? divicion - 2 : divicion - 1) : divicion] + listAbc[residuo > 0 ? residuo - 1 : (residuo == 0 ? 25 : residuo)];

                /*
                if (adress.Contains("ay"))
                {
                    Console.WriteLine("debug");
                }*/

                return adress;
            }
        }
    }
}
