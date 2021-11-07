using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ImageToExcelEPPlus.Class
{
    public class ColorAgrupado
    {
        public Color ColorOriginal { get; set; }
        public List<string> Posiciones { get; set; }
    }

    public class ColorPosicion
    {
        public Color ColorOriginal { get; set; }
        public string Posicion { get; set; }
    }
}
