using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CobranzaSP.Modelos
{
    public class Impresora
    {
        public int IdMarca { get; set; }

        public int IdModelo { get; set; }

        public string Serie { get; set; }

        public int Contador { get; set; }
        public double Valor { get; set; }

        public DateTime FechaVenta { get; set; }



    }
}
