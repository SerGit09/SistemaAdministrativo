using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CobranzaSP.Modelos
{
    public class Equipo:Impresora
    {
        public int IdEquipo { get; set; }

        public int IdCliente { get; set; }

        public string Ubicacion { get; set; }

        public int IdRenta { get; set; }

        public double Precio { get; set; }

        public string FechaPago { get; set; }

    }
}
