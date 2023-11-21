using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CobranzaSP.Modelos
{
    internal class Fusor: Factura
    {
        public int IdFusor { get; set; }

        public string SerieO { get; set; }

        public string SerieS { get; set; }

        public string Estado { get; set; }
        public string Proveedor { get; set; }

        public int DiasGarantia { get; set; }

        public string Modelo { get; set; }

    }
}
