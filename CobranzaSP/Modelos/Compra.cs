using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CobranzaSP.Modelos
{
    internal class Compra
    {
        public string Factura { get; set; }

        public double Cantidad { get; set;}

        public DateTime Fecha { get; set; }

        public int IdProveedor { get; set; }
    }
}
