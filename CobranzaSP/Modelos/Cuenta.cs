using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CobranzaSP.Modelos
{
    internal class Cuenta
    {
        public int IdCliente { get; set; }

        public string Factura { get; set; }

        public DateTime FechaFactura { get; set; }

        public DateTime FechaPago { get; set; }

        public double Cantidad { get; set; }
    }
}
