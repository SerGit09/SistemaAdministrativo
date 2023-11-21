using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CobranzaSP.Modelos
{
    internal class CuentaEstado
    {
        public string Folio { get; set; }
        public string Estado { get; set; }

        public DateTime FechaFactura { get; set; }
        public DateTime FechaVencimiento { get; set; }

        public double Saldo { get; set; }

        public int IdCliente { get; set; }
    }
}
