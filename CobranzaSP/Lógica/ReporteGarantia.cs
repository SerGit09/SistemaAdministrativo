using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CobranzaSP.Lógica
{
    internal class ReporteGarantia
    {
        public string Cartucho { get; set; }
        public string Marca { get; set; }
        public string Cliente { get; set; }

        public string Descripcion { get; set; }
        public int Cantidad { get; set; }
        public DateTime Fecha { get; set; }
    }
}
