using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CobranzaSP.Modelos
{
    internal class ReporteHojaEntrega:Modulo
    {

        public string NumeroParte { get; set; }
        public int Cantidad { get; set; }

        public string Descripcion { get; set; }

        public string Estado { get; set; }

        public int Folio { get; set; }
    }
}
