using CobranzaSP.Formularios;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CobranzaSP.Modelos
{
    internal class ReporteModulo:ReporteServicios
    {
        public int Paginas { get; set; }

        public string Clave { get; set; }

        public string Estado { get; set; }

        public string Modulo { get; set; }
    }
}
