using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CobranzaSP.Modelos
{
    public class ReporteEquipo:Equipo
    {
        public string Cliente { get; set; }

        public string Marca { get; set; }

        public string Modelo { get; set; }

        public string TipoRenta { get; set; }
    }
}
