using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CobranzaSP.Modelos
{
    public class Reporte
    {
        public DateTime FechaInicio { get; set; }

        public DateTime FechaFinal { get; set; }

        public string TipoBusqueda { get; set; }

        public string ParametroBusqueda { get; set; }
        

        public string TipoBusquedaAdicional { get; set; }

        public bool RangoHabilitado { get; set; }
    }
}
