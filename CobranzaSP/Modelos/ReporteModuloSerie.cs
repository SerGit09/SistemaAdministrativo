using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CobranzaSP.Modelos
{
    internal class ReporteModuloSerie
    {
        public string Clave { get; set; }
        public string Modulo { get; set; }
        public DateTime Fecha { get; set; }
        public string Folio { get; set; }
        public string Serie { get; set; }
        public string Observacion { get; set; }
        public string Marca { get; set; }
        public string Modelo { get; set; }
        public string Cliente { get; set; }
        public int Contador { get; set; }
    }
}
