using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CobranzaSP.Modelos
{
    internal class ReporteEquipoBodega:Impresora
    {

        public string Marca { get; set; }

        public string Modelo { get; set; }

        public string Estado { get; set; }
        public string Notas { get; set; }

        public string Precio { get; set; }

    }
}
