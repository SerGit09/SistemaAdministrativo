using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CobranzaSP.Modelos
{
    internal class ReporteServicios:Impresora
    {
        public string NumeroFolio { get; set; }

        public string Cliente { get; set; }

        public string Marca { get; set; }

        public string Modelo { get; set; }

        public DateTime Fecha { get; set; }

        public string Tecnico { get; set; }

        public string Fusor { get; set; }

        public string FusorSaliente { get; set; }

        public string ServicioRealizado { get; set; }

        public string ReporteFallo { get; set; }
    }
}
