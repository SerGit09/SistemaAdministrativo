using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CobranzaSP.Modelos
{
    public class RegistroInventario:Toner
    {
        public int Id { get; set; }
        public int CantidadSalida { get; set; }

        public int CantidadEntrada { get; set; }
        public int CantidadGarantia { get; set; }

        public string Cliente { get; set; }

        public DateTime Fecha { get; set; }


        public string NumeroSerie { get; set; }
    }
}
