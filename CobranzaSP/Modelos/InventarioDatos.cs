using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CobranzaSP.Modelos
{
    internal class InventarioDatos:Toner
    {
        public int Id { get; set; }

        public int CantidadOficina { get; set; }

        public DateTime Fecha { get; set; }

        public double Precio { get; set; }
    }
}
