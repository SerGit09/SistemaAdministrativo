using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CobranzaSP.Modelos
{
    internal class ReporteRegistro
    {
        public string FechaActual { get; set; }

        public string Marca { get; set; }

        public string Modelo { get; set; }

        public int CantidadSalida { get; set; }

        public int CantidadEntrada { get; set; }
        public int CantidadGarantia { get; set; }

        public string Cliente_Proveedor { get; set; }

        public bool BusquedaCliente { get; set; }

        public string ClaveFusor { get; set; }

        public bool ClienteEspecifico { get; set; } 
    }
}
