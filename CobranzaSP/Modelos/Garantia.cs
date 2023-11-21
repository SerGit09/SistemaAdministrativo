using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CobranzaSP.Modelos
{
    public class Garantia : RegistroInventario
    {
        public int IdCliente { get; set; }

        public string Descripcion { get; set; }
        public int Cantidad { get; set; }

        public string Marca { get; set; }

        public string Modelo { get; set;}
    }
}
