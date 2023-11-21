using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CobranzaSP.Modelos
{
    internal class Contacto:Cliente
    {

        public string Correo { get; set; }

        public int DiasCredito { get; set; }

    }
}
