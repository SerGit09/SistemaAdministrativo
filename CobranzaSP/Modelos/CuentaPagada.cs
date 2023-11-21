using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CobranzaSP.Modelos
{
    internal class CuentaPagada:Cuenta
    {
        public int Id { get; set; }

        public DateTime FechaPago { get;set; }
    }
}
