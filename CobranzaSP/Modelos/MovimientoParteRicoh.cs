using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CobranzaSP.Modelos
{
    public class MovimientoParteRicoh:ParteRicoh
    {
        public int IdTipoPersona { get; set; }
        public int IdRegistro { get; set; }

        public string TipoMovimiento { get; set; }
        public string ClienteProveedor { get; set; }

        public DateTime Fecha { get; set; }


    }
}
