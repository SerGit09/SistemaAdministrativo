using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.SqlClient;
using CobranzaSP.Modelos;

namespace CobranzaSP.Lógica
{
    internal class VentaMostradorLogica
    {
        private CD_Conexion conexion = new CD_Conexion();
        SqlCommand comando = new SqlCommand();
        public string AgregarVentaMostrador(VentaMostrador nuevaVentaMostrador)
        {
            int respuesta = 0;
            string mensaje = "";
            comando.Connection = conexion.AbrirConexion();
            comando.CommandText = "AgregarVentaMostrador";
            comando.CommandType = CommandType.StoredProcedure;

            comando.Parameters.AddWithValue("@IdCliente", nuevaVentaMostrador.IdCliente);
            comando.Parameters.AddWithValue("@Cantidad", nuevaVentaMostrador.Cantidad);
            comando.Parameters.AddWithValue("@FechaPago", nuevaVentaMostrador.FechaPago);
            comando.Parameters.AddWithValue("@NumeroFolio", 0000);
            comando.Parameters.AddWithValue("@TipoPago", "MOSTRADOR");

            respuesta = comando.ExecuteNonQuery();
            mensaje = (respuesta > 0) ? "Venta mostrador agregada correctamente" : "Algo ha salido mal, no se añadio el registro";

            comando.Parameters.Clear();
            conexion.CerrarConexion();
            return mensaje;
        }
    }
}
