using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.Data;
using CobranzaSP.Lógica;
using CobranzaSP.Modelos;

namespace CobranzaSP.Lógica
{
    internal class LogicaCompras
    {
        private CD_Conexion conexion = new CD_Conexion();
        SqlCommand comando = new SqlCommand();

        public string RegistroCompra(Compra nuevaCompra, string sp)
        {
            int respuesta = 0;
            string AccionRealizada;
            comando.Connection = conexion.AbrirConexion();
            comando.CommandText = sp;
            comando.CommandType = CommandType.StoredProcedure;

            //Preguntamos que accion realizaremos en la base de datos para posteriormente mostrale al usuario la accion que realizo
            AccionRealizada = (sp == "AgregarCompra") ? "agrego" : "modifico";

            comando.Parameters.AddWithValue("@Factura", nuevaCompra.Factura);
            comando.Parameters.AddWithValue("@Cantidad", nuevaCompra.Cantidad);
            comando.Parameters.AddWithValue("Fecha", nuevaCompra.Fecha);
            comando.Parameters.AddWithValue("IdProveedor", nuevaCompra.IdProveedor);

            respuesta = comando.ExecuteNonQuery();
            string Mensaje = (respuesta > 0) ? "La compra se " + AccionRealizada + " correctamente" : "Algo salio mal, no se " + AccionRealizada + " el registro";

            comando.Parameters.Clear();
            conexion.CerrarConexion();
            return Mensaje;
        }

        public void EliminarCompra(string Factura, string sp)
        {
            comando.Connection = conexion.AbrirConexion();
            comando.CommandText = sp;
            comando.CommandType = CommandType.StoredProcedure;
            comando.Parameters.AddWithValue("@Factura", Factura);
            comando.ExecuteNonQuery();
            comando.Parameters.Clear();

            conexion.CerrarConexion();
        }

        public void AgregarProveedor(string Proveedor)
        {
            comando.Connection = conexion.AbrirConexion();
            comando.CommandText = "AgregarProveedor";
            comando.CommandType = CommandType.StoredProcedure;
            comando.Parameters.AddWithValue("@Nombre", Proveedor);
            comando.ExecuteNonQuery();
            comando.Parameters.Clear();

            conexion.CerrarConexion();
        }
    }
}
