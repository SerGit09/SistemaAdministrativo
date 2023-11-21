using CobranzaSP.Modelos;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CobranzaSP.Lógica
{
    internal class LogicaPartesUsadas
    {
        private CD_Conexion conexion = new CD_Conexion();
        SqlCommand comando = new SqlCommand();
        public string RegistroParte(Parte nuevaParte, string sp)
        {
            int respuesta = 0;
            int valor = 0;
            string AccionRealizada;
            string Mensaje;

            valor = ActualizarInventarioPartesRicoh(nuevaParte);

            if(valor > 0)
            {
                comando.Connection = conexion.AbrirConexion();
                comando.CommandText = sp;
                comando.CommandType = CommandType.StoredProcedure;
                comando.Parameters.Clear();
                //Preguntamos que accion realizaremos en la base de datos para posteriormente mostrale al usuario la accion que realizo
                AccionRealizada = (sp == "AgregarParteUsadaServicio") ? "agrego" : "modifico";

                if (nuevaParte.Id != 0)
                {
                    comando.Parameters.AddWithValue("@IdRegistro", nuevaParte.Id);
                }
                comando.Parameters.AddWithValue("@NumeroFolio", nuevaParte.NumeroFolio);
                comando.Parameters.AddWithValue("@IdNumeroParte", nuevaParte.IdNumeroParte);
                comando.Parameters.AddWithValue("@Cantidad", nuevaParte.Cantidad);
                comando.Parameters.AddWithValue("@Descripcion", nuevaParte.Descripcion);
                comando.Parameters.AddWithValue("@Estado", nuevaParte.Estado);

                respuesta = comando.ExecuteNonQuery();
                Mensaje = (respuesta > 0) ? "Registro se " + AccionRealizada + " correctamente" : "Algo salio mal, no se " + AccionRealizada + " el registro";

                comando.Parameters.Clear();
                conexion.CerrarConexion();
            }
            else
            {
                Mensaje = "La cantidad excede la cantidad en inventario";
            }
            

            return Mensaje;
        }

        public void EliminarRegistroParteUsada(Parte parte)
        {
            comando.Connection = conexion.AbrirConexion();
            comando.CommandText = "EliminarParteUsadaServicio";
            comando.CommandType = CommandType.StoredProcedure;
            comando.Parameters.Clear();
            comando.Parameters.AddWithValue("@Id", parte.Id);
            comando.Parameters.AddWithValue("@IdNumeroParte", parte.IdNumeroParte);
            comando.Parameters.AddWithValue("@Cantidad", parte.Cantidad);
            comando.Parameters.AddWithValue("@Estado", parte.Estado);
            comando.ExecuteNonQuery();

            conexion.CerrarConexion();
        }

        public int ActualizarInventarioPartesRicoh(Parte nuevaParte)
        {
            int respuesta = 0;
            comando.Connection = conexion.AbrirConexion();
            comando.CommandText = "RestarInventarioRicoh";
            comando.CommandType = CommandType.StoredProcedure;
            comando.Parameters.Clear();

            //Primero deberemos de comprobar que se pueda realizar la resta del inventario de las partes de ricoh
            comando.Parameters.AddWithValue("@IdNumeroParte", nuevaParte.IdNumeroParte);
            comando.Parameters.AddWithValue("@Estado", nuevaParte.Estado);
            comando.Parameters.AddWithValue("@Cantidad", nuevaParte.Cantidad);

            respuesta = comando.ExecuteNonQuery();

            comando.Parameters.Clear();
            conexion.CerrarConexion();
            return respuesta;
        }

        public DataTable MostrarPartes(string NumeroFolio)
        {
            DataTable tabla = new DataTable();
            SqlDataReader leer;
            comando.Connection = conexion.AbrirConexion();
            comando.CommandText = "MostrarRegistrosPartesUsadas";
            comando.CommandType = CommandType.StoredProcedure;
            comando.Parameters.Clear();
            comando.Parameters.AddWithValue("@NumeroFolio", NumeroFolio);
            leer = comando.ExecuteReader();
            tabla.Load(leer);
            conexion.CerrarConexion();
            leer.Close();
            return tabla;
        }
    }
}
