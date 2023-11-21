using CobranzaSP.Modelos;
using DocumentFormat.OpenXml.Office2010.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CobranzaSP.Lógica
{
    internal class LogicaServicio
    {
        private CD_Conexion conexion = new CD_Conexion();
        SqlCommand comando = new SqlCommand();
        public string RegistroServicio(Servicio nuevoServicio, string sp)
        {
            int respuesta = 0;
            string AccionRealizada;
            comando.Connection = conexion.AbrirConexion();
            comando.CommandText = sp;
            comando.CommandType = CommandType.StoredProcedure;

            //Preguntamos que accion realizaremos en la base de datos para posteriormente mostrale al usuario la accion que realizo
            AccionRealizada = (sp == "AgregarServicio") ? "agrego" : "modifico";

            comando.Parameters.AddWithValue("@NumeroFolio", nuevoServicio.NumeroFolio);
            comando.Parameters.AddWithValue("@IdCliente", nuevoServicio.IdCliente);
            comando.Parameters.AddWithValue("IdMarca", nuevoServicio.IdMarca);
            comando.Parameters.AddWithValue("IdModelo", nuevoServicio.IdModelo);
            comando.Parameters.AddWithValue("@Serie", nuevoServicio.Serie);
            comando.Parameters.AddWithValue("@Contador", nuevoServicio.Contador);
            comando.Parameters.AddWithValue("@Fecha", nuevoServicio.Fecha);
            comando.Parameters.AddWithValue("@Tecnico", nuevoServicio.Tecnico);
            comando.Parameters.AddWithValue("@Fusor", nuevoServicio.Fusor);
            comando.Parameters.AddWithValue("@FusorRetirado", nuevoServicio.FusorSaliente);
            comando.Parameters.AddWithValue("@ServicioRealizado", nuevoServicio.ServicioRealizado);
            comando.Parameters.AddWithValue("@ReporteFalla", nuevoServicio.ReporteFallo);

            respuesta = comando.ExecuteNonQuery();
            string Mensaje = (respuesta > 0) ? "Registro se " + AccionRealizada + " correctamente" : "Algo salio mal, no se " + AccionRealizada + " el registro";

            comando.Parameters.Clear();
            conexion.CerrarConexion();
            return Mensaje;
        }

        //Unicamente se creo este metodo eliminar en servicio debido a que su Id es de tipo varchar
        public void EliminarRegistro(string Id, string sp)
        {
            comando.Connection = conexion.AbrirConexion();
            comando.CommandText = sp;
            comando.CommandType = CommandType.StoredProcedure;
            comando.Parameters.AddWithValue("@Id", Id);
            comando.ExecuteNonQuery();
            comando.Parameters.Clear();

            conexion.CerrarConexion();
        }

        public void AñadirModelo(string Modelo, int IdMarca)
        {
            comando.Connection = conexion.AbrirConexion();
            comando.CommandText = "AñadirModelo";
            comando.CommandType = CommandType.StoredProcedure;
            comando.Parameters.AddWithValue("@Modelo", Modelo);
            comando.Parameters.AddWithValue("@IdMarca", IdMarca);
            comando.ExecuteNonQuery();
            comando.Parameters.Clear();

            conexion.CerrarConexion();
        }


        public int ObtenerMarcaFusor(int IdCartucho)
        {
            SqlDataReader leer;
            int IdMarca = 0;
            comando.Connection = conexion.AbrirConexion();
            comando.CommandText = "ObtenerMarcaFusor";
            comando.CommandType = CommandType.StoredProcedure;

            comando.Parameters.AddWithValue("@IdCartucho", IdCartucho);
            leer = comando.ExecuteReader();
            if (leer.Read())
            {
                IdMarca = int.Parse(leer[0].ToString());
            }
            comando.Parameters.Clear();
            leer.Close();
            conexion.CerrarConexion();
            return IdMarca;
        }
    }
}
