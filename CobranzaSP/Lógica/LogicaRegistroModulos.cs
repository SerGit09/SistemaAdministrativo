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
    internal class LogicaRegistroModulos
    {
        private CD_Conexion conexion = new CD_Conexion();
        SqlCommand comando = new SqlCommand();

        public string RegistroInsumo(Modulo nuevoModulo, string sp)
        {
            int respuesta = 0;
            string AccionRealizada;
            comando.Connection = conexion.AbrirConexion();
            comando.CommandText = sp;
            comando.CommandType = CommandType.StoredProcedure;
            comando.Parameters.Clear();

            //Preguntamos que accion realizaremos en la base de datos para posteriormente mostrale al usuario la accion que realizo
            AccionRealizada = (sp == "AgregarRegistroInsumo") ? "agrego" : "modifico";

            comando.Parameters.AddWithValue("@Folio", nuevoModulo.Folio);
            if(nuevoModulo.IdRegistro != 0)
            {
                comando.Parameters.AddWithValue("@IdRegistro", nuevoModulo.IdRegistro);
            }
            comando.Parameters.AddWithValue("@Fecha", nuevoModulo.Fecha); 
            comando.Parameters.AddWithValue("@Recibido", nuevoModulo.Recibido);
            comando.Parameters.AddWithValue("@Concepto", nuevoModulo.Concepto);


            respuesta = comando.ExecuteNonQuery();
            string Mensaje = (respuesta > 0) ? "Registro se " + AccionRealizada + " correctamente" : "Algo salio mal, no se " + AccionRealizada + " el registro";

            comando.Parameters.Clear();
            conexion.CerrarConexion();
            return Mensaje;
        }

        public string RegistroParte(Parte nuevaParte, string sp)
        {
            int respuesta = 0;
            int valor = 0;
            string AccionRealizada;
            string Mensaje;
            comando.Connection = conexion.AbrirConexion();
            comando.CommandText = sp;
            comando.CommandType = CommandType.StoredProcedure;
            comando.Parameters.Clear();
            //Preguntamos que accion realizaremos en la base de datos para posteriormente mostrale al usuario la accion que realizo
            AccionRealizada = (sp == "AgregarRegistroParte") ? "agrego" : "modifico";

            if(nuevaParte.Id > 0)
            {
                comando.Parameters.AddWithValue("@IdRegistro", nuevaParte.Id);
            }
            else
            {
                comando.Parameters.AddWithValue("@Folio", nuevaParte.Folio);
            }

            comando.Parameters.AddWithValue("@Cantidad", nuevaParte.Cantidad);
            comando.Parameters.AddWithValue("@Descripcion", nuevaParte.Descripcion);
            comando.Parameters.AddWithValue("@Estado", nuevaParte.Estado);
            comando.Parameters.AddWithValue("@IdNumeroParte", nuevaParte.IdNumeroParte);


            respuesta = comando.ExecuteNonQuery();
            Mensaje = (respuesta > 0) ? "Registro se " + AccionRealizada + " correctamente" : "Algo salio mal, no se " + AccionRealizada + " el registro";

            comando.Parameters.Clear();
            conexion.CerrarConexion();

            return Mensaje;
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


        public DataTable MostrarPartes(int NumeroFolio)
        {
            DataTable tabla = new DataTable();
            SqlDataReader leer;
            comando.Connection = conexion.AbrirConexion();
            comando.CommandText = "MostrarRegistrosPartes";
            comando.CommandType = CommandType.StoredProcedure;
            comando.Parameters.Clear();
            comando.Parameters.AddWithValue("@Folio", NumeroFolio);
            leer = comando.ExecuteReader();
            tabla.Load(leer);
            conexion.CerrarConexion();
            leer.Close();
            return tabla;
        }

        public bool VerificarExistenciaFolio(int ParametroBusqueda, string sp)
        {
            SqlDataReader leer;
            bool Encontrado = false;
            comando.Connection = conexion.AbrirConexion();
            comando.CommandText = sp;
            comando.CommandType = CommandType.StoredProcedure;
            comando.Parameters.Clear();
            comando.Parameters.AddWithValue("@ParametroBusqueda", ParametroBusqueda);
            leer = comando.ExecuteReader();
            comando.Parameters.Clear();

            Encontrado = (leer.Read()) ? true : false;
            leer.Close();
            return Encontrado;
        }

        public string ObtenerDescripcionParteRicoh(string NumeroParte)
        {
            SqlDataReader leer;
            string Descripcion = "";
            comando.Connection = conexion.AbrirConexion();
            comando.CommandText = "SeleccionarDescripcionParte";
            comando.CommandType = CommandType.StoredProcedure;
            comando.Parameters.Clear();


            comando.Parameters.AddWithValue("@NumeroParte", NumeroParte);
            leer = comando.ExecuteReader();
            while (leer.Read())
            {
                Descripcion = leer[0].ToString();
            }
            comando.Parameters.Clear();
            conexion.CerrarConexion();
            leer.Close();
            return Descripcion;
        }
    }
}
