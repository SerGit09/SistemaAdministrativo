using CobranzaSP.Modelos;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CobranzaSP.Lógica
{
    internal class LogicaModulosCliente
    {
        CD_Conexion conexion = new CD_Conexion();
        SqlCommand comando = new SqlCommand();

        
        //public static List<string> Claves { get; set; } = new List<string>();
        public static List<ModuloEquipo> ModulosEquipo { get; set; } = new List<ModuloEquipo>();
        //Lista que nos servira para las claves a las cuales se le modifico su observacion
        public static List<string> ClavesObservaciones = new List<string>();
        LogicaModulosBodega lgModuloBodega = new LogicaModulosBodega();
        LogicaReportesModulo lgReporteModulo = new LogicaReportesModulo();
        public string RegistrarModulo(Modulo_Cliente NuevoModulo, string sp)
        {
            string AccionRealizada = "agrego";
            int respuesta;
            comando.Connection = conexion.AbrirConexion();
            comando.CommandText = sp;
            comando.CommandType = CommandType.StoredProcedure;
            comando.Parameters.Clear();
            if (NuevoModulo.Id > 0)
            {
                AccionRealizada = "modifico";
                comando.Parameters.AddWithValue("@Id", NuevoModulo.Id);
            }
            comando.Parameters.AddWithValue("@IdModulo", NuevoModulo.IdModulo);
            comando.Parameters.AddWithValue("@Clave", NuevoModulo.Clave);
            comando.Parameters.AddWithValue("@Folio", NuevoModulo.Reporte);
            comando.Parameters.AddWithValue("@Contador", NuevoModulo.Paginas);
            comando.Parameters.AddWithValue("@Estado", NuevoModulo.Estado);
            comando.Parameters.AddWithValue("@Serie", NuevoModulo.Serie);
            comando.Parameters.AddWithValue("@Observacion", NuevoModulo.Observacion);

            respuesta = comando.ExecuteNonQuery();
            string Mensaje = (respuesta > 0) ? "Se " + AccionRealizada + " el modulo correctamente" : "Algo salio mal, no se " + AccionRealizada + " el modulo";


            comando.Parameters.Clear();
            conexion.CerrarConexion();
            return Mensaje;
        }

        public void RetirarModulo(Modulo_Cliente NuevoModulo, bool ModuloSeleccionado)
        {
            int respuesta;
            comando.Connection = conexion.AbrirConexion();
            comando.CommandText = "AgregarEstadoModuloEquipo";
            comando.CommandType = CommandType.StoredProcedure;
            comando.Parameters.Clear();
            comando.Parameters.AddWithValue("@IdModulo", NuevoModulo.IdModulo);
            comando.Parameters.AddWithValue("@Clave", NuevoModulo.ClaveAnterior);
            comando.Parameters.AddWithValue("@Folio", NuevoModulo.Reporte);
            comando.Parameters.AddWithValue("@Contador", NuevoModulo.Paginas);
            comando.Parameters.AddWithValue("@Estado", "RETIRADO");
            comando.Parameters.AddWithValue("@Serie", NuevoModulo.Serie);
            comando.Parameters.AddWithValue("@Observacion", NuevoModulo.Observacion);


            respuesta = comando.ExecuteNonQuery();
            comando.Parameters.Clear();
            conexion.CerrarConexion();

            //Una vez que acabemos de retirar el modulo, sera necesario borrar del inventario de bodega el 
            //Modulo que se dejo instalado, y agregar el que se retiro
            //DEBEMOS DE VALIDAR CUANDO SEA UN RETIRO CON UN MODULO SELECCIONADO
            if (ModuloSeleccionado)
            {
                //Deberemos de agregar en el inventario la que retiramos y eliminar la que dejamos instalada
                ModificarInventarioModulosBodega(NuevoModulo);
            }
            else//Si no se selecciono un modulo
            {
                //Solamente agregaremos la que se retiro
                RegistrarModuloABodega(NuevoModulo);
            }
        }

        public string ObtenerSerieClaveRepetida(string Clave)
        {
            SqlDataReader ver;
            string Serie = "";
            string Equipo = "";
            comando.Connection = conexion.AbrirConexion();
            comando.CommandText = "ObtenerSerieClaveAsignada";
            comando.CommandType = CommandType.StoredProcedure;
            comando.Parameters.Clear();
            comando.Parameters.AddWithValue("@Clave", Clave);
            ver = comando.ExecuteReader();

            while (ver.Read())
            {
                Serie = ver[0].ToString();
            }

            conexion.CerrarConexion();
            ver.Close();
            Equipo = lgReporteModulo.ColocarMarcaModeloEquipo(Serie) + " CON LA SERIE: " +Serie;
            return Equipo;
        }

        public void RegistrarModuloABodega(Modulo_Cliente NuevoModulo)
        {
            int respuesta;
            comando.Connection = conexion.AbrirConexion();
            comando.CommandText = "AgregarModuloBodega";
            comando.CommandType = CommandType.StoredProcedure;
            comando.Parameters.Clear();
            comando.Parameters.AddWithValue("@IdModulo", NuevoModulo.IdModulo);
            comando.Parameters.AddWithValue("@Clave", NuevoModulo.ClaveAnterior);

            comando.ExecuteNonQuery();
            conexion.CerrarConexion();
        }

        //Quizas una funcion solo para agregar el modelo que fue retirado en bodega de modulos

        //Metodo que nos ayuda a agregar el modulo que fue retirado al inventario de modulos
        //Y retirar el que se le coloco ahora al equipo
        public void ModificarInventarioModulosBodega(Modulo_Cliente NuevoModulo)
        {
            comando.Connection = conexion.AbrirConexion();
            comando.CommandText = "AgregarYEliminarModuloBodega";
            comando.CommandType = CommandType.StoredProcedure;
            comando.Parameters.Clear();
            comando.Parameters.AddWithValue("@IdModulo", NuevoModulo.IdModulo);
            comando.Parameters.AddWithValue("@Clave", NuevoModulo.Clave);
            comando.Parameters.AddWithValue("@ClaveAnterior", NuevoModulo.ClaveAnterior);

            comando.ExecuteNonQuery();
            comando.Parameters.Clear();
            conexion.CerrarConexion();
        }

        public DataTable MostrarModulosCliente(string Serie, string Folio)
        {
            DataTable tabla = new DataTable();
            SqlDataReader leer;
            comando.Connection = conexion.AbrirConexion();
            //comando.CommandText = "MostrarModulosCliente";
            comando.CommandText = "MostrarModulosEquipo";
            comando.CommandType = CommandType.StoredProcedure;
            comando.Parameters.Clear();
            comando.Parameters.AddWithValue("@Serie", Serie);
            comando.Parameters.AddWithValue("@NumeroFolio", Folio);

            leer = comando.ExecuteReader();
            tabla.Load(leer);
            conexion.CerrarConexion();
            return tabla;
        }

        public int BuscarIdModulo(string campo, int IdModelo)
        {
            SqlDataReader ver;
            int id = 0;
            comando.Connection = conexion.AbrirConexion();
            comando.CommandText = "BuscarIdModulo";
            comando.CommandType = CommandType.StoredProcedure;
            comando.Parameters.Clear();
            comando.Parameters.AddWithValue("@CampoBusqueda", campo);
            comando.Parameters.AddWithValue("@IdModelo", IdModelo);
            ver = comando.ExecuteReader();

            while (ver.Read())
            {
                id = int.Parse(ver[0].ToString());
            }

            conexion.CerrarConexion();
            ver.Close();
            return id;
        }

        //Nos ayudara a saber si un cliente ya tiene ese modulo con una clave
        //CAMBIAR AQU Y EN BASE DE DATOS
        public bool VerificarModuloAsignado(string Serie, int IdModulo)
        {
            SqlDataReader leer;
            bool FolioRepetido = false;
            comando.Connection = conexion.AbrirConexion();
            comando.CommandText = "VerificarModuloCliente";
            comando.CommandType = CommandType.StoredProcedure;
            comando.Parameters.Clear();
            comando.Parameters.AddWithValue("@NumeroSerie", Serie);
            comando.Parameters.AddWithValue("@IdModulo", IdModulo);


            leer = comando.ExecuteReader();
            

            //Nos ayuda a comprobar si el inventario fue modificado(Dependiendo si se haya modificado algo o no)
            if (leer.Read())
            {
                leer.Close();
                return FolioRepetido = true;
            }
            else
            {
                leer.Close();
                return FolioRepetido;
            }
        }

        
    }
}
