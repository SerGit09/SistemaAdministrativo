using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.Data;
using CobranzaSP.Lógica;
using CobranzaSP.Modelos;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.Xml.Linq;
using System.Drawing;
using System.Diagnostics;
using CobranzaSP.Formularios;
using System.IO;

namespace CobranzaSP.Lógica
{
    internal class LogicaModulosBodega
    {
        private CD_Conexion conexion = new CD_Conexion();
        SqlCommand comando = new SqlCommand();

        public string RegistrarModulo(ModuloBodega NuevoModulo, string sp)
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

            respuesta = comando.ExecuteNonQuery();
            string Mensaje = (respuesta > 0) ? "Se " + AccionRealizada + " correctamente el catalogo" : "Algo salio mal, no se " + AccionRealizada + " el registro";


            comando.Parameters.Clear();
            conexion.CerrarConexion();
            return Mensaje;
        }

        public int BuscarIdModulo(string campo,int IdModelo)
        {
            SqlDataReader ver;
            int id = 0;
            comando.Connection = conexion.AbrirConexion();
            comando.CommandText = "BuscarIdModulo";
            comando.CommandType = CommandType.StoredProcedure;
            comando.Parameters.AddWithValue("@CampoBusqueda", campo);
            comando.Parameters.AddWithValue("@IdModelo", IdModelo);
            ver = comando.ExecuteReader();
            comando.Parameters.Clear();

            while (ver.Read())
            {
                id = int.Parse(ver[0].ToString());
            }

            conexion.CerrarConexion();
            ver.Close();
            return id;
        }
    }
}
