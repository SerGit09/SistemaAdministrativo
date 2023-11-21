using CobranzaSP.Modelos;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.Data;
using CobranzaSP.Lógica;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.Xml.Linq;
using System.Drawing;
using System.Diagnostics;
using CobranzaSP.Formularios;
using System.IO;
using DocumentFormat.OpenXml.Bibliography;

namespace CobranzaSP.Lógica
{
    internal class LogicaCatalogoModulos
    {
        private CD_Conexion conexion = new CD_Conexion();
        SqlCommand comando = new SqlCommand();
        public string RegistrarModulo(CatalogoModulo NuevoModulo, string sp)
        {
            string AccionRealizada = "agrego";
            int respuesta;
            comando.Connection = conexion.AbrirConexion();
            comando.CommandText = sp;
            comando.CommandType = CommandType.StoredProcedure;
            comando.Parameters.Clear();
            if (NuevoModulo.IdModulo > 0)
            {
                AccionRealizada = "modifico";
                comando.Parameters.AddWithValue("@IdModulo", NuevoModulo.IdModulo);
            }
            comando.Parameters.AddWithValue("@IdModelo", NuevoModulo.IdModelo);
            comando.Parameters.AddWithValue("@Modulo", NuevoModulo.Modulo);

            respuesta = comando.ExecuteNonQuery();
            string Mensaje = (respuesta > 0) ? "Se " + AccionRealizada + " correctamente el catalogo" : "Algo salio mal, no se " + AccionRealizada + " el registro";


            comando.Parameters.Clear();
            conexion.CerrarConexion();
            return Mensaje;
        }

        public bool VerificarModuloDuplicado(CatalogoModulo NuevoModulo)
        {
            SqlDataReader leer;
            bool FolioRepetido = false;
            comando.Connection = conexion.AbrirConexion();
            comando.CommandText = "VerificarModuloDuplicado";
            comando.CommandType = CommandType.StoredProcedure;

            comando.Parameters.Clear();
            comando.Parameters.AddWithValue("@ParametroBusqueda", NuevoModulo.Modulo);
            comando.Parameters.AddWithValue("@IdModelo", NuevoModulo.IdModelo);
            leer = comando.ExecuteReader();

            if (!leer.HasRows)
            {
                leer.Close();
                return FolioRepetido;
            }

            leer.Close();
            FolioRepetido = true;
            return FolioRepetido;
        }
    }
}
