using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.SqlClient;
using System.Net.Mail;
using System.Configuration;
using System.Net;
using System.Security.Cryptography.X509Certificates;
using System.Net.Security;
using System.Net.Mime;
using CobranzaSP.Modelos;


namespace CobranzaSP.Lógica
{
    internal class ContactoLogica
    {
        private CD_Conexion conexion = new CD_Conexion();
        SqlCommand comando = new SqlCommand();
        AccionesLógica NuevaAccion = new AccionesLógica();

        public string AgregarContacto(Contacto nuevoContacto, bool NuevoCliente)
        {
            int respuesta = 0; 
            string mensaje = "";
            comando.Connection = conexion.AbrirConexion();
            if (NuevoCliente)
            {
                comando.CommandText = "AgregarCliente";
                comando.CommandType = CommandType.StoredProcedure;
                comando.Parameters.AddWithValue("@DiasCredito", nuevoContacto.DiasCredito);
                comando.Parameters.AddWithValue("@Cliente", nuevoContacto.Nombre);

            }
            else//Si no, solamente estariamos agregando un correo
            {
                comando.CommandText = "AgregarContacto";
                comando.CommandType = CommandType.StoredProcedure;
                comando.Parameters.AddWithValue("@IdCliente", nuevoContacto.IdCliente);
            }
            
            comando.Parameters.AddWithValue("@Correo", nuevoContacto.Correo);

            respuesta = comando.ExecuteNonQuery();
            mensaje = (respuesta > 0) ? "Contacto agregado correctamente" : "Algo ha salido mal, no agrego el registro";

            comando.Parameters.Clear();
            return mensaje;
        }

        public string ModificarContacto(Contacto nuevoContacto, int Id)
        {
            int respuesta = 0;
            string mensaje = "";
            comando.Connection = conexion.AbrirConexion();
            comando.CommandText = "ModificarContacto";
            comando.CommandType = CommandType.StoredProcedure;

            comando.Parameters.AddWithValue("@Id", Id);
            comando.Parameters.AddWithValue("@IdCliente", nuevoContacto.IdCliente);
            comando.Parameters.AddWithValue("@Correo", nuevoContacto.Correo);

            respuesta = comando.ExecuteNonQuery();
            mensaje = (respuesta > 0) ? "Contacto modificado correctamente" : "Algo ha salido mal, no agrego el registro";

            comando.Parameters.Clear();
            return mensaje;
        }

        public string EliminarContacto(int Id)
        {
            int respuesta = 0;
            string mensaje = "";
            comando.Connection = conexion.AbrirConexion();
            comando.CommandText = "EliminarContacto";
            comando.CommandType = CommandType.StoredProcedure;

            comando.Parameters.AddWithValue("@Id", Id);

            respuesta = comando.ExecuteNonQuery();
            mensaje = (respuesta > 0) ? "Contacto eliminado correctamente" : "Algo ha salido mal, no se elimino el registro";

            comando.Parameters.Clear();

            conexion.CerrarConexion();
            return mensaje;
        }
    }
}
