using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.SqlClient;

namespace CobranzaSP.Lógica
{
    internal class CD_Conexion
    {
        //private SqlConnection Conexion = new SqlConnection("Server=DESKTOP-7KHO1J8;DataBase= CobranzaSP;Integrated Security=true");
        private SqlConnection Conexion = new SqlConnection("data source=192.168.0.17,1433;initial catalog=CobranzaSP; user id=user1;password=1234;");


        //private SqlConnection Conexion = new SqlConnection("Server=DESKTOP-QFUKV11;DataBase= SpeedToner;Integrated Security=true");

        public SqlConnection AbrirConexion()
        {
            if (Conexion.State == ConnectionState.Closed)
            {
                Conexion.Open();
            }
            return Conexion;
        }

        public SqlConnection CerrarConexion()
        {
            if (Conexion.State == ConnectionState.Open)
            {
                Conexion.Close();
            }
            return Conexion;
        }
    }
}
