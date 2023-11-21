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
using System.Windows.Forms;
using System.Globalization;

namespace CobranzaSP.Lógica
{
    internal class GraficaLogica
    {
        private CD_Conexion conexion = new CD_Conexion();
        SqlCommand comando = new SqlCommand();
        public double ObtenerTotalMes(int Año,int Mes, string sp)
        {
            double Total = 0;
            SqlDataReader dr;

            comando.Connection = conexion.AbrirConexion();
            comando.CommandText = sp;
            comando.CommandType = CommandType.StoredProcedure;

            comando.Parameters.AddWithValue("@Mes", Mes);
            comando.Parameters.AddWithValue("Año", Año);
            dr = comando.ExecuteReader();
            comando.Parameters.Clear();

            if (dr.Read())
            {
                Total = double.Parse(dr[0].ToString());
            }

            comando.Parameters.Clear();
            conexion.CerrarConexion();
            return Total;
        }


        public DataTable ObtenerTablaResultados(string sp)
        {
            DataTable tabla = new DataTable();
            SqlDataReader leer;
            comando.Connection = conexion.AbrirConexion();
            comando.CommandText = sp;
            comando.CommandType = CommandType.StoredProcedure;
            leer = comando.ExecuteReader();

            tabla.Columns.Add("Mes", typeof(string));
            tabla.Columns.Add("Remisiones", typeof(string));
            tabla.Columns.Add("Mostrador", typeof(string));
            tabla.Columns.Add("Facturas", typeof(string));
            tabla.Columns.Add("Total", typeof(string));


            while (leer.Read())
            {
                DataRow row =tabla.NewRow();
                double TotalRemisiones = double.Parse(leer[1].ToString());
                double TotalMostrador = double.Parse(leer[2].ToString());
                double TotalFacturas = double.Parse(leer[3].ToString());
                double Total = TotalMostrador + TotalRemisiones + TotalFacturas;

                row["Mes"] = leer[0];
                row["Remisiones"] = TotalRemisiones.ToString("C", CultureInfo.CurrentCulture);
                row["Mostrador"] = TotalMostrador.ToString("C", CultureInfo.CurrentCulture); 
                row["Facturas"] = TotalFacturas.ToString("C", CultureInfo.CurrentCulture); 
                row["Total"] = Total.ToString("C", CultureInfo.CurrentCulture);

                tabla.Rows.Add(row);
            }
            tabla.Load(leer);
            conexion.CerrarConexion();
            leer.Close();
            return tabla;
        }

        public DataTable ObtenerTablaCobradoMeses(string sp)
        {
            DataTable tabla = new DataTable();
            SqlDataReader leer;
            comando.Connection = conexion.AbrirConexion();
            comando.CommandText = sp;
            comando.CommandType = CommandType.StoredProcedure;
            leer = comando.ExecuteReader();

            tabla.Columns.Add("Mes", typeof(string));
            tabla.Columns.Add("Total", typeof(string));
            tabla.Columns.Add("Compras", typeof(string));
            tabla.Columns.Add("Ganancia", typeof(string));


            while (leer.Read())
            {
                DataRow row = tabla.NewRow();

                double Total = double.Parse(leer[1].ToString());
                double Compras = double.Parse(leer[2].ToString());
                double Ganancia = double.Parse(leer[3].ToString());
                row["Mes"] = leer[0];
                row["Total"] = Total.ToString("C", CultureInfo.CurrentCulture);
                row["Compras"] = Compras.ToString("C", CultureInfo.CurrentCulture);
                row["Ganancia"] = Ganancia.ToString("C", CultureInfo.CurrentCulture);

                tabla.Rows.Add(row);
            }
            tabla.Load(leer);
            conexion.CerrarConexion();
            leer.Close();
            return tabla;
        }
    }
}
