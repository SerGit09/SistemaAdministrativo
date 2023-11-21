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
    internal class CuentasPagadasLogica
    {
        private CD_Conexion cn = new CD_Conexion();
        SqlCommand comando = new SqlCommand();
        AccionesLógica NuevaAccion = new AccionesLógica();

        public string ActualizarCobranzaMesEspecifico(int Mes, int Año)
        {
            string TotalCobranza;
            SqlCommand comando = new SqlCommand();
            comando.Connection = cn.AbrirConexion();
            comando.CommandText = "MostrarCuentasMesEspecifico";
            comando.CommandType = CommandType.StoredProcedure;
            comando.Parameters.AddWithValue("@Mes", Mes);
            //Mandamos la variable y especificamos el tipo de dato sql que sera
            SqlParameter total = new SqlParameter("@totalCobrado", SqlDbType.Money); total.Direction = ParameterDirection.Output;

            comando.Parameters.AddWithValue("@Año", Año);

            comando.Parameters.Add(total);
            //comando.CommandText = "ObtenerTotales";

            comando.ExecuteNonQuery();
            //lblTotalCobrado.Text = comando.Parameters["@totalCobranza"].Value.ToString();
            double ValorRecbido = double.Parse(comando.Parameters["@totalCobrado"].Value.ToString());//LINEA QUE ESTA FALLANDO
            comando.Parameters.Clear();
            cn.CerrarConexion();
            return TotalCobranza = String.Format("{0:c}", ValorRecbido);
            //lblTotalCobrado.Text = ValorRecbido +"";
        }

        public DataTable MostrarCuentasPagadasMesEspecifico(int Mes, int Año)
        {
            DataTable tabla = new DataTable();
            SqlDataReader leer;
            comando.Connection = cn.AbrirConexion();
            comando.CommandText = "MostrarCuentasPagadasMesElegido";
            comando.CommandType = CommandType.StoredProcedure;
            comando.Parameters.AddWithValue("@Mes", Mes);
            comando.Parameters.AddWithValue("@Año", Año);
            leer = comando.ExecuteReader();
            tabla.Load(leer);
            comando.Parameters.Clear();
            cn.CerrarConexion();
            return tabla;
        }


        public void ModificarCuentaPagada(CuentaPagada CuentaPagada)
        {
            comando.Connection = cn.AbrirConexion();
            comando.CommandText = "ModificarCuentaPagada";
            comando.CommandType = CommandType.StoredProcedure;
            comando.Parameters.AddWithValue("@Id", CuentaPagada.Id);
            comando.Parameters.AddWithValue("@FechaPago", CuentaPagada.FechaPago);
            comando.ExecuteNonQuery();

            comando.Parameters.Clear();
        }
    }
}
