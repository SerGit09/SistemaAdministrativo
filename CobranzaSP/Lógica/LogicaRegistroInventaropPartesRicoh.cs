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
    internal class LogicaRegistroInventaropPartesRicoh
    {
        private CD_Conexion conexion = new CD_Conexion();
        SqlCommand comando = new SqlCommand();

        public string AgregarRegistroInventario(MovimientoParteRicoh NuevoMovimiento)
        {
            SqlDataReader leer;
            int valor = 0;
            comando.Connection = conexion.AbrirConexion();

            comando.CommandText = "ModificarCantidadInventarioPartes";
            comando.CommandType = CommandType.StoredProcedure;

            comando.Parameters.AddWithValue("@IdNumeroParte", NuevoMovimiento.IdNumeroParte);
            comando.Parameters.AddWithValue("@Cantidad", NuevoMovimiento.Cantidad);
            comando.Parameters.AddWithValue("@TipoMovimiento", NuevoMovimiento.TipoMovimiento);


            valor = comando.ExecuteNonQuery();
            comando.Parameters.Clear();
            //Nos ayuda a comprobar si el inventario fue modificado(Dependiendo si se haya modificado algo o no)
            if (valor > 0)
            {
                //En dado caso de que haya modificado el inventario, se agregara el registro a la tabla de registros

                //Este es el que se utiliza en mi base de datos
                comando.CommandText = "AgregarRegistroInventarioPartesRicoh";

                //BD PRINCIPAL
                //comando.CommandText = "AgregarRegistroInventarioRespaldo";

                comando.CommandType = CommandType.StoredProcedure;

                comando.Parameters.AddWithValue("@IdNumeroParte", NuevoMovimiento.IdNumeroParte);
                comando.Parameters.AddWithValue("@IdTipoPersona", NuevoMovimiento.IdTipoPersona);
                comando.Parameters.AddWithValue("@TipoMovimiento", NuevoMovimiento.TipoMovimiento);
                comando.Parameters.AddWithValue("@Cantidad", NuevoMovimiento.Cantidad);
                comando.Parameters.AddWithValue("@Fecha", NuevoMovimiento.Fecha);

                valor = comando.ExecuteNonQuery();
                comando.Parameters.Clear();
                conexion.CerrarConexion();
                return "Se ha agregado el resgitro correctamente. Se ha actualizado el inventario";
            }
            else
            {
                conexion.CerrarConexion();
                return "No se ha agregado el registro. La cantidad excede la cantidad en existencia";
            }
            return "";
        }
    }
}
