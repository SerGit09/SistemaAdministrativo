using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Data;

namespace CobranzaSP.Lógica
{
    internal class FuncionesFormularios
    {
        AccionesLógica nuevaAccion = new AccionesLógica();
        CD_Conexion conexion = new CD_Conexion();

        CD_Conexion cn = new CD_Conexion();
        SqlCommand comando = new SqlCommand();



        public void PropiedadesDtg(DataGridView datos)
        {
            //Solo lectura
            datos.ReadOnly = true;
            //No agregar renglones
            datos.AllowUserToAddRows = false;
            //No borrar renglones
            datos.AllowUserToDeleteRows = false;
            //Ajustar automaticamente el ancho de las columnas
            datos.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            //dtgServicios.AutoResizeColumns(DataGridViewAutoSizeColumnsMo‌​de.Fill);
            datos.AutoResizeColumns();
        }

        public void AbrirForm(object formNuevo)
        {
            //Declaramos la forma
            Form fh = formNuevo as Form;

            //Mostramos la forma 
            fh.Show();
        }


        public void LlenarComboBox(ComboBox cb, string sp, int Marca)
        {
            SqlDataReader dr;
            cb.Items.Clear();
            if (Marca != 0 || sp == "SeleccionarModelos")
            {
                dr = nuevaAccion.LlenarComboBoxEspecifico(sp, Marca);
            }
            else
            {
                dr = nuevaAccion.LlenarComboBox(sp);
            }

            while (dr.Read())
            {
                //Agregamos las opciones dependiendo los registros que nos devolvieron
                cb.Items.Add(dr[0].ToString());
            }
            //Agregamos un espacio en blanco y lo asignamos como opcion por defecto
            cb.Items.Insert(0, " ");
            cb.SelectedIndex = 0;
            dr.Close();
            cn.CerrarConexion();
        }

        public void LlenarComboBoxModelos(ComboBox cb, string sp, int indice)
        {
            SqlDataReader dr;
            cb.Items.Clear(); 
            dr = nuevaAccion.LlenarComboBoxModelos(sp, indice);

            while (dr.Read())
            {
                //Agregamos las opciones dependiendo los registros que nos devolvieron
                cb.Items.Add(dr[0].ToString());
            }

            //Agregamos un espacio en blanco y lo asignamos como opcion por defecto
            cb.Items.Insert(0, " ");
            cb.SelectedIndex = 0;
            dr.Close();
            cn.CerrarConexion();
        }
    }
}
