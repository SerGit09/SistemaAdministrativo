using CobranzaSP.Lógica;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CobranzaSP.Formularios
{
    public partial class EquiposVendidos : Form
    {
        public EquiposVendidos()
        {
            InitializeComponent();
            InicioAplicacion();
        }
        AccionesLógica NuevaAccion = new AccionesLógica();

        #region Inicio
        public void InicioAplicacion()
        {
            MostrarDatosEquipos();
            PropiedadesDtgEquipos();
        }

        public void MostrarDatosEquipos()
        {
            //Limpiamos los datos del datagridview
            dtgEquiposVendidos.DataSource = null;
            dtgEquiposVendidos.Refresh();
            DataTable tabla = new DataTable();
            //Guardamos los registros dependiendo la consulta
            tabla = NuevaAccion.Mostrar("MostrarEquiposVendidos");
            //Asignamos los registros que optuvimos al datagridview
            dtgEquiposVendidos.DataSource = tabla;
            dtgEquiposVendidos.Columns["IdEquipo"].Visible = false;
        }
        public void PropiedadesDtgEquipos()
        {
            //Solo lectura
            dtgEquiposVendidos.ReadOnly = true;
            //No agregar renglones
            dtgEquiposVendidos.AllowUserToAddRows = false;
            //No borrar renglones
            dtgEquiposVendidos.AllowUserToDeleteRows = false;
            //Ajustar automaticamente el ancho de las columnas
            dtgEquiposVendidos.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            //dtgServicios.AutoResizeColumns(DataGridViewAutoSizeColumnsMo‌​de.Fill);
            dtgEquiposVendidos.AutoResizeColumns();
        }
        #endregion
    }
}
