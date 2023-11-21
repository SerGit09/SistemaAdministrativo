using CobranzaSP.Lógica;
using CobranzaSP.Modelos;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CobranzaSP.Formularios
{
    public partial class ReporteInventario : Form
    {
        public ReporteInventario()
        {
            InitializeComponent();
            InicioAplicacion();
        }
        AccionesLógica NuevaAccion = new AccionesLógica();
        CD_Conexion cn = new CD_Conexion();
        #region Inicio

        public void InicioAplicacion()
        {
            string[] Opciones = { "Marca"};
            cboOpcionReporte.Items.AddRange(Opciones);
            cboOpcionReporte.SelectedIndex = 0;

            LlenarComboBox(cboMarcas, "SeleccionarMarca", 0);

            //Radio buttons seleccionado
            rdTodasLasMarcas.Checked = true;
            rdTodosLosModelos.Checked = true;
        }

        public void LlenarComboBox(ComboBox cb, string sp, int Marca)
        {
            SqlDataReader dr;
            cb.Items.Clear();
            if (sp == "SeleccionarModelos")
            {
                dr = NuevaAccion.LlenarComboBoxEspecifico(sp, Marca);
            }
            else
            {
                dr = NuevaAccion.LlenarComboBox(sp);
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

        #endregion

        #region Validacion
        #endregion

        #region PanelSuperior
        [DllImport("user32.DLL", EntryPoint = "ReleaseCapture")]
        private extern static void ReleaseCapture();
        [DllImport("user32.DLL", EntryPoint = "SendMessage")]

        private extern static void SendMessage(System.IntPtr hWnd, int wMsg, int wParam, int lparam);
        private void pSuperior_MouseDown(object sender, MouseEventArgs e)
        {
            ReleaseCapture();
            SendMessage(this.Handle, 0x112, 0xf012, 0);
        }

        private void btnCerrar_Click(object sender, EventArgs e)
        {
            this.Close();
        }
        #endregion

        #region Botones
        private void btnGenerarReporte_Click(object sender, EventArgs e)
        {

        }
        #endregion

        #region Eventos
        #endregion

        #region MetodosLocales
        #endregion

        #region OpcionesReporte
        private void rdTodasLasMarcas_CheckedChanged(object sender, EventArgs e)
        {
            grpModelo.Visible = false;
            OcultarMostrarElementos(true);
        }

        private void OcultarMostrarElementos(bool ocultar)
        {
            cboMarcas.Visible = !ocultar;
            rdTodosLosModelos.Visible = true;
            radUnModelo.Visible = true;
        }

        private void rdUnaMarca_CheckedChanged(object sender, EventArgs e)
        {
            OcultarMostrarElementos(false);
        }

        private void cboMarcas_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cboMarcas.SelectedItem.ToString() != " ")
            {
                grpModelo.Visible = true;
                int IdMarca = NuevaAccion.BuscarId(cboMarcas.SelectedItem.ToString(), "ObtenerIdMarca");
                LlenarComboBox(cboModelos, "SeleccionarModelos", IdMarca);
            }
        }

        private void rdTodosLosModelos_CheckedChanged(object sender, EventArgs e)
        {
            cboModelos.Visible = false;
        }

        private void radUnModelo_CheckedChanged(object sender, EventArgs e)
        {
            cboModelos.Visible = true;
        }
        #endregion
    }
}
