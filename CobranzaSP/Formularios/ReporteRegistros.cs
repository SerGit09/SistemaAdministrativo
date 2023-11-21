using CobranzaSP.Lógica;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Data.SqlClient;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;

namespace CobranzaSP.Formularios
{
    public partial class ReporteRegistros : Form
    {
        public ReporteRegistros()
        {
            InitializeComponent();
            InicioAplicacion();
        }
        AccionesLógica NuevaAccion = new AccionesLógica();
        CD_Conexion cn = new CD_Conexion();
        LogicaRegistro AccionRegistro = new LogicaRegistro();
        string TipoBusqueda;

        #region Inicio
        public void InicioAplicacion()
        {
            cboMarca.DropDownStyle = ComboBoxStyle.DropDownList;
            cboBusqueda.DropDownStyle = ComboBoxStyle.DropDownList;
            cboTipoBusqueda.DropDownStyle = ComboBoxStyle.DropDownList;
            LlenarComboBox(cboBusqueda, "SeleccionarClientesServicios", 0);
            LlenarComboBox(cboClienteEspecifico, "SeleccionarClientes", 0);
            LlenarComboBoxTipoBusqueda();
        }

        public void LlenarComboBoxTipoBusqueda()
        {
            string[] Opciones = { "", "Cliente", "Modelo" };
            cboTipoBusqueda.Items.AddRange(Opciones);
        }

        public void LlenarComboBox(ComboBox cb, string sp, int indice)
        {
            SqlDataReader dr;
            cb.Items.Clear();
            if (sp == "SeleccionarModelos" || sp == "SeleccionarCartuchos")
            {
                dr = NuevaAccion.LlenarComboBoxModelos(sp, indice);
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

        private void btnCerrar_Click(object sender, EventArgs e)
        {
            this.Close();
        }


        private void btnGenerar_Click(object sender, EventArgs e)
        {
            string Parametro;
            if (cboBusqueda.SelectedIndex == -1)
                Parametro = "";
            else
                Parametro = cboBusqueda.SelectedItem.ToString();

            //if (TipoBusqueda == "Modelo")
            //    Parametro = cboMarca.SelectedItem.ToString() +" " + Parametro;

            DateTime FechaInicio = dtpFechaInicio.Value;
            DateTime FechaFinal = dtpFechaFinal.Value;

            if(TipoBusqueda == "Modelo")
            {
                string Marca = cboMarca.SelectedItem.ToString();
                string Cliente = "";
                if (chkCliente.Checked)
                {
                    Cliente = cboClienteEspecifico.SelectedItem.ToString();
                }
                //En esta parte podemos comenzar con el filtro de cartuchos
                MessageBox.Show(AccionRegistro.ReporteRegistroCartucho(FechaInicio, FechaFinal, TipoBusqueda, Parametro,Marca,Cliente), "REPORTE REGISTRO", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show(AccionRegistro.ReporteRegistrosInventario(FechaInicio, FechaFinal, TipoBusqueda, Parametro), "REPORTE REGISTRO", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            cboBusqueda.SelectedIndex = 0;

        }

        #region PanelSuperior
        [DllImport("user32.DLL", EntryPoint = "ReleaseCapture")]
        private extern static void ReleaseCapture();
        [DllImport("user32.DLL", EntryPoint = "SendMessage")]

        private extern static void SendMessage(System.IntPtr hWnd, int wMsg, int wParam, int lparam);

        private void panelSuperior_MouseDown(object sender, MouseEventArgs e)
        {
            ReleaseCapture();
            SendMessage(this.Handle, 0x112, 0xf012, 0);
        }
        #endregion

        private void cboBusqueda_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void cboTipoBusqueda_SelectedIndexChanged(object sender, EventArgs e)
        {
            btnGenerar.Enabled = true;
            TipoBusqueda = cboTipoBusqueda.SelectedItem.ToString();

            cboBusqueda.Visible = false;

            switch (cboTipoBusqueda.SelectedItem.ToString())
            {
                case "Cliente": 
                    LlenarComboBox(cboBusqueda, "SeleccionarClientes", 0); 
                    cboMarca.Visible = false;
                    HabilitarClienteEspecifico(false);
                    cboBusqueda.Visible = true;
                    break;
                case "Modelo": 
                    LlenarComboBox(cboMarca, "SeleccionarMarca", 0); 
                    cboMarca.Visible = true;
                    HabilitarClienteEspecifico(true);
                    break;
            }
        }

        public void HabilitarClienteEspecifico(bool mostrar)
        {
            chkCliente.Visible = mostrar;
        }

        private void cboMarca_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cboMarca.SelectedItem.ToString() != " ")
            {
                int IdMarca = NuevaAccion.BuscarId(cboMarca.SelectedItem.ToString(), "ObtenerIdMarca");
                LlenarComboBox(cboBusqueda, "SeleccionarCartuchos", IdMarca);
            }
        }

        private void chkCliente_CheckedChanged(object sender, EventArgs e)
        {
            cboClienteEspecifico.Visible = (chkCliente.Checked) ? true: false;
        }

        private void cboMarca_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            if (cboMarca.SelectedItem.ToString() != " ")
            {
                int IdMarca = NuevaAccion.BuscarId(cboMarca.SelectedItem.ToString(), "ObtenerIdMarca");
                LlenarComboBox(cboBusqueda, "SeleccionarCartuchos", IdMarca);
                cboBusqueda.Visible = true;
            }
        }

        private void cboBusqueda_SelectedIndexChanged_1(object sender, EventArgs e)
        {

        }
    }
}
