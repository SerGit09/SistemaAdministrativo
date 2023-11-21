using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Data.SqlClient;
using CobranzaSP.Lógica;
using CobranzaSP.Modelos;

namespace CobranzaSP.Formularios
{
    public partial class ReporteServicio : Form
    {
        public ReporteServicio()
        {
            InitializeComponent();
            InicioAplicacion();
        }

        string Parametro = " ";
        string TipoBusqueda = " ";

        AccionesLógica NuevaAccion = new AccionesLógica();
        CD_Conexion cn = new CD_Conexion();
        LogicaReporteServicio LgReporteServicio = new LogicaReporteServicio();

        [DllImport("user32.DLL", EntryPoint = "ReleaseCapture")]
        private extern static void ReleaseCapture();
        [DllImport("user32.DLL", EntryPoint = "SendMessage")]
        private extern static void SendMessage(System.IntPtr hWnd, int wMsg, int wParam, int lparam);

        #region Inicio
        public void InicioAplicacion()
        {
            //Controles desactivados inicialmente
            btnGenerarReporte.Enabled = true;
            txtDato.Enabled = false;

            //Agregamos opciones al combo box
            AgregarOpcionesBusqueda();

            //Denegamos escritura en el combo box
            cboOpcionReporte.DropDownStyle = ComboBoxStyle.DropDownList;

            //Llenamos nuestro combobox de Clientes
            //LlenarComboBox(cboClientes, "SeleccionarClientes", 0);

            dtpFechaInicial.MaxDate = DateTime.Today;
            dtpFechaFinal.MaxDate = DateTime.Today;
        }

        public void AgregarOpcionesBusqueda()
        {
            cboOpcionReporte.Items.Add("CLIENTE");
            cboOpcionReporte.Items.Add("SERIE");
            cboOpcionReporte.Items.Add("FECHA");
            cboOpcionReporte.Items.Add("FUSOR");
            cboOpcionReporte.Items.Add("MARCA");
            cboOpcionReporte.Items.Add("TODOS LOS FUSORES");
        }

        public void LlenarComboBox(ComboBox cb, string sp, int indice)
        {
            cb.Items.Clear();

            SqlDataReader dr = NuevaAccion.LlenarComboBox(sp);

            while (dr.Read())
            {
                cb.Items.Add(dr[indice].ToString());
            }

            cb.Items.Insert(0, " ");
            cb.SelectedIndex = 0;
            dr.Close();
            cn.CerrarConexion();
        }
        #endregion

        #region Validaciones
        private bool ValidarCampos()
        {
            bool Validado = true;
            erReporte.Clear();
            if (cboOpcionReporte.SelectedItem == null)
            {
                erReporte.SetError(cboOpcionReporte, "Escoja una opcion");
                Validado = false;
            }
            else
            {
                switch (cboOpcionReporte.SelectedItem.ToString())
                {
                    case "SERIE": Validado = ValidarCb(); break;
                    case "TECNICO": Validado = ValidarTxtDato(); break;
                    case "FUSOR": Validado = ValidarCb(); break;
                    case "CLIENTE": Validado = ValidarCb(); break;
                    default:
                        break;
                }
            }
            return Validado;
        }

        private bool ValidarTxtDato()
        {
            bool Validado = true;
            if (txtDato.Text == "")
            {
                erReporte.SetError(txtDato, "Campo Obligatorio");
                Validado = false;
            }
            return Validado;
        }

        private bool ValidarCb()
        {
            bool Validado = true;
            if (cboClientes.SelectedItem == " ")
            {
                erReporte.SetError(cboClientes, "Escoga un cliente");
                Validado = false;
            }
            return Validado;
        }

        private void txtDato_KeyPress(object sender, KeyPressEventArgs e)
        {
            Validacion.SoloLetrasYNumeros(e);
        }
        #endregion

        #region Botones
        private void btnCerrar_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnGenerarReporte_Click(object sender, EventArgs e)
        {
            //Nos ayudara a saber si se encontro algun registro de fusor o serie
            bool Encontrado = true;
            bool Fusor = false;
            DateTime FechaInicial = dtpFechaInicial.Value;
            DateTime FechaFinal = dtpFechaFinal.Value;
            bool RangoFecha = chkRango.Checked;

            Reporte nuevoReporte = new Reporte()
            {
                FechaInicio = dtpFechaInicial.Value,
                FechaFinal = dtpFechaFinal.Value,
                RangoHabilitado = chkRango.Checked,
                TipoBusqueda = TipoBusqueda
            };

            nuevoReporte.ParametroBusqueda = (cboClientes.SelectedIndex > 0) ?cboClientes.SelectedItem.ToString(): "";

            try
            {
                if (!ValidarCampos())
                    return;
                bool DatosEncontrados = LgReporteServicio.DeterminarTipoReporte(nuevoReporte,"");
                if (!DatosEncontrados)
                    MessageBox.Show("¡NO SE ENCONTRARON REGISTROS!", "DATOS NO ENCONTRADOS", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        #endregion


        #region Eventos
        private void pSuperior_Paint(object sender, PaintEventArgs e)
        {

        }
        private void cboOpcionReporte_SelectedIndexChanged(object sender, EventArgs e)
        {
            btnGenerarReporte.Enabled = true;
            TipoBusqueda = cboOpcionReporte.SelectedItem.ToString();
            ReiniciarControles();
            //Parametro = cboClientes.SelectedItem.ToString();
            switch (cboOpcionReporte.SelectedItem.ToString())
            {
                case "SERIE": MostrarComboBox("SeleccionarSeriesEquipos", false); chkRango.Visible = true; break;
                case "FECHA": Parametro = " "; chkRango.Visible = false; MostrarRangosFechas(true); break;
                case "FUSOR": MostrarComboBox("SeleccionarFusores", false);  break;
                case "MARCA": MostrarComboBox("SeleccionarMarca", false);chkRango.Visible = true;  break;
                case "TODOS LOS FUSORES":TipoBusqueda = "Fusor"; Parametro = ""; chkRango.Visible = true; break;
                case "CLIENTE": MostrarComboBox("SeleccionarClientesServicios", true); break;
            }
        }

        private void pSuperior_MouseDown(object sender, MouseEventArgs e)
        {
            ReleaseCapture();
            SendMessage(this.Handle, 0x112, 0xf012, 0);
        }

        private void chkRango_CheckedChanged(object sender, EventArgs e)
        {
            if (chkRango.Checked)
                MostrarRangosFechas(true);
            else
                MostrarRangosFechas(false);
        }
        #endregion

        #region MetodosLocales

        private void MostrarComboBox(string sp, bool MostrarRango)
        {
            LlenarComboBox(cboClientes, sp, 0);
            txtDato.Enabled = false;
            txtDato.Visible = false;
            cboClientes.Visible = true;
            chkRango.Visible = MostrarRango;
            erReporte.Clear();
        }

        private void MostrarRangosFechas(bool mostrar)
        {
            lblFechaFinal.Visible = mostrar;
            lblFechaInicio.Visible = mostrar;
            dtpFechaFinal.Visible = mostrar;
            dtpFechaInicial.Visible = mostrar;
        }

        private void ReiniciarControles()
        {
            MostrarRangosFechas(false);
            txtDato.Visible = false;
            cboClientes.Visible = false;
            chkRango.Visible = true;
        }
        #endregion
    }
}
