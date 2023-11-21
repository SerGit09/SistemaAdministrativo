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
    public partial class ReporteModulos : Form
    {
        public ReporteModulos()
        {
            InitializeComponent();
            InicioAplicacion();
        }
        string Parametro = " ";
        string TipoBusqueda = " ";
        AccionesLógica NuevaAccion = new AccionesLógica();
        LogicaReportesModulo lgReporteModulo = new LogicaReportesModulo();
        CD_Conexion cn = new CD_Conexion();
        [DllImport("user32.DLL", EntryPoint = "ReleaseCapture")]
        private extern static void ReleaseCapture();
        [DllImport("user32.DLL", EntryPoint = "SendMessage")]
        private extern static void SendMessage(System.IntPtr hWnd, int wMsg, int wParam, int lparam);

        #region Inicio

        public void InicioAplicacion()
        {
            string[] Opciones = { "", "TECNICO","RANGO DE FECHA"};
            cboOpcionReporte.Items.AddRange(Opciones);
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
        public bool ValidarCamposReporte()
        {
            bool Validado = true;
            erReporte.Clear();
            if (cboOpcionReporte.SelectedItem == null)
            {
                erReporte.SetError(cboOpcionReporte, "Escoja una opcion");
                Validado = false;
            }
            return Validado;
        }
        #endregion

        #region Botones
        private void btnCerrar_Click(object sender, EventArgs e)
        {
            this.Close();
        }
        #endregion

        #region Eventos
        private void panelSuperior_MouseDown(object sender, MouseEventArgs e)
        {
            ReleaseCapture();
            SendMessage(this.Handle, 0x112, 0xf012, 0);
        }

        private void cboOpcionReporte_SelectedIndexChanged(object sender, EventArgs e)
        {
            btnGenerarReporte.Enabled = true;
            TipoBusqueda = cboOpcionReporte.SelectedItem.ToString();
            ReiniciarControles();

            switch (cboOpcionReporte.SelectedItem.ToString())
            {
                case "TECNICO": MostrarComboBox("SeleccionarSeriesEquipos", false); chkRango.Visible = true; break;
                case "RANGO DE FECHA": Parametro = " "; MostrarRangosFechas(true); break;
            }
        }
        private void chkRango_CheckedChanged(object sender, EventArgs e)
        {
            bool mostrar = (chkRango.Checked) ? true : false;
            MostrarRangosFechas(mostrar);
        }
        #endregion

        #region Metodos Locales

        public void MostrarRangosFechas(bool mostrar)
        {
            lblFechaInicio.Visible = mostrar;
            lblFechaFinal.Visible = mostrar;
            dtpFechaFinal.Visible = mostrar;
            dtpFechaInicial.Visible = mostrar;
        }

        private void ReiniciarControles()
        {
            MostrarRangosFechas(false);
            txtDato.Visible = false;
            cboClientes.Visible = false;
        }

        private void MostrarComboBox(string sp, bool MostrarRango)
        {
            LlenarComboBoxTecnicos();
            txtDato.Enabled = false;
            txtDato.Visible = false;
            cboClientes.Visible = true;
            chkRango.Visible = MostrarRango;
            //erReporte.Clear();
        }

        public void LlenarComboBoxTecnicos()
        {
            cboClientes.Items.Clear();
            string[] Tecnicos = { "", "ALVARO", "MIGUEL" };
            cboClientes.Items.AddRange(Tecnicos);
        }

        #endregion

        private void btnGenerarReporte_Click(object sender, EventArgs e)
        {
            //Nos ayudara a saber si se encontro algun registro de fusor o serie
            bool ReporteGenerado = true;
            
            DateTime FechaInicial = dtpFechaInicial.Value;
            DateTime FechaFinal = dtpFechaFinal.Value;
            bool RangoFecha = chkRango.Checked;

            try
            {
                if (!ValidarCamposReporte())
                    return;

                string mensajeError = string.Empty;
                //switch (TipoBusqueda)
                //{
                //    case "TECNICO":
                //        if (ValidarParametro())
                //            //ReporteGenerado = lgReporteModulo.GenerarPdf(TipoBusqueda,Parametro, FechaInicial,FechaFinal,RangoFecha);
                //        break;
                //    case "RANGO DE FECHA": ReporteGenerado = lgReporteModulo.GenerarPdf(TipoBusqueda, Parametro, FechaInicial, FechaFinal, true); ; break;
                //}

                if (!string.IsNullOrEmpty(mensajeError) || !ReporteGenerado)
                    MessageBox.Show("¡NO SE ENCONTRARON REGISTROS!", "DATOS NO ENCONTRADOS", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private bool ValidarParametro()
        {
            Parametro = cboClientes.SelectedItem.ToString();
            return NuevaAccion.VerificarExistenciaRegistro(Parametro, "BuscarExistenciaHojaEntregaTecnico");
        }
    }
}
