using CobranzaSP.Lógica;
using CobranzaSP.Modelos;
using DocumentFormat.OpenXml.Vml.Spreadsheet;
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
    public partial class ReporteServiciosRicoh : Form
    {
        public ReporteServiciosRicoh()
        {
            InitializeComponent();
            InicioAplicacion();
        }
        FuncionesFormularios Formulario = new FuncionesFormularios();
        LogicaServiciosRicoh lgServicioRicoh =  new LogicaServiciosRicoh();
        AccionesLógica NuevaAccion = new AccionesLógica();
        LogicaReportesModulo lgReporteModulo = new LogicaReportesModulo();

        //Variables para reporte
        string TipoBusqueda = "";
        bool Validado = true;

        #region Inicio

        public void InicioAplicacion()
        {
            OpcionesReporte();

            MostrarOpcionesModulo(false);
            btnGenerarReporte.Enabled = false;
        }

        public void OpcionesReporte()
        {
            string[] Opciones = { "", "HISTORIAL MODULO","MODULOS POR SERIE DE EQUIPO","UBICACION MODULO","CLIENTE", "SERIE","RANGO DE FECHA"};
            cboOpcionReporte.Items.AddRange(Opciones);
            cboOpcionReporte.SelectedIndex = 0;
        }

        #endregion

        #region Validaciones
        public bool ValidarReporte()
        {
            erReporte.Clear();
            Validado = false;
            ValidarOpcionElegida();
            return Validado;
        }

        public void ValidarOpcionElegida()
        {
            if (TipoBusqueda == "RANGO DE FECHA")
                Validado = true;

            if(TipoBusqueda == "HISTORIAL MODULO" || TipoBusqueda == "UBICACION MODULO")
            {
                ValidarClave();
            }
            else
            {
                if (string.IsNullOrWhiteSpace(cboOpcionElegida.SelectedItem.ToString()))
                {
                    switch (TipoBusqueda)
                    {
                        case "CLIENTE": erReporte.SetError(cboOpcionElegida, "Eliga un cliente"); ; break;
                        case "SERIE": erReporte.SetError(cboOpcionElegida, "Eliga una serie de equipo"); break;
                        case "MODULOS POR SERIE DE EQUIPO": erReporte.SetError(cboOpcionElegida, "Eliga una serie de equipo"); break;
                        case "RANGO DE FECHA": Validado = true; break;
                        default: erReporte.SetError(cboOpcionElegida, "Eliga un modelo"); ; break;
                    }
                    Validado = false;
                }
                else
                {
                    Validado = true;
                }
            }
            
        }

        public void ValidarClave()
        {
            if (string.IsNullOrWhiteSpace(txtClaveModulo.Text))
            {
                erReporte.SetError(txtClaveModulo, "Escriba una clave de modulo");
                Validado = false;
            }
            else
            {
                Validado = true;
            }
        }
        #endregion

        #region OpcionesReporte
        private void cboOpcionReporte_SelectedIndexChanged(object sender, EventArgs e)
        {
            TipoBusqueda = cboOpcionReporte.SelectedItem.ToString();
            ReiniciarOpciones();

            if (TipoBusqueda != "")
            {
                chkRango.Visible = true;
                btnGenerarReporte.Enabled = true;
                switch (TipoBusqueda)
                {
                    case "CLIENTE":
                        cboOpcionElegida.Visible = true;
                        lblOpcionReporte.Visible = true;
                        lblOpcionReporte.Text = "Cliente:";
                        Formulario.LlenarComboBox(cboOpcionElegida, "SeleccionarClientesServicios", 0);
                        break;
                    case "SERIE":
                        MostrarSeries(true);
                        Formulario.LlenarComboBox(cboOpcionElegida, "SeleccionarSeriesRicoh", 0);
                        break;
                    case "RANGO DE FECHA":
                        MostrarSeleccionarFechas(true);
                        chkRango.Visible = false;
                        chkRango.Checked = true;
                        break;
                    case "HISTORIAL MODULO":
                        MostrarSeleccionarFechas(true);
                        chkRango.Visible = false;
                        chkRango.Checked = true;
                        cboOpcionElegida.Visible = false;
                        MostrarTextBoxClave(true);
                        break;
                    case "MODULOS POR SERIE DE EQUIPO":
                        MostrarSeries(true);
                        chkRango.Visible = false;
                        Formulario.LlenarComboBox(cboOpcionElegida, "SeleccionarSeriesRicoh", 0);
                        cboOpcionElegida.Items.Add("TODOS");
                        break;
                    default:
                        MostrarTextBoxClave(true);
                        cboOpcionElegida.Visible = false;
                        chkRango.Visible = false;
                        MostrarOpcionesModulo(true);
                        break;
                }
            }
        }

        public void MostrarSeries(bool MostrarControl)
        {
            cboOpcionElegida.Visible = MostrarControl;
            lblOpcionReporte.Visible = MostrarControl;
            lblOpcionReporte.Text = "Serie:";
        }

        public void MostrarTextBoxClave(bool MostrarControl)
        {
            txtClaveModulo.Visible = MostrarControl;
            lblOpcionReporte.Visible = MostrarControl;
            lblOpcionReporte.Text = "Clave:";
        }

        public void MostrarOpcionesModulo(bool MostrarControl)
        {
            lblOpcionReporte.Visible = MostrarControl;
        }

        public void ReiniciarOpciones()
        {
            cboOpcionElegida.Visible = false;
            chkRango.Checked = false;
            MostrarSeleccionarFechas(false);
            MostrarClavesModulos(false);
            MostrarOpcionesModulos(false);
            btnGenerarReporte.Enabled= false;
            lblOpcionReporte.Visible = false;
            txtClaveModulo.Text = "";
            erReporte.Clear();
            MostrarTextBoxClave(false);
            MostrarSeries(false);
        }
        public void MostrarSeleccionarFechas(bool mostrar)
        {
            lblFechaInicio.Visible = mostrar;
            lblFechaFinal.Visible = mostrar;
            dtpFechaFinal.Visible = mostrar;
            dtpFechaInicial.Visible = mostrar;
        }
        #endregion

        private void btnCerrar_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnGenerarReporte_Click(object sender, EventArgs e)
        {
            bool DatosEncontrados;
            if (!ValidarReporte())
                return;
            try
            {
                Reporte NuevoReporte = new Reporte()
                {
                    FechaInicio = dtpFechaInicial.Value,
                    FechaFinal = dtpFechaFinal.Value,
                    TipoBusqueda = TipoBusqueda,
                    RangoHabilitado = chkRango.Checked,
                    ParametroBusqueda = DeterminarParametroBusqueda()
                };
                //MessageBox.Show(NuevoReporte.TipoBusqueda);
                //MessageBox.Show(NuevoReporte.ParametroBusqueda);

                DatosEncontrados = lgReporteModulo.Pdf(NuevoReporte);

                if (!DatosEncontrados)
                {
                    MessageBox.Show("NO EXISTEN REGISTROS", "DATOS NO ENCONTRADOS", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

            }
            catch(Exception ex)
            {
                MessageBox.Show("Ocurrio un error: " + ex.Message);
            }
        }

        public string DeterminarParametroBusqueda()
        {
            string Parametro = "";
            switch (TipoBusqueda)
            {
                case "HISTORIAL MODULO":
                    Parametro = txtClaveModulo.Text;
                    break;
                case "UBICACION MODULO":
                    Parametro = txtClaveModulo.Text;
                    break;
                case "RANGO DE FECHA":
                    Parametro = "";break;
                default:
                    Parametro = cboOpcionElegida.SelectedItem.ToString();
                    break;
            }
            //MessageBox.Show(Parametro);
            return Parametro;
        }

        private void chkRango_CheckedChanged(object sender, EventArgs e)
        {
            if (chkRango.Checked)
            {
                MostrarSeleccionarFechas(true);
            }
            else
            {
                MostrarSeleccionarFechas(false);
            }

        }

        

        #region ReporteModulos
        private void cboOpcionElegida_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cboOpcionElegida.SelectedItem.ToString() != " " && (TipoBusqueda == "HISTORIAL MODULO" || TipoBusqueda == "UBICACION MODULO"))
            {
                MostrarOpcionesModulos(true);
                string opcion = cboOpcionElegida.SelectedItem.ToString();
                int IdModelo = NuevaAccion.BuscarId(opcion, "ObtenerIdModeloModulo");

                Formulario.LlenarComboBox(cboModulos, "SeleccionarModulos", IdModelo);
            }
        }

        public void MostrarOpcionesModulos(bool Mostrar)
        {
            cboModulos.Visible = Mostrar;
            lblModulo.Visible = Mostrar;
        }

        private void cboModulos_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cboModulos.SelectedItem.ToString() != " ")
            {
                MostrarClavesModulos(true);
                int IdModelo = NuevaAccion.BuscarId(cboOpcionElegida.SelectedItem.ToString(), "ObtenerIdModeloModulo");

                int IdModulo = lgReporteModulo.BuscarIdModulo(cboModulos.SelectedItem.ToString(), "ObtenerIdModulo", IdModelo);

                Formulario.LlenarComboBox(cboClaves, "ObtenerClaves", IdModulo);
            }
        }

        public void MostrarClavesModulos(bool Mostrar)
        {
            lblClaves.Visible = Mostrar;
            cboClaves.Visible = Mostrar;
        }
        #endregion
    }
}
