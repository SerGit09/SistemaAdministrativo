using CobranzaSP.Lógica;
using DocumentFormat.OpenXml;
using System;
using System.Data.SqlClient;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using CobranzaSP.Modelos;
using System.Runtime.InteropServices;

namespace CobranzaSP.Formularios
{
    public partial class ReportesGarantias : Form
    {
        string TipoBusqueda = "";
        FuncionesFormularios Formulario = new FuncionesFormularios();
        AccionesLógica NuevaAccion = new AccionesLógica();
        LogicaGarantia lgGarantia = new LogicaGarantia();
        public ReportesGarantias()
        {
            InitializeComponent();
            InicioAplicacion();
        }

        #region Inicio

        public void InicioAplicacion()
        {
            //Opciones reporte
            string[] Opciones = { "", "Modelo", "Cliente", "Rango Fecha" };
            cboOpcionReporte.Items.AddRange(Opciones);
            cboOpcionReporte.SelectedIndex = 0;
        }

        #endregion

        #region Validaciones
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

        #region OpcionesReporte
        private void cboOpcionReporte_SelectedIndexChanged(object sender, EventArgs e)
        {
            btnGenerarReporte.Enabled = true;
            TipoBusqueda = cboOpcionReporte.SelectedItem.ToString();
            ReiniciarOpciones();
            switch (TipoBusqueda)
            {
                case "Modelo":
                    MostrarComboBox(cboMarca, "SeleccionarMarca", true);
                    MostrarLabelOpcion(true);
                    break;
                case "Cliente":
                    grpCliente.Visible = true;
                    radTodosLosClientes.Checked = true;
                    MostrarRadioButtonsCliente(true);
                    ; break;
            }
        }

        public void ReiniciarOpciones()
        {
            MostrarRadioButtonsCliente(false);
            MostrarModelos(false);
            radTodosLosClientes.Checked = false;
            grpCliente.Visible = false;
            cboMarca.Visible = false;
            MostrarRadioButtonsModelos(false);
            grpModelos.Visible = false;
        }

        public void MostrarLabelOpcion(bool Mostrar)
        {
            lblOpcion.Visible = Mostrar;
        }

        public void MostrarRadioButtonsCliente(bool Mostrar)
        {
            rdUnCliente.Visible = Mostrar;
            radTodosLosClientes.Visible = Mostrar;
        }

        private void MostrarComboBox(ComboBox cbo, string sp, bool Mostrar)
        {
            cbo.Visible = Mostrar;
            if (Mostrar)
            {
                Formulario.LlenarComboBox(cbo, sp, 0);
            }
        }

        public void MostrarModelos(bool Mostrar)
        {
            lblModelos.Visible = Mostrar;
            cboModelos.Visible = Mostrar;
        }

        private void radTodosLosClientes_CheckedChanged(object sender, EventArgs e)
        {
            MostrarComboBox(cboClientes, "", false);
            lblCliente.Visible = false;
        }

        private void rdUnCliente_CheckedChanged(object sender, EventArgs e)
        {
            lblCliente.Visible = true;
            MostrarComboBox(cboClientes, "SeleccionarClientesServicios", true);
        }

        private void cboOpcionElegida_SelectedIndexChanged(object sender, EventArgs e)
        {
            MostrarModelos(false);
            grpModelos.Visible = false;
            //En dado caso que se haya seleccionado algo de las marcas y mientras no estemos buscando un registro en especifico
            if (cboMarca.SelectedItem.ToString() != " ")
            {
                grpModelos.Visible = true;
                int IdMarca = NuevaAccion.BuscarId(cboMarca.SelectedItem.ToString(), "ObtenerIdMarca");
                Formulario.LlenarComboBoxModelos(cboModelos, "SeleccionarCartuchos", IdMarca);
                if(TipoBusqueda != "Cliente")
                {
                    MostrarModelos(true);
                }
            }
        }
        public void MostrarRadioButtonsModelos(bool Mostrar)
        {
            radTodosModelos.Visible =Mostrar;
            radUnModelo.Visible = Mostrar;
            radTodosModelos.Checked = Mostrar;
        }

        #endregion

        #region Botones
        private void btnGenerarReporte_Click(object sender, EventArgs e)
        {
            bool DatosEncontrados;
            //if (!ValidarReporte())
            //    return;
            try
            {
                DatosReporteGarantia NuevoReporte = new DatosReporteGarantia()
                {
                    FechaInicio = dtpFechaInicial.Value,
                    FechaFinal = dtpFechaFinal.Value,
                    TipoBusqueda = TipoBusqueda,
                    Marca = ColocarMarca(),
                    ParametroBusqueda = DeterminarParametroBusqueda(),
                    SegundoParametro = DeterminarSegundoParametro()
                };
                //MessageBox.Show(NuevoReporte.ParametroBusqueda);
                //MessageBox.Show(NuevoReporte.SegundoParametro);
                //MessageBox.Show(NuevoReporte.Marca);

                DatosEncontrados = lgGarantia.DeterminarTipoReporte(NuevoReporte);

                //if (!DatosEncontrados)
                //{
                //    MessageBox.Show("NO EXISTEN REGISTROS", "DATOS NO ENCONTRADOS", MessageBoxButtons.OK, MessageBoxIcon.Error);
                //}
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ocurrio un error: " + ex.Message);
            }
        }

        public string ColocarMarca()
        {
            string Marca = "";
            if (TipoBusqueda != "Rango Fecha" && radTodosLosClientes.Checked == false)
            {
                Marca = cboMarca.SelectedItem.ToString();
            }
            return Marca;
        }

        public string DeterminarParametroBusqueda()
        {
            string Parametro = "";
            switch (TipoBusqueda)
            {
                case "Modelo":
                    Parametro = cboModelos.SelectedItem.ToString();
                    break;
                case "Cliente":
                    Parametro = (radTodosLosClientes.Checked) ? "" : cboClientes.SelectedItem.ToString();
                    ; break;
            }
            return Parametro;
        }

        public string DeterminarSegundoParametro()
        {
            string SegundoParametro = "";
            if (TipoBusqueda == "Cliente")
            {
                if (!radTodosLosClientes.Checked)
                {
                    SegundoParametro = cboModelos.SelectedItem.ToString();
                }
            }
            return SegundoParametro;
        }

        #endregion

        #region Eventos
        #endregion

        private void chkRango_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void cboClientes_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cboClientes.SelectedItem.ToString() != " ")
            {
                MostrarComboBox(cboMarca, "SeleccionarMarca", true);
                MostrarRadioButtonsModelos(true);
                lblOpcion.Visible = true;
            }
        }

        private void radTodosModelos_CheckedChanged(object sender, EventArgs e)
        {
            if(TipoBusqueda == "Cliente")
            {
                cboModelos.Visible = false;
                lblModelos.Visible = false;
            }
        }

        private void radUnModelo_CheckedChanged(object sender, EventArgs e)
        {
            cboModelos.Visible = true;
            lblModelos.Visible = true;
        }
    }
}
