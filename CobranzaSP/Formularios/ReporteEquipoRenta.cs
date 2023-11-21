using CobranzaSP.Lógica;
using CobranzaSP.Modelos;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CobranzaSP.Formularios
{
    public partial class ReporteEquipoRenta : Form
    {
        public ReporteEquipoRenta()
        {
            InitializeComponent();
            InicioAplicacion();
        }
        FuncionesFormularios AccionFormulario = new FuncionesFormularios();
        AccionesLógica NuevaAccion = new AccionesLógica();
        LogicaEquipos lgEquipo = new LogicaEquipos();

        string Parametro = "";
        string TipoBusqueda = "";

        //Parametros para los reportes
        string Marca = "";
        string Modelo = "";


        #region Inicio

        public void InicioAplicacion()
        {
            LlenarCboOpcionesMostrar();
            btnMostrar.Enabled = false;
        }

        public void LlenarCboOpcionesMostrar()
        {
            string[] Opciones = { "", "Precios de Equipos", "Cliente" };
            cboOpcionMostrar.Items.AddRange(Opciones);
        }
        #endregion

        #region Validaciones
        //public bool ValidarCamposReporte()
        //{
        //    bool Validado = true;
        //    erEquipos.Clear();

        //    if (cboOpcionMostrar.SelectedItem == null || string.IsNullOrWhiteSpace(cboOpcionMostrar.SelectedItem.ToString()))
        //    {
        //        erEquipos.SetError(cboOpcionMostrar, "Campo obligatorio");
        //        Validado = false;
        //    }

        //    if (TipoBusqueda != "Todos")
        //    {
        //        if (cboBusqueda.SelectedItem == null || string.IsNullOrWhiteSpace(cboBusqueda.SelectedItem.ToString()))
        //        {
        //            erEquipos.SetError(cboBusqueda, "Campo obligatorio");
        //            Validado = false;
        //        }
        //    }


        //    return Validado;
        //}

        //public bool ValidarMarcaSeleccionada()
        //{
        //    bool Validado = true;
        //    erEquipos.Clear();

        //    if (string.IsNullOrWhiteSpace(cboMarca.SelectedItem.ToString()))
        //    {
        //        erEquipos.SetError(cboMarca, "Eliga una marca");
        //        Validado = false;
        //    }
        //    return Validado;
        //}

        //public bool ValidarModeloSeleccionado()
        //{
        //    bool Validado = true;
        //    erEquipos.Clear();

        //    if (string.IsNullOrWhiteSpace(cboModelo.SelectedItem.ToString()))
        //    {
        //        erEquipos.SetError(cboModelo, "Eliga un modelo");
        //        Validado = false;
        //    }
        //    return Validado;
        //}
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

        private void ReporteEquipoRenta_MouseDown(object sender, MouseEventArgs e)
        {

        }

        private void btnCerrar_Click(object sender, EventArgs e)
        {
            this.Close();
        }
        #endregion

        #region Opciones Reporte
        private void cboOpcionMostrar_SelectedIndexChanged(object sender, EventArgs e)
        {
            TipoBusqueda = cboOpcionMostrar.SelectedItem.ToString();
            ReiniciarOpciones();
            if (TipoBusqueda != "")
            {
                btnMostrar.Enabled = true;
                switch (TipoBusqueda)
                {
                    case "Cliente":
                        OpcionesClientes(true);
                        break;
                    case "Precios de Equipos":
                        MostrarOpcionesMarca(true);
                        ;
                        break;
                }
            }
        }

        public void ReiniciarOpciones()
        {
            OpcionesClientes(false);
            rdTodosLosModelos.Checked = false;
            btnMostrar.Enabled = false;
        }

        public void OpcionesClientes(bool Habilitar)
        {
            grpClientes.Visible = Habilitar;
            radTodosLosClientes.Checked = Habilitar;
            MostrarOpcionesMarca(Habilitar);
        }

        public void MostrarOpcionesMarca(bool Habilitar)
        {
            grpMarcas.Visible = Habilitar;
            rdTodasLasMarcas.Checked = Habilitar;
        }

        //RADIO BUTTONS CLIENTES
        private void radTodosLosClientes_CheckedChanged(object sender, EventArgs e)
        {
            cboClienteElegido.Visible = false;
        }

        private void radUnCliente_CheckedChanged(object sender, EventArgs e)
        {
            cboClienteElegido.Visible = true;
            AccionFormulario.LlenarComboBox(cboClienteElegido, "SeleccionarClientes", 0);
        }

        //RADIO BUTTONS MARCAS
        private void rdTodasLasMarcas_CheckedChanged(object sender, EventArgs e)
        {
            cboMarca.Visible = false;
            AccionFormulario.LlenarComboBox(cboMarca, "SeleccionarMarca", 0);
            grpModelo.Visible = false;
            rdTodosLosModelos.Checked = true;
        }

        private void rdUnaMarca_CheckedChanged(object sender, EventArgs e)
        {
            cboMarca.Visible = true;
            grpModelo.Visible = true;
        }

        //RADIO BUTTONS MODELOS
        private void rdTodosLosModelos_CheckedChanged(object sender, EventArgs e)
        {
            cboModelo.Visible = false;
        }

        private void radUnModelo_CheckedChanged(object sender, EventArgs e)
        {
            cboModelo.Visible = true;
        }

        private void cboMarca_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cboMarca.SelectedIndex != 0)
            {
                int IdMarca = NuevaAccion.BuscarId(cboMarca.SelectedItem.ToString(), "ObtenerIdMarca");
                AccionFormulario.LlenarComboBox(cboModelo, "SeleccionarModelos", IdMarca);
            }
        }

        #endregion

        #region Botones
        #endregion

        #region Eventos
        #endregion

        #region Metodos Locales
        #endregion

        private void btnMostrar_Click(object sender, EventArgs e)
        {
            bool DatosEncontrados = true;
            Parametro = "";
            string StoredProcedure = "";


            //if (!ValidarCamposReporte())
            //    return;

            switch (TipoBusqueda)
            {
                case "Cliente":
                    VerificarSeleccionDeCliente();
                    VerificarSeleccionDeMarca();
                    StoredProcedure = "ReporteEquiposPrueba";
                    break;
                case "Precios de Equipos":
                    Parametro = (rdTodasLasMarcas.Checked) ? "" : cboMarca.SelectedItem.ToString();
                    VerificarSeleccionDeMarca();
                    StoredProcedure = "ReporteEquiposPrecios";
                    break;
            }

            ReporteEquipoParametro NuevoReporte = new ReporteEquipoParametro()
            {
                Parametro = Parametro,
                TipoBusqueda = TipoBusqueda,
                Marca = Marca,
                Modelo = Modelo
            };
            DatosEncontrados = lgEquipo.OrdenarEquiposPrueba(NuevoReporte, StoredProcedure);

            if (!DatosEncontrados)
            {
                MessageBox.Show("!!DATOS NO ENCONTRADOS!!", "NO SE PUDO GENERAR EL REPORTE", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            //cboOpcionMostrar.ResetText();
        }

        public void VerificarSeleccionDeCliente()
        {
            if (radTodosLosClientes.Checked)
            {
                Parametro = "";
            }
            else
            {
                Parametro = cboClienteElegido.SelectedItem.ToString();
            }
        }

        public void VerificarSeleccionDeMarca()
        {
            if (rdTodasLasMarcas.Checked)
            {
                Marca = "";
                Modelo = "";
            }
            else
            {
                Marca = cboMarca.SelectedItem.ToString();
                VerificarSeleccionDeModelo();
            }
        }

        public void VerificarSeleccionDeModelo()
        {
            if (rdTodosLosModelos.Checked)
            {
                Modelo = "";
            }
            else
            {
                Modelo = cboModelo.SelectedItem.ToString();
            }
        }
    }
}
