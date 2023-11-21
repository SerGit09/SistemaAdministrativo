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

namespace CobranzaSP.Formularios
{
    public partial class ServicioReportes : Form
    {
        public ServicioReportes()
        {
            InitializeComponent();
            InicioAplicacion();
        }
        //Variable para validaciones
        bool Validado = true;
        
        //Variables para generar el reporte
        string Parametro = " ";
        string TipoBusqueda = " ";
        string TipoBusquedaAdicional = "";
        string Cliente = "";

        AccionesLógica NuevaAccion = new AccionesLógica();
        CD_Conexion cn = new CD_Conexion();
        LogicaReporteServicio LgReporteServicio = new LogicaReporteServicio();

        #region Inicio
        public void InicioAplicacion()
        {
            string[] OpcionesMostrar = { "", "Cliente", "Serie", "Fusor", "Fecha" };
            cboOpcionReporte.Items.AddRange(OpcionesMostrar);
        }

        public void LlenarComboBox(ComboBox cb, string sp, int Marca)
        {
            SqlDataReader dr;
            cb.Items.Clear();
            if (Marca != 0 || sp == "SeleccionarModelos")
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
            dr.Close();

            cb.Items.Insert(0, " ");
            cb.SelectedIndex = 0;
            cn.CerrarConexion();
        }
        #endregion

        #region Validaciones
        public bool DeterminarTipoValidacionDatos()
        {
            switch (TipoBusqueda)
            {
                case "Serie":
                    ValidarComboBoxOpcionElegida();
                    break;
                case "Fecha":
                    ValidarSeleccionDeMarca();
                    break;
                case "Fusor":
                    ValidarSeleccionFusor();
                    break;
                case "Cliente":
                    ValidarComboBoxOpcionElegida();
                    break;
            }
            return Validado;
        }

        public void ValidarComboBoxOpcionElegida()
        {
            erReporte.Clear();
            if (cboOpcionElegida.SelectedIndex == 0)
            {
                erReporte.SetError(cboOpcionElegida, "Campo obligatorio");
                Validado = false;
                return;
            }
            ValidarSeleccionDeMarca();
        }

        public void ValidarSeleccionDeMarca()
        {
            erReporte.Clear();
            if (rdUnaMarca.Checked)
            {
                if (cboMarca.SelectedIndex == 0)
                {
                    erReporte.SetError(cboMarca, "Seleccione una marca");
                    Validado = false;
                    return;
                }
            }
            ValidarSeleccionModelo();
        }

        public void ValidarSeleccionModelo()
        {
            erReporte.Clear();
            if (radUnModelo.Checked)
            {
                if (cboModelo.SelectedIndex == 0)
                {
                    erReporte.SetError(cboModelo, "Seleccione un modelo");
                    Validado = false;
                }
            }
        }

        public void ValidarSeleccionFusor()
        {
            erReporte.Clear();
            if (rdUnFusor.Checked)
            {
                if (cboFusor.SelectedIndex == 0)
                {
                    erReporte.SetError(cboFusor, "Seleccione un fusor");
                    Validado = false;
                }
            }
        }
        #endregion

        #region OpcionesReportes
        private void cboOpcionReporte_SelectedIndexChanged(object sender, EventArgs e)
        {
            btnGenerarReporte.Enabled = true;
            TipoBusqueda = cboOpcionReporte.SelectedItem.ToString();
            grpMarca.Visible = false;
            switch (TipoBusqueda)
            {
                case "Serie":
                    MostrarComboBox("SeleccionarSeriesEquipos", true); break;
                case "Fecha":
                    chkRango.Visible = false;
                    chkRango.Checked= true;
                    grpMarca.Visible = true;
                    break;
                case "Fusor":
                    grpFusor.Visible = true;
                    chkRango.Visible = true;
                    rdTodosLosFusores.Checked = true;
                    break;
                case "Todos Los Fusores":
                    TipoBusqueda = "Fusor";
                    chkRango.Visible = true; break;
                case "Cliente":
                    MostrarComboBox("SeleccionarClientesServicios", true);
                    grpMarca.Visible = true;//Mostramos opciones de las marcas
                    TipoBusqueda = "Cliente";
                    break;
            }
        }

        //Metodo que nos ayudara a llenar el combobox, dependiendo el stored procedure que le pasemos y si queremos tener o no un rango de fechas definido
        private void MostrarComboBox(string sp, bool MostrarRango)
        {
            cboOpcionElegida.Visible = true;
            LlenarComboBox(cboOpcionElegida, sp, 0);

            chkRango.Visible = MostrarRango;
            //erReporte.Clear();
        }

        private void ReiniciarControles()
        {
            MostrarRangosFechas(false);
            cboOpcionElegida.Visible = false;
            grpFusor.Visible = false;
            chkRango.Visible = false;
            chkRango.Checked = false;
            Cliente = "";
            TipoBusqueda = "";
            Parametro = "";
            TipoBusquedaAdicional = "";
            grpMarca.Visible = false;
            grpModelo.Visible = false;
        }

        private void MostrarRangosFechas(bool mostrar)
        {
            lblFechaFinal.Visible = mostrar;
            lblFechaInicio.Visible = mostrar;
            dtpFechaFinal.Visible = mostrar;
            dtpFechaInicial.Visible = mostrar;
            chkRango.Checked = mostrar;
        }

        public void MostrarOpcionesMarca(bool mostrar)
        {
            rdTodasLasMarcas.Visible = mostrar;
            rdUnaMarca.Visible = mostrar;
        }

        private void rdTodasLasMarcas_CheckedChanged(object sender, EventArgs e)
        {
            cboMarca.Visible = false;
            grpModelo.Visible = false;
            rdTodosLosModelos.Checked = true;
            TipoBusqueda = cboOpcionReporte.SelectedItem.ToString();
        }

        //Evento que nos llena de las marcas con las que contamos para poder elegir alguna
        private void rdUnaMarca_CheckedChanged(object sender, EventArgs e)
        {
            cboMarca.Visible = true;
            rdTodosLosModelos.Checked = true;
            LlenarComboBox(cboMarca, "SeleccionarMarca", 0);

            VerificarTipoBusquedaCliente();
        }

        public void VerificarTipoBusquedaCliente()
        {
            if (TipoBusqueda == "Cliente")
            {
                TipoBusquedaAdicional = "CLIENTE";
                Cliente = cboOpcionElegida.SelectedItem.ToString();
                TipoBusqueda = "Marca";
            }
            else
            {
                TipoBusqueda = "Marca";
            }
        }

        //Evento dependiendo si escogemos una opcion, nos mostrara las ocpiones para elegir uno o todos los modelos
        private void cboMarca_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cboMarca.SelectedItem.ToString() != " ")
            {
                grpModelo.Visible = true;
            }
        }

        //Evento que nos ayudara a llenar los modelos en el combo box dependiendo la marca que se haya seleccionado
        private void radUnModelo_CheckedChanged(object sender, EventArgs e)
        {
            cboModelo.Visible = true;
            int IdMarca = NuevaAccion.BuscarId(cboMarca.SelectedItem.ToString(), "ObtenerIdMarca");
            LlenarComboBox(cboModelo, "SeleccionarModelos", IdMarca);
            TipoBusqueda = "Modelo";
        }

        //Evento que nos ayudara a mostrar la fecha incicial y final, en dado caso que asi se requiera en el inventario
        private void chkRango_CheckedChanged(object sender, EventArgs e)
        {
            if (chkRango.Checked)
                MostrarRangosFechas(true);
            else
                MostrarRangosFechas(false);
        }

        //Evento para que en dado caso de que este seleccionado, no mostrará los modelos
        private void rdTodosLosModelos_CheckedChanged(object sender, EventArgs e)
        {
            cboModelo.Visible = false;
            VerificarTipoBusquedaCliente();
        }
        private void rdUnFusor_CheckedChanged(object sender, EventArgs e)
        {
            cboFusor.Visible = true;
            LlenarComboBox(cboFusor, "SeleccionarFusores", 0);
            chkRango.Visible = false;
            TipoBusqueda = "Fusor";
        }

        private void rdTodosLosFusores_CheckedChanged(object sender, EventArgs e)
        {
            cboFusor.Visible = false;
            chkRango.Visible = true;
            TipoBusqueda = "Fusores";
        }

        private void cboOpcionElegida_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cboOpcionElegida.SelectedItem.ToString() != "")
            {
                rdTodasLasMarcas.Checked = true;
            }
        }
        #endregion

        #region PanelSuperior
        private void btnCerrar_Click(object sender, EventArgs e)
        {
            this.Close();
        }
        #endregion

        private void btnGenerarReporte_Click(object sender, EventArgs e)
        {
            if (!DeterminarTipoValidacionDatos())
                return;

            try
            {
                VerificarParametroBusqueda();
                Reporte nuevoReporte = new Reporte()
                {
                    FechaInicio = dtpFechaInicial.Value,
                    FechaFinal = dtpFechaFinal.Value,
                    RangoHabilitado = chkRango.Checked,
                    TipoBusqueda = TipoBusqueda,
                    TipoBusquedaAdicional = TipoBusquedaAdicional,
                    ParametroBusqueda = Parametro
                };
                
                MessageBox.Show(nuevoReporte.TipoBusqueda);
                bool DatosEncontrados = LgReporteServicio.DeterminarTipoReporte(nuevoReporte, Cliente);
                ReiniciarControles();

                if (!DatosEncontrados)
                    MessageBox.Show("¡NO SE ENCONTRARON REGISTROS!", "DATOS NO ENCONTRADOS", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch(Exception ex)
            {
                MessageBox.Show("Ocurrio un error " + ex.Message);
            }
        }

        public void VerificarParametroBusqueda()
        {
            switch (TipoBusqueda)
            {
                case "Serie":
                    Parametro = cboOpcionElegida.SelectedItem.ToString();
                    break;
                case "Fusor":
                    VerificarFusorEspecifico();
                    break;
                case "Cliente":
                    Parametro = cboOpcionElegida.SelectedItem.ToString();
                    Cliente = Parametro;
                    VerificarMarcaEspecifica();
                    break;
                case "Fecha":
                    VerificarMarcaEspecifica();
                    break;
                default:
                    VerificarMarcaEspecifica(); break;
            }
        }


        //Obtenemos el parametro de busqueda para el reporte ya sea una marca, fusor o modelo especifico
        public void VerificarFusorEspecifico()
        {
            if (rdUnFusor.Checked)
            {
                Parametro = cboFusor.SelectedItem.ToString();
            }
        }

        public void VerificarMarcaEspecifica()
        {
            //Cliente = cboOpcionElegida.SelectedItem.ToString();
            if (rdUnaMarca.Checked)
            {
                Parametro = cboMarca.SelectedItem.ToString();
            }
            if (radUnModelo.Checked)
            {
                Parametro = cboModelo.SelectedItem.ToString();
            }
        }

    }

}
