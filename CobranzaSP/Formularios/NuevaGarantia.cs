using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data;
using System.Data.SqlClient;
using CobranzaSP.Lógica;
using CobranzaSP.Modelos;
using System.Runtime.InteropServices;
using DocumentFormat.OpenXml.Office2010.Excel;

namespace CobranzaSP.Formularios
{
    public partial class NuevaGarantia : Form
    {
        Garantias Garantia = new Garantias();
        Garantia GarantiaSeleccionada = new Garantia();

        int IdGarantia = 0;
        bool ModificarGarantia = false;
        public NuevaGarantia(Garantias NuevaGarantia)
        {
            InitializeComponent();
            this.Garantia = NuevaGarantia;
            ModificarGarantia = false;
            InicioAplicacion();
        }

        public NuevaGarantia(Garantias NuevaGarantia, Garantia GarantiaSeleccionada)
        {
            InitializeComponent();
            this.Garantia = NuevaGarantia;
            this.GarantiaSeleccionada = GarantiaSeleccionada;
            InicioAplicacion();
            LlenarDatosGarantia();
            ModificarGarantia = true;
        }

        public void LlenarDatosGarantia()
        {
            dtpFechaRegistro.Value = GarantiaSeleccionada.Fecha;
            cboMarcas.SelectedItem = GarantiaSeleccionada.Marca;
            cboModelos.SelectedItem = GarantiaSeleccionada.Modelo;
            cboClientes.SelectedItem = GarantiaSeleccionada.Cliente;
            txtCantidad.Text = GarantiaSeleccionada.Cantidad.ToString();
            rtxtDescripcion.Text = GarantiaSeleccionada.Descripcion;
            IdGarantia = GarantiaSeleccionada.Id;
        }

        FuncionesFormularios FuncionFormulario = new FuncionesFormularios();
        AccionesLógica NuevaAccion = new AccionesLógica();
        LogicaGarantia lgGarantia = new LogicaGarantia();
        bool BuscandoRegistro = false;


        public void InicioAplicacion()
        {
            FuncionFormulario.LlenarComboBox(cboClientes, "SeleccionarClientesServicios", 0);
            FuncionFormulario.LlenarComboBox(cboMarcas, "SeleccionarMarca", 0);
            //FuncionFormulario.LlenarComboBox(cboModelos, "SeleccionarCartuchos", 0);
        }

        private void label12_Click(object sender, EventArgs e)
        {

        }

        private void btnGuardar_Click(object sender, EventArgs e)
        {
            //if (!ValidarRegistros())
            //    return;

            int IdMarcaEncontrada = NuevaAccion.BuscarId(cboMarcas.SelectedItem.ToString(), "ObtenerIdMarca");
            Garantia NuevaGarantia = new Garantia()
            { 
                Id = IdGarantia,
                IdCliente = NuevaAccion.BuscarId(cboClientes.SelectedItem.ToString(), "ObtenerIdCliente"),
                IdMarca = IdMarcaEncontrada,
                IdCartucho = NuevaAccion.BuscarIdCartucho(cboModelos.SelectedItem.ToString(), IdMarcaEncontrada, "ObtenerIdCartucho"),
                Cantidad = int.Parse(txtCantidad.Text),
                Fecha = dtpFechaRegistro.Value,
                Descripcion = rtxtDescripcion.Text
            };
            if (ModificarGarantia)
            {
                if (MessageBox.Show("¿Esta seguro de modificar la garantía?", "CONFIRME LA MODIFICACION", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                {
                    MessageBox.Show("Modificacion cancelada!!", "CANCELADO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                string Mensaje = lgGarantia.RegistrarGarantia(NuevaGarantia, "ModificarGarantia");
                MessageBox.Show(Mensaje, "REGISTRO GARANTIA", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                string Mensaje = lgGarantia.RegistrarGarantia(NuevaGarantia, "AgregarGarantia");
                MessageBox.Show(Mensaje, "REGISTRO GARANTIA", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            Garantia.MostrarGarantias();
            this.Close();
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
        private void btnCerrar_Click(object sender, EventArgs e)
        {
            this.Close();
        }
        #endregion


        #region Eventos
        private void cboMarcas_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            if (cboMarcas.SelectedItem.ToString() != " ")
            {
                int IdMarca = NuevaAccion.BuscarId(cboMarcas.SelectedItem.ToString(), "ObtenerIdMarca");
                FuncionFormulario.LlenarComboBoxModelos(cboModelos, "SeleccionarCartuchos", IdMarca);
            }
        }
        private void cboModelos_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
        #endregion

    }
}
