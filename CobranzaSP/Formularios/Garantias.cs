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

namespace CobranzaSP.Formularios
{
    public partial class Garantias : Form
    {
        public Garantias()
        {
            InitializeComponent();
            InicioAplicacion();
        }
        FuncionesFormularios Formulario = new FuncionesFormularios();
        AccionesLógica NuevaAccion = new AccionesLógica();
        Garantia GarantiaSeleccionada;

        #region Inicio
        public void InicioAplicacion()
        {
            Formulario.PropiedadesDtg(dtgGarantias);
            MostrarGarantias();
            HabilitarBotones(false);
        }

        public void MostrarGarantias()
        {
            //Limpiamos los datos del datagridview
            dtgGarantias.DataSource = null;
            dtgGarantias.Refresh();
            DataTable tabla = new DataTable();
            //Guardamos los registros dependiendo la consulta
            tabla = NuevaAccion.Mostrar("MostrarGarantias");
            //Asignamos los registros que optuvimos al datagridview
            dtgGarantias.DataSource = tabla;
            dtgGarantias.Columns["IdGarantia"].Visible = false;

        }

        public void HabilitarBotones(bool Habilitado)
        {
            btnModificarGarantia.Enabled = Habilitado;
            btnEliminar.Enabled = Habilitado;
            btnCancelar.Enabled = Habilitado;
        }
        #endregion

        #region Validaciones
        #endregion

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

        #region Botones
        private void btnAgregar_Click(object sender, EventArgs e)
        {
            NuevaGarantia Garantia = new NuevaGarantia(this);
            Garantia.Show();
        }
        #endregion

        #region Eventos
        #endregion

        #region Metodos
        #endregion

        private void dtgGarantias_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            HabilitarBotones(true);
            GarantiaSeleccionada = new Garantia()
            {
                Id = int.Parse(dtgGarantias.CurrentRow.Cells[0].Value.ToString()),
                Marca = dtgGarantias.CurrentRow.Cells[1].Value.ToString(),
                Modelo = dtgGarantias.CurrentRow.Cells[2].Value.ToString(),
                Cliente = dtgGarantias.CurrentRow.Cells[3].Value.ToString(),
                Fecha = DateTime.Parse(dtgGarantias.CurrentRow.Cells[4].Value.ToString()),
                Cantidad = int.Parse(dtgGarantias.CurrentRow.Cells[5].Value.ToString()),
                Descripcion = dtgGarantias.CurrentRow.Cells[6].Value.ToString()
            };
        }

        private void btnCancelar_Click(object sender, EventArgs e)
        {
            HabilitarBotones(false);
        }

        private void btnModificarGarantia_Click(object sender, EventArgs e)
        {
            NuevaGarantia Garantia = new NuevaGarantia(this,GarantiaSeleccionada);
            Garantia.Show();
        }

        private void btnEliminar_Click(object sender, EventArgs e)
        {
            try
            {
                if (MessageBox.Show("¿Esta seguro de eliminar la garantía?", "CONFIRME ELIMINACION", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                {
                    MessageBox.Show("Eliminacion cancelada!!", "CANCELADO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                NuevaAccion.Eliminar(GarantiaSeleccionada.Id, "EliminarGarantia");
                MessageBox.Show("Se elimino la garantía", "OPERACION EXITOSA", MessageBoxButtons.OK, MessageBoxIcon.Information);
                MostrarGarantias();
                HabilitarBotones(false);
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void btnAbrirReportes_Click(object sender, EventArgs e)
        {
            Formulario.AbrirForm(new ReportesGarantias());
        }
    }
}
