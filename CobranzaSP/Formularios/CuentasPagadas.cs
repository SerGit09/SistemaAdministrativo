using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using CobranzaSP.Lógica;
using CobranzaSP.Modelos;
using System.Data;
using System.Data.SqlClient;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace CobranzaSP.Formularios
{
    public partial class CuentasPagadas : Form
    {
        public CuentasPagadas()
        {
            InitializeComponent();
            InicioAplicacion();
        }

        CD_Conexion cn = new CD_Conexion();
        AccionesLógica NuevaAccion = new AccionesLógica();
        CuentasPagadasLogica AccionCuentasPagadas = new CuentasPagadasLogica();
        int Id = 0;

        #region Inicio
        public void InicioAplicacion()
        {
            PropiedadesDtgCobranza();
            MostrarDatosCuentasPagadas();
            ActualizarTotalCobrado();
            Meses();
            dtpAño.Format = DateTimePickerFormat.Custom;
            dtpAño.CustomFormat = "yyyy";
            dtpAño.ShowUpDown = true;
            //dtpAño.CustomFormat("yyyy");
            btnGuardar.Enabled = false;
        }

        public void PropiedadesDtgCobranza()
        {
            //Solo lectura
            dtgPagado.ReadOnly = true;

            //No agregar renglones
            dtgPagado.AllowUserToAddRows = false;

            //No borrar renglones
            dtgPagado.AllowUserToDeleteRows = false;

            //Ajustar automaticamente el ancho de las columnas
            dtgPagado.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            //dtgServicios.AutoResizeColumns(DataGridViewAutoSizeColumnsMo‌​de.Fill);
            dtgPagado.AutoResizeColumns();
        }

        public void Meses()
        {
            cboMeses.Items.Add("Enero");
            cboMeses.Items.Add("Febrero");
            cboMeses.Items.Add("Marzo");
            cboMeses.Items.Add("Abril");
            cboMeses.Items.Add("Mayo");
            cboMeses.Items.Add("Junio");
            cboMeses.Items.Add("Julio");
            cboMeses.Items.Add("Agosto");
            cboMeses.Items.Add("Septiembre");
            cboMeses.Items.Add("Octubre");
            cboMeses.Items.Add("Noviembre");
            cboMeses.Items.Add("Diciembre");
        }

        public void MostrarDatosCuentasPagadas()
        {
            //Limpiamos los datos del datagridview
            dtgPagado.DataSource = null;
            dtgPagado.Refresh();
            DataTable tabla = new DataTable();
            //Guardamos los registros dependiendo la consulta
            tabla = NuevaAccion.Mostrar("MostrarCuentasPagadas");
            //Asignamos los registros que optuvimos al datagridview
            dtgPagado.DataSource = tabla;
        }

        public void MostrarDatosMesElegido(int Mes, int Año)
        {
            //Limpiamos los datos del datagridview
            dtgPagado.DataSource = null;
            dtgPagado.Refresh();
            DataTable tabla = new DataTable();
            //Guardamos los registros dependiendo la consulta
            tabla = AccionCuentasPagadas.MostrarCuentasPagadasMesEspecifico(Mes, Año);
            //Asignamos los registros que optuvimos al datagridview
            dtgPagado.DataSource = tabla;
            if (tabla.Rows.Count <= 0)
            {
                MessageBox.Show("No se encontraron datos sobre dicho mes", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ActualizarTotalCobrado()
        {
            //CHECAR CUANDO SE QUEDE EN 0
            lblCobrado.Text = AccionCuentasPagadas.ActualizarCobranzaMesEspecifico(DateTime.Now.Month, DateTime.Now.Year);
        }
        #endregion

        #region Botones
        private void btnMostrar_Click(object sender, EventArgs e)
        {
            //VALIDAR QUE DICHO MES CONTENGA REGISTROS
            int Mes = cboMeses.SelectedIndex + 1;
            int Año = int.Parse(dtpAño.Value.Year.ToString());
            lblCobrado.Text = AccionCuentasPagadas.ActualizarCobranzaMesEspecifico(Mes, Año);
            MostrarDatosMesElegido(Mes, Año);
        }

        private void btnBusqueda_Click(object sender, EventArgs e)
        {

            SqlDataReader dr;
            if (txtBusqueda.Text != "")
            {
                dr = NuevaAccion.Buscar(txtBusqueda.Text, "BuscarCuentaPagada");
                if (dr.Read())
                {
                    DatosCuentasPagadas datos = new DatosCuentasPagadas();
                    datos.txtFactura.Text = dr[0].ToString();
                    datos.txtCliente.Text = dr[1].ToString();
                    datos.txtPrecio.Text = dr[2].ToString();
                    datos.dtFechaFactura.Value = Convert.ToDateTime(dr[3].ToString());
                    datos.dtFechaDePago.Value = Convert.ToDateTime(dr[4].ToString());
                    datos.txtTipoPago.Text = dr[5].ToString();
                    datos.Show();
                    dr.Close();
                }
                else
                {
                    MessageBox.Show("El numero de factura no existe", "REGISTRO NO ENCONTRADO", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    dr.Close();
                }
            }
        }

        private void btnGuardar_Click(object sender, EventArgs e)
        {
            try
            {
                CuentaPagada nuevaCuentaPagada = new CuentaPagada()
                {
                    Id = Id,
                    FechaPago = dtpFechaPago.Value
                };

                if (MessageBox.Show("¿Esta seguro la fecha de pago?", "CONFIRME LA MODIFICACION", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                {
                    MessageBox.Show("Modificacion cancelada!!", "CANCELADO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    ReiniciarControles();
                    return;
                }

                AccionCuentasPagadas.ModificarCuentaPagada(nuevaCuentaPagada);
                MessageBox.Show("Fecha de pago modificada correctamente");
                ReiniciarControles();
                MostrarDatosCuentasPagadas();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        #endregion

        #region MetodosLocales
        public void ReiniciarControles()
        {
            btnGuardar.Enabled = false;
            dtpFechaPago.Value = DateTime.Now;
            Id = 0;
        }
        #endregion

        #region Eventos
        private void dtgPagado_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            btnGuardar.Enabled = true;
            Id = int.Parse(dtgPagado.CurrentRow.Cells[0].Value.ToString());
            dtpFechaPago.Value = DateTime.Parse(dtgPagado.CurrentRow.Cells[5].Value.ToString());
        }
        #endregion
    }
}
