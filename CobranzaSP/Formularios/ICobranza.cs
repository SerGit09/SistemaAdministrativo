using System;
using System.Collections.Generic;
using CobranzaSP.Lógica;
using CobranzaSP.Modelos;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using DocumentFormat.OpenXml.Vml.Spreadsheet;

namespace CobranzaSP.Formularios
{
    public partial class ICobranza : Form
    {
        public ICobranza()
        {
            InitializeComponent();
            InicioAplicacion();
        }

        CD_Conexion cn = new CD_Conexion();
        AccionesLógica NuevaAccion = new AccionesLógica();
        CobranzaLógica nuevoCobro = new CobranzaLógica();
        bool Modificando = false;
        int Id;

        #region Inicio
        public void InicioAplicacion()
        {
            string[] Opciones = { "", "SI", "NO"};
            cboPromesaPago.Items.AddRange(Opciones);
            BotonesDesactivados(false);

            PropiedadesDtgCobranza();
            MostrarDatosCobranza();
            LlenarComboBox(cboClientes, "SeleccionarClientes", 0);

            ActualizarTotalCobranza();
            cboPromesaPago.DropDownStyle = ComboBoxStyle.DropDownList;
        }

        public void BotonesDesactivados(bool activado)
        {
            btnCancelar.Enabled = activado;
            btnEnviarCorreo.Enabled = activado;
            btnEliminar.Enabled = activado;
            btnCobrado.Enabled = activado;
        }

        public void PropiedadesDtgCobranza()
        {
            //Solo lectura
            dtgCobranza.ReadOnly = true;

            //No agregar renglones
            dtgCobranza.AllowUserToAddRows = false;

            //No borrar renglones
            dtgCobranza.AllowUserToDeleteRows = false;

            //Ajustar automaticamente el ancho de las columnas
            dtgCobranza.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            //dtgServicios.AutoResizeColumns(DataGridViewAutoSizeColumnsMo‌​de.Fill);
            dtgCobranza.AutoResizeColumns();
        }

        public void LlenarComboBox(ComboBox cb, string sp, int Marca)
        {
            SqlDataReader dr;
            cb.Items.Clear();
            dr = NuevaAccion.LlenarComboBox(sp);


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

        public void MostrarDatosCobranza()
        {
            //Limpiamos los datos del datagridview
            dtgCobranza.DataSource = null;
            dtgCobranza.Refresh();
            DataTable tabla = new DataTable();
            //Guardamos los registros dependiendo la consulta
            tabla = NuevaAccion.Mostrar("MostrarCobranza");
            //Asignamos los registros que optuvimos al datagridview
            dtgCobranza.DataSource = tabla;
        }
        #endregion

        #region Validaciones

        public bool ValidarCampos()
        {
            bool Validado = true;
            erCobranza.Clear();
            foreach (Control c in grpDatos.Controls)
            {
                if (c is ComboBox || c is TextBox)
                {
                    if (c.Text == "" || c.Text == " ")
                    {
                        erCobranza.SetError(c, "Campo Obligatorio");
                        Validado = false;
                    }
                }
            }
            return Validado;
        }

        private void txtFactura_KeyPress(object sender, KeyPressEventArgs e)
        {
            Validacion.SoloLetrasYNumeros(e);
        }

        private void txtCantidad_KeyPress(object sender, KeyPressEventArgs e)
        {
            //Validacion.SoloNumeros(e);
        }
        #endregion

        #region Botones

        private void btnGuardar_Click(object sender, EventArgs e)
        {
            string FormaPago = "";
            bool FacturaDuplicada = false;
            if (!ValidarCampos())
                return;
            FormaPago = (nuevoCobro.MostrarDiasCredito(cboClientes.SelectedItem.ToString()) == 0 || chkPue.Checked) ? "PUE" : "";
            Cobranza nuevaCobranza = new Cobranza()
            {
                IdCliente = NuevaAccion.BuscarId(cboClientes.SelectedItem.ToString(), "BuscarIdCliente"),
                NumeroFactura = txtFactura.Text,
                DiasCredito = nuevoCobro.MostrarDiasCredito(cboClientes.SelectedItem.ToString()),
                Cantidad = double.Parse(txtCantidad.Text.Replace(",", "")),
                FormaPago = FormaPago,
                FechaFactura = dtpFechaFactura.Value,
                Observaciones = rtxtObservaciones.Text,
                PromesaPago = cboPromesaPago.SelectedItem.ToString()
            };
            if (Modificando)
            {
                if (MessageBox.Show("¿Esta seguro de modificar el registro?", "CONFIRME LA MODIFICACION", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                {
                    MessageBox.Show("!!Modificación cancelada!!", "OPERACION CANCELADA", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    LimpiarForm();
                    return;
                }
                MessageBox.Show(nuevoCobro.ModificarCobro(nuevaCobranza, Id), "MODIFICACION DE COBRO", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                //Evitamos duplicados de facturas
                FacturaDuplicada = NuevaAccion.VerificarDuplicados(nuevaCobranza.NumeroFactura, "VerificarFacturaDuplicada");
                //En esta parte deberemos de comprobar si existe la cuenta de dicho cliente, para saber si la agregamos o simplemente la dejamos ahi
                if (FacturaDuplicada)
                {
                    MessageBox.Show("Ingrese un numero de factura distinto", "NUMERO DE FACTURA YA EXISTENTE", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    MessageBox.Show(nuevoCobro.Registrar(nuevaCobranza), "AGREGANDO NUEVO COBRO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            ActualizarTotalCobranza();
            MostrarDatosCobranza();
            LimpiarForm();
        }
        private void btnCobrado_Click(object sender, EventArgs e)
        {
            try
            {
                if (MessageBox.Show("¿Esta seguro de que la factura fue pagada?", "CONFIRME EL PAGO", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                {
                    MessageBox.Show("!!CUENTA NO PAGADA!!", "OPERACION CANCELADA", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    LimpiarForm();
                    return;
                }
                CuentaPagada nuevaCuentaPagada = new CuentaPagada()
                {
                    IdCliente = NuevaAccion.BuscarId(cboClientes.SelectedItem.ToString(), "BuscarIdCliente"),
                    Factura = txtFactura.Text,
                    Cantidad = double.Parse(txtCantidad.Text.Replace(",", "")),
                    FechaFactura = dtpFechaFactura.Value,
                };
                MessageBox.Show(nuevoCobro.CuentaPagada(nuevaCuentaPagada, Id), "AGREGANDO NUEVO COBRO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                MostrarDatosCobranza();
                ActualizarTotalCobranza();
                LimpiarForm();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void btnEliminar_Click(object sender, EventArgs e)
        {
            try
            {
                if (MessageBox.Show("¿Esta seguro de eliminar el registro?", "CONFIRME LA ELIMINACION", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                {
                    MessageBox.Show("!!Eliminacion cancelada!!", "OPERACION CANCELADA", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    LimpiarForm();
                    return;
                }
                MessageBox.Show(NuevaAccion.Eliminar(Id, "EliminarCobro"));
                ActualizarTotalCobranza();
                MostrarDatosCobranza();
                LimpiarForm();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void btnCancelar_Click(object sender, EventArgs e)
        {
            Modificando = false;
            LimpiarForm();
            BotonesDesactivados(false);
        }

        private void btnEnviarCorreo_Click(object sender, EventArgs e)
        { 
            MessageBox.Show(nuevoCobro.GenerarCorreo(cboClientes.SelectedItem.ToString()), "ENVIO DE CORREO", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void btnBusqueda_Click(object sender, EventArgs e)
        {
            LimpiarForm();
            SqlDataReader dr;
            if (txtBusqueda.Text != "")
            {
                dr = NuevaAccion.Buscar(txtBusqueda.Text, "BuscarFactura");
                if (dr.Read())
                {
                    txtFactura.Text = dr[1].ToString();
                    cboClientes.SelectedItem = dr[2].ToString();
                    txtCantidad.Text = dr[5].ToString();
                    dtpFechaFactura.Value = Convert.ToDateTime(dr[6].ToString());
                    rtxtObservaciones.Text = dr[9].ToString();
                    cboPromesaPago.SelectedItem = dr[12].ToString();
                    dr.Close();
                }
                else
                {
                    MessageBox.Show("El numero de factura no existe", "REGISTRO NO ENCONTRADO", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    dr.Close();
                }
            }
        }
        #endregion

        #region Eventos
        private void dtgCobranza_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            Modificando = true;
            BotonesDesactivados(true);
            Id = int.Parse(dtgCobranza.CurrentRow.Cells[0].Value.ToString());
            txtFactura.Text = dtgCobranza.CurrentRow.Cells[1].Value.ToString();
            cboClientes.SelectedItem = dtgCobranza.CurrentRow.Cells[2].Value.ToString();
            dtpFechaFactura.Value = Convert.ToDateTime(dtgCobranza.CurrentRow.Cells[6].Value.ToString());
            txtCantidad.Text = dtgCobranza.CurrentRow.Cells[5].Value.ToString().Replace("$", "");
            rtxtObservaciones.Text = dtgCobranza.CurrentRow.Cells[8].Value.ToString();
            cboPromesaPago.SelectedItem = dtgCobranza.CurrentRow.Cells[9].Value.ToString();
            string Estado = dtgCobranza.CurrentRow.Cells[10].Value.ToString();
            if(Estado == "PUE")
                chkPue.Checked = true;
            else
                chkPue.Checked = false;

        }

        private void dtgCobranza_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (this.dtgCobranza.Columns[e.ColumnIndex].Name == "Estado")
            {
                switch (e.Value.ToString())
                {
                    case "PUE":
                        e.CellStyle.ForeColor = Color.Black;
                        e.CellStyle.BackColor = Color.Yellow; break;
                    case "Disponible":
                        e.CellStyle.ForeColor = Color.Black;
                        e.CellStyle.BackColor = Color.Green; break;
                    case "Vencida":
                        e.CellStyle.ForeColor = Color.White;
                        e.CellStyle.BackColor = Color.Red; break;
                }
            }
        }
        #endregion

        #region Metodos Locales
        private void ActualizarTotalCobranza()
        {
            SqlCommand comando = new SqlCommand();
            comando.Connection = cn.AbrirConexion();
            comando.CommandText = "ActualizarDatos";
            comando.CommandType = CommandType.StoredProcedure;
            //Mandamos la variable y especificamos el tipo de dato sql que sera
            SqlParameter total = new SqlParameter("@totalCobranza", SqlDbType.Money); total.Direction = ParameterDirection.Output;
            SqlParameter totalPue = new SqlParameter("@totalPue", SqlDbType.Money); totalPue.Direction = ParameterDirection.Output;

            comando.Parameters.Add(total);
            comando.Parameters.Add(totalPue);
            //comando.CommandText = "ObtenerTotales";

            comando.ExecuteNonQuery();
            //lblTotalCobrado.Text = comando.Parameters["@totalCobranza"].Value.ToString();
            decimal ValorRecbido = decimal.Parse(comando.Parameters["@totalCobranza"].Value.ToString());
            lblTotalCobrado.Text = String.Format("{0:c}", ValorRecbido);
            //lblTotalCobrado.Text = ValorRecbido +"";
            decimal Pue = decimal.Parse(comando.Parameters["@totalPue"].Value.ToString());
            lblTotalPue.Text = String.Format("{0:c}", Pue);

            cn.CerrarConexion();
        }

        public void LimpiarForm()
        {
            foreach (Control c in grpDatos.Controls)
            {
                if (c is TextBox || c is RichTextBox)
                {
                    c.Text = "";
                }
            }
            dtpFechaFactura.Value = DateTime.Now;
            cboClientes.SelectedIndex = 0;
            cboPromesaPago.SelectedIndex = 0;
            chkPue.Checked = false;
        }
        #endregion
    }
}
