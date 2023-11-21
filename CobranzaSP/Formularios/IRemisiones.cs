using CobranzaSP.Lógica;
using CobranzaSP.Modelos;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CobranzaSP.Formularios
{
    public partial class IRemisiones : Form
    {
        public IRemisiones()
        {
            InitializeComponent();
            InicioAplicacion();
        }
        bool Remision = true;
        bool Modificando = false;
        int IdVentaMostrador = 0;
        CD_Conexion cn = new CD_Conexion();
        AccionesLógica NuevaAccion = new AccionesLógica();
        RemisionLogica RemisionNueva = new RemisionLogica();
        VentaMostradorLogica nuevaVentaMostrador = new VentaMostradorLogica();
        #region Inicio

        public void InicioAplicacion()
        {
            BotonesDesactivados(false);
            PropiedadesDtgCobranza(dtgRemisiones);
            LlenarComboBox(cboClientes, "SeleccionarClientes", 0);
            MostrarDatos(dtgRemisiones, "MostrarRemisiones");
            radRemisiones.Checked= true;
        }

        public void OcultarControlesFormulario(bool mostrar)
        {
            lblDiasCredito.Visible = mostrar;
            lblRemisiones.Visible = mostrar;
            txtFolio.Visible = mostrar;
            txtDiasCredito.Visible = mostrar;
            Remision = mostrar;
        }

        public void BotonesDesactivados(bool activado)
        {
            btnCancelar.Enabled = activado;
            btnEliminar.Enabled = activado;
            btnCobrado.Enabled = activado;
        }

        public void PropiedadesDtgCobranza(DataGridView dtg)
        {
            //Solo lectura
            dtg.ReadOnly = true;

            //No agregar renglones
            dtg.AllowUserToAddRows = false;

            //No borrar renglones
            dtg.AllowUserToDeleteRows = false;

            //Ajustar automaticamente el ancho de las columnas
            dtg.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            //dtgServicios.AutoResizeColumns(DataGridViewAutoSizeColumnsMo‌​de.Fill);
            dtg.AutoResizeColumns();
        }

        public void MostrarDatos(DataGridView dtg, string sp)
        {
            //Limpiamos los datos del datagridview
            dtg.DataSource = null;
            dtg.Refresh();
            DataTable tabla = new DataTable();
            //Guardamos los registros dependiendo la consulta
            tabla = NuevaAccion.Mostrar(sp);
            //Asignamos los registros que optuvimos al datagridview
            dtg.DataSource = tabla;
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


        private void radRemisiones_CheckedChanged(object sender, EventArgs e)
        {
            OcultarControlesFormulario(true);
        }

        private void radMostrador_CheckedChanged(object sender, EventArgs e)
        {
            OcultarControlesFormulario(false);
        }
        #endregion

        #region Validaciones
        public bool ValidarRemision()
        {
            erRemisiones.Clear();
            bool Validado = true;
            if (Remision)
            {
                if (txtFolio.Text == "")
                {
                    erRemisiones.SetError(txtFolio, "Ingrese Folio"); 
                    Validado = false;
                }
                    
                if (txtDiasCredito.Text == "")
                {
                    erRemisiones.SetError(txtDiasCredito, "Ingrese Dias");
                    Validado = false;
                }
                    
            }

            if (txtCantidad.Text == "")
            {
                erRemisiones.SetError(txtCantidad, "Ingrese Cantidad"); 
                Validado = false;
            }
                
            if(cboClientes.Text == " ")
            {
                erRemisiones.SetError(cboClientes, "Ingrese cliente"); 
                Validado = false;
            }
                

            return Validado;
        }
        private void txtFolio_KeyPress(object sender, KeyPressEventArgs e)
        {
            Validacion.SoloLetrasYNumeros(e);
        }

        private void txtDiasCredito_KeyPress(object sender, KeyPressEventArgs e)
        {
            Validacion.SoloNumeros(e);
        }

        private void txtCantidad_KeyPress(object sender, KeyPressEventArgs e)
        {
            
        }
        #endregion


        #region Botones

        private void btnGuardar_Click(object sender, EventArgs e)
        {
            try
            {
                if (!ValidarRemision())
                    return;

                if (Remision)
                {
                    GuardarRemision();
                }
                //Entonces agregaremos una venta de mostrador
                else
                {
                    GuardarVentaMostrador();
                }
                LimpiarForm();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public void GuardarRemision()
        {
            string Mensaje;
            Remision nuevaRemision = new Remision()
            {
                NumeroFolio = int.Parse(txtFolio.Text),
                DiasCredito = int.Parse(txtDiasCredito.Text),
                Cantidad = double.Parse(txtCantidad.Text),
                IdCliente = NuevaAccion.BuscarId(cboClientes.SelectedItem.ToString(), "BuscarIdCliente"),
                FechaPago = dtpFechaRemision.Value
            };

            if (Modificando)
            {
                if (MessageBox.Show("¿Esta seguro de modificar el registro?", "CONFIRME LA MODIFICACION", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                {
                    MessageBox.Show("!!Modificación cancelada!!", "OPERACION CANCELADA", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    LimpiarForm();
                    return;
                }
                Mensaje = RemisionNueva.AgregarRemision(nuevaRemision, "ModificarRemision", false);
            }
            else
            {
                Mensaje = RemisionNueva.AgregarRemision(nuevaRemision, "AgregarRemision", true);
            }
            MessageBox.Show(Mensaje, "REGISTRO DE DATOS", MessageBoxButtons.OK, MessageBoxIcon.Information);
            MostrarDatos(dtgRemisiones, "MostrarRemisiones");
        }

        public void GuardarVentaMostrador()
        {
            string Mensaje;
            VentaMostrador nuevaVenta = new VentaMostrador()
            {
                IdVenta = IdVentaMostrador,
                IdCliente = NuevaAccion.BuscarId(cboClientes.SelectedItem.ToString(), "BuscarIdCliente"),
                Cantidad = double.Parse(txtCantidad.Text),
                FechaPago = dtpFechaRemision.Value
            };
            Mensaje = nuevaVentaMostrador.AgregarVentaMostrador(nuevaVenta);
            
            MessageBox.Show(Mensaje, "REGISTRO NUEVA VENTA", MessageBoxButtons.OK, MessageBoxIcon.Information);
            LimpiarForm();
        }

        private void btnEnviarCorreo_Click(object sender, EventArgs e)
        {
            MessageBox.Show(RemisionNueva.GenerarCorreo(cboClientes.SelectedItem.ToString()), "ENVIO DE CORREO", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void btnCobrado_Click(object sender, EventArgs e)
        {
            try
            {
                if (MessageBox.Show("¿Esta seguro de que la remision fue pagada?", "CONFIRME EL PAGO", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                {
                    MessageBox.Show("!!REMISION NO PAGADA!!", "OPERACION CANCELADA", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    LimpiarForm();
                    return;
                }
                CuentaPagada nuevaCuentaPagada = new CuentaPagada()
                {
                    IdCliente = NuevaAccion.BuscarId(cboClientes.SelectedItem.ToString(), "BuscarIdCliente"),
                    Factura = txtFolio.Text,
                    Cantidad = double.Parse(txtCantidad.Text.Replace(",", "")),
                    FechaFactura = dtpFechaRemision.Value
                };
                MessageBox.Show(RemisionNueva.RemisionPagada(nuevaCuentaPagada), "AGREGANDO NUEVO COBRO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                MostrarDatos(dtgRemisiones, "MostrarRemisiones");
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
                string Mensaje;
                if (MessageBox.Show("¿Esta seguro de eliminar el registro?", "CONFIRME LA ELIMINACION", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                {
                    MessageBox.Show("!!Eliminación cancelada!!", "OPERACION CANCELADA", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    LimpiarForm();
                    return;
                }
                Mensaje = NuevaAccion.Eliminar(int.Parse(txtFolio.Text), "EliminarRemision");
                MostrarDatos(dtgRemisiones, "MostrarRemisiones");

                MessageBox.Show(Mensaje, "ELIMINACION EXITOSA", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
            BotonesDesactivados(false);
            LimpiarForm();
        }

        #endregion

        public void LimpiarForm()
        {
            foreach (Control c in grpDatos.Controls)
            {
                if (c is TextBox || c is RichTextBox)
                {
                    c.Text = "";
                }
            }
            dtpFechaRemision.Value = DateTime.Now;
            cboClientes.SelectedIndex = 0;
        }


        #region Eventos
        

        private void dtgRemisiones_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            Modificando = true;
            Remision = true;
            //Activamos los botones para eliminar o cancelar
            BotonesDesactivados(true);
            //Ocultamos los controles que no necesitamos
            OcultarControlesFormulario(true);
            txtFolio.Text = dtgRemisiones.CurrentRow.Cells[0].Value.ToString();
            cboClientes.SelectedItem = dtgRemisiones.CurrentRow.Cells[1].Value.ToString();
            txtDiasCredito.Text = dtgRemisiones.CurrentRow.Cells[2].Value.ToString();
            txtCantidad.Text = dtgRemisiones.CurrentRow.Cells[3].Value.ToString().Replace("$", "");
            dtpFechaRemision.Value = DateTime.Parse(dtgRemisiones.CurrentRow.Cells[4].Value.ToString());
        }
        #endregion

    }
}
