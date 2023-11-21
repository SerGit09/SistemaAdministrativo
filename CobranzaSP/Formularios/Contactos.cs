using System;
using CobranzaSP.Lógica;
using CobranzaSP.Modelos;
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
    public partial class Contactos : Form
    {
        public Contactos()
        {
            InitializeComponent();
            InicioAplicacion();
        }
        CD_Conexion cn = new CD_Conexion();
        AccionesLógica NuevaAccion = new AccionesLógica();
        ContactoLogica nuevoContactoLogica = new ContactoLogica();
        SqlCommand comando = new SqlCommand();
        bool EstaModificando = false;
        bool NuevoCliente = false;
        int Id;

        #region Inicio
        public void InicioAplicacion()
        {
            BotonesDesactivados(false);
            MostrarClientesAdeudos();

            PropiedadesDtgCobranza(dtgCobranza);
            MostrarDatosCobranza();
            LlenarComboBox(cboClientes, "SeleccionarClientes", 0);
        }

        public void BotonesDesactivados(bool activado)
        {
            btnCancelar.Enabled = activado;
            btnEliminar.Enabled = activado;
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
            dtg.AutoResizeColumns();
            dtg.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
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
            tabla = NuevaAccion.Mostrar("MostrarContactos");
            //Asignamos los registros que optuvimos al datagridview
            dtgCobranza.DataSource = tabla;
        }

        public void MostrarClientesAdeudos()
        {
            SqlDataReader leer;
            comando.Connection = cn.AbrirConexion();
            //comando.CommandText = "ClientesAdeudos";
            comando.CommandText = "CuentasAdeudas";
            comando.CommandType = CommandType.StoredProcedure;

            leer = comando.ExecuteReader();
            while (leer.Read())
            {
                lstListaCorreos.Items.Add(leer[0].ToString());
            }
        }
        #endregion

        #region Validaciones
        public bool ValidarDatos()
        {
            bool Validado = true;
            erContacto.Clear();
            

            if (NuevoCliente)
            {
                Validado = Validar(txtDiasCredito);

                Validado = Validar(txtCliente);
            }
            else
            {
                Validado = Validar(cboClientes);
            }

            Validado = Validar(txtCorreo);

            return Validado;
        }

        public bool Validar(Control c)
        {
            if (string.IsNullOrWhiteSpace(c.Text))
            {
                erContacto.SetError(c, "Campo Obligatorio");
                return false;
            }

            return true;
            //if (c.Text == "" || c.Text == " ")
            //{
            //    return false;
            //} else
            //    return true;
        }
        #endregion

        #region Botones
        private void btnGuardar_Click(object sender, EventArgs e)
        {
            bool FacturaDuplicada = false;
            Contacto nuevoContacto = new Contacto()
            {
                IdCliente = NuevaAccion.BuscarId(cboClientes.SelectedItem.ToString(), "BuscarIdCliente"),
                Correo = txtCorreo.Text,
            };
            

            if (!ValidarDatos())
            {
                return;
            }

            if (NuevoCliente)
            {
                nuevoContacto.DiasCredito = int.Parse(txtDiasCredito.Text);
                nuevoContacto.Nombre = txtCliente.Text;
            }

            if (EstaModificando)
            {
                if (MessageBox.Show("¿Esta seguro de modificar el contacto?", "CONFIRME LA MODIFICACION", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                {
                    MessageBox.Show("!!Modificación cancelada!!", "OPERACION CANCELADA", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    LimpiarForm();
                    return;
                }
                MessageBox.Show(nuevoContactoLogica.ModificarContacto(nuevoContacto, Id), "MODIFICANDO CONTACTO", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show(nuevoContactoLogica.AgregarContacto(nuevoContacto,NuevoCliente), "AGREGANDO NUEVO CONTACTO", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            MostrarDatosCobranza();
            LimpiarForm();
        }
        private void btnEliminar_Click(object sender, EventArgs e)
        {
            try
            {
                if (MessageBox.Show("¿Esta seguro de eliminar el contacto?", "CONFIRME LA ELIMINACION", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                {
                    MessageBox.Show("!!Eliminacion cancelada!!", "OPERACION CANCELADA", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    LimpiarForm();
                    return;
                }
                MessageBox.Show(NuevaAccion.Eliminar(Id, "EliminarContacto"));
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
            EstaModificando = false;
            LimpiarForm();
            BotonesDesactivados(false);
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
            cboClientes.SelectedIndex = 0;
        }

        public void ActivarAgregarCliente(bool Activado)
        {
            if (Activado)
            {
                cboClientes.Visible = false;
                txtCliente.Visible = Activado;
            }
            else
            {
                cboClientes.Visible = true;
                txtCliente.Visible = false;
            }
        }
        #region Eventos
        private void dtgCobranza_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            EstaModificando = true;
            BotonesDesactivados(true);
            Id = int.Parse(dtgCobranza.CurrentRow.Cells[0].Value.ToString());
            cboClientes.SelectedItem = dtgCobranza.CurrentRow.Cells[1].Value.ToString();
            txtCorreo.Text = dtgCobranza.CurrentRow.Cells[02].Value.ToString();
        }
        #endregion

        private void chkNuevoCliente_CheckedChanged(object sender, EventArgs e)
        {
            if (!this.chkNuevoCliente.Checked)
            {
                cboClientes.Visible = true;
                HabilitarAgregarContacto(false);
            } 
            else
            {
                cboClientes.Visible = false;
                HabilitarAgregarContacto(true);
            }
        }

        public void HabilitarAgregarContacto(bool ClienteNuevo)
        {
            
            txtCliente.Visible = ClienteNuevo;
            NuevoCliente = ClienteNuevo;
            lblDias.Visible = ClienteNuevo;
            txtDiasCredito.Visible = ClienteNuevo;
            //NuevoCliente = ClienteNuevo;
        }
    }
}
