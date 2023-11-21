using CobranzaSP.Lógica;
using CobranzaSP.Modelos;
using DocumentFormat.OpenXml.Office2010.Excel;
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
    public partial class Compras : Form
    {
        public Compras()
        {
            InitializeComponent();
            IniciarAplicacion();
        }

        AccionesLógica NuevaAccion = new AccionesLógica();
        CD_Conexion cn = new CD_Conexion();
        LogicaCompras lgCompra = new LogicaCompras();
        bool Modificando = false;

        #region Inicio

        public void IniciarAplicacion()
        {
            LlenarComboBox(cboProveedores, "SeleccionarProveedores");

            PropiedadesDtg();
            ControlesDesactivados(false);

            Mostrar("MostrarCompras");
        }

        public void LlenarComboBox(ComboBox cb, string sp)
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

        public void PropiedadesDtg()
        {
            //Solo lectura
            dtgCompras.ReadOnly = true;

            //No agregar renglones
            dtgCompras.AllowUserToAddRows = false;

            //No borrar renglones
            dtgCompras.AllowUserToDeleteRows = false;

            //Ajustar automaticamente el ancho de las columnas
            dtgCompras.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            //dtgServicios.AutoResizeColumns(DataGridViewAutoSizeColumnsMo‌​de.Fill);
            dtgCompras.AutoResizeColumns();

            dtgCompras.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
        }

        public void ControlesDesactivados(bool Desactivado)
        {
            btnEliminar.Enabled = Desactivado;
            btnCancelar.Enabled = Desactivado;
            //btnMostrar.Enabled = Desactivado;
        }

        public void Mostrar(string sp)
        {
            //Limpiamos los datos del datagridview
            dtgCompras.DataSource = null;
            dtgCompras.Refresh();
            DataTable tabla = new DataTable();
            //Guardamos los registros dependiendo la consulta
            tabla = NuevaAccion.Mostrar(sp);
            //Asignamos los registros que optuvimos al datagridview
            dtgCompras.DataSource = tabla;
        }

        #endregion

        #region Validaciones

        public bool ValidarCampos()
        {
            bool Validado = true;
            erCompras.Clear();



            if (string.IsNullOrEmpty(txtFolio.Text))
            {
                erCompras.SetError(txtFolio, "Ingrese número de folio");
                Validado = false;
            }
            if (string.IsNullOrEmpty(txtCantidad.Text))
            {
                erCompras.SetError(txtCantidad, "Ingrese una cantidad");
                Validado = false;
            }

            if (chkNuevo.Checked)
            {
                if (string.IsNullOrEmpty(txtProveedor.Text))
                {
                    erCompras.SetError(txtProveedor, "Registre proveedor");
                    Validado = false;
                }
            }
            else
            {
                if (cboProveedores.SelectedIndex == 0)
                {
                    erCompras.SetError(cboProveedores, "Seleccione un proveedor");
                    Validado = false;
                }
            }
            return Validado;
        }

        private void txtFolio_KeyPress(object sender, KeyPressEventArgs e)
        {
            Validacion.SoloLetrasYNumeros(e);
        }

        private void txtCantidad_KeyPress(object sender, KeyPressEventArgs e)
        {
            Validacion.SoloNumeros(e);
        }

        private void txtProveedor_KeyPress(object sender, KeyPressEventArgs e)
        {
            Validacion.SoloLetrasYNumeros(e);
        }

        #endregion

        #region Botones
        private void btnGuardar_Click(object sender, EventArgs e)
        {
            try
            {
                bool FolioRepetido = false;
                if (!ValidarCampos())
                    return;

                Compra nuevaCompra = new Compra()
                {
                    Factura = txtFolio.Text,
                    Cantidad = double.Parse(txtCantidad.Text.Replace(",","").Replace("$","")),
                    Fecha = dtpFecha.Value
                };
                VerificarProveedor(nuevaCompra);

                if (Modificando)
                {
                    //Modificar stop procedure de modificar, para que se pueda cambiar la fecha
                    if (MessageBox.Show("¿Esta seguro de modificar el registro?", "CONFIRME LA MODIFICACION", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                    {
                        MessageBox.Show("!!Modificación cancelada!!", "OPERACION EXITOSA", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        Reiniciar();
                        return;
                    }
                    lgCompra.RegistroCompra(nuevaCompra, "ModificarCompra");
                    MessageBox.Show("Equipo modificado correctamente", "OPERACION EXITOSA", MessageBoxButtons.OK, MessageBoxIcon.Information);

                }
                else
                {
                    FolioRepetido = NuevaAccion.VerificarDuplicados(nuevaCompra.Factura, "VerificarFolioCompra");
                    if (FolioRepetido)
                    {
                        MessageBox.Show("Ingrese un numero de serie distinto", "SERIE YA EXISTENTE", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    else
                    {
                        lgCompra.RegistroCompra(nuevaCompra, "AgregarCompra");
                        MessageBox.Show("Equipo agregado correctamente", "OPERACION EXITOSA", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                Mostrar("MostrarCompras");
                Reiniciar();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ocurrio un error" + ex);
            }
        }

        

        private void btnEliminar_Click(object sender, EventArgs e)
        {
            //Preguntamos si se esta seguro de eliminar el registro 
            if (MessageBox.Show("¿Esta seguro de eliminar el registro?", "CONFIRME LA ELIMINACION", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                MessageBox.Show("Elimacion cancelada!!", "CANCELADO", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                Reiniciar();
                return;
            }
            try
            {
                string NumeroFolio = txtFolio.Text;
                lgCompra.EliminarCompra(NumeroFolio, "EliminarCompra");
                MessageBox.Show("Se elimino el registro", "OPERACION EXITOSA", MessageBoxButtons.OK, MessageBoxIcon.Information);
                Mostrar("MostrarCompras");
                Reiniciar(); 
            }
            catch (Exception ex)
            {
                MessageBox.Show("No se pudo eliminar el registro: " + ex, "OCURRIO UN PROBLEMA", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void btnCancelar_Click(object sender, EventArgs e)
        {
            Reiniciar();
        }
        #endregion

        #region Eventos
        private void dtgCompras_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            Modificando = true;
            ControlesDesactivados(true);

            txtFolio.Text = dtgCompras.CurrentRow.Cells[0].Value.ToString();
            txtCantidad.Text = dtgCompras.CurrentRow.Cells[1].Value.ToString();
            dtpFecha.Value = DateTime.Parse(dtgCompras.CurrentRow.Cells[2].Value.ToString());
            cboProveedores.SelectedItem = dtgCompras.CurrentRow.Cells[3].Value.ToString();

        }
        #endregion

        #region MetodosLocales
        private void LimpiarForm()
        {
            foreach (Control c in grpDatos.Controls)
            {
                if (c is TextBox)
                {
                    c.Text = "";
                }
            }
            cboProveedores.SelectedIndex = 0;
            dtpFecha.Value = DateTime.Today;
        }

        private void Reiniciar()
        {
            LimpiarForm();

            ControlesDesactivados(false);
            Modificando = false;
        }

        private void VerificarProveedor(Compra nuevaCompra)
        {
            if (chkNuevo.Checked)
            {
                //Agregamos el proveedor
                lgCompra.AgregarProveedor(txtProveedor.Text);
                //Lo buscamos
                nuevaCompra.IdProveedor = NuevaAccion.BuscarId(txtProveedor.Text, "ObtenerIdProveedor");
            }
            else
            {
                nuevaCompra.IdProveedor = NuevaAccion.BuscarId(cboProveedores.SelectedItem.ToString(), "ObtenerIdProveedor");
            }
        }
        #endregion

        private void chkNuevo_CheckedChanged(object sender, EventArgs e)
        {
            txtProveedor.Visible = chkNuevo.Checked;
            cboProveedores.Visible = !chkNuevo.Checked;
        }

    }
}
