using CobranzaSP.Lógica;
using CobranzaSP.Modelos;
using DocumentFormat.OpenXml.Office2010.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;


namespace CobranzaSP.Formularios
{
    public partial class InventarioPartesRicoh : Form
    {
        public InventarioPartesRicoh()
        {
            InitializeComponent();
            InicioAplicacion();
        }

        CD_Conexion cn = new CD_Conexion();
        LogicaInventarioPartesRicoh lgInventarioPartes = new LogicaInventarioPartesRicoh();
        LogicaRegistroInventaropPartesRicoh lgRegistroPartes = new LogicaRegistroInventaropPartesRicoh();
        //Objeto que nos servirá para poder tener capturado en dado caso de una modificacion
        RegistroInventario RegistroAnterior;
        AccionesLógica NuevaAccion = new AccionesLógica();

        bool Inventario = true;
        bool EstaModificando = false;
        int Id = 0;
        bool BuscandoRegistro = false;

        #region Inicio

        public void InicioAplicacion()
        {
            PropiedadesDtg();
            Mostrar("MostrarInventarioPartesRicoh");
            LlenarComboBox(cboModelo, "SeleccionarDescripcionesInventarioPartes", 0);
            ControlesDesactivados(false);
            radNuevo.Checked = true;
            radSalida.Checked = true;
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
            cb.Items.Insert(0, " ");
            cb.SelectedIndex = 0;
            dr.Close();
            cn.CerrarConexion();
        }

        public void ControlesDesactivados(bool Desactivado)
        {
            btnEliminar.Enabled = Desactivado;
            btnCancelar.Enabled = Desactivado;
        }

        public void PropiedadesDtg()
        {
            //Solo lectura
            dtgInventario.ReadOnly = true;

            //No agregar renglones
            dtgInventario.AllowUserToAddRows = false;

            //No borrar renglones
            dtgInventario.AllowUserToDeleteRows = false;

            //Ajustar automaticamente el ancho de las columnas
            dtgInventario.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            //dtgServicios.AutoResizeColumns(DataGridViewAutoSizeColumnsMo‌​de.Fill);
            dtgInventario.AutoResizeColumns();

            dtgInventario.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
        }
        public void Mostrar(string sp)
        {
            //Limpiamos los datos del datagridview
            dtgInventario.DataSource = null;
            dtgInventario.Refresh();
            DataTable tabla = new DataTable();
            //Guardamos los registros dependiendo la consulta
            tabla = NuevaAccion.Mostrar(sp);
            //Asignamos los registros que optuvimos al datagridview
            dtgInventario.DataSource = tabla;
            if (Inventario)
            {
                dtgInventario.Columns["IdNumeroParte"].Visible = false;
            }
        }

        #endregion

        #region Validaciones
        #endregion

        #region Botones
        private void btnGuardar_Click(object sender, EventArgs e)
        {
            
            try
            {
                if (Inventario)
                {
                    CapturaAInventario();
                }
                else
                {
                    CapturaMovimientoDeInventario();
                }
                

                LimpiarForm(grpDatosInventario);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public void CapturaMovimientoDeInventario()
        {
            MovimientoParteRicoh nuevoMovimiento = new MovimientoParteRicoh()
            {
                IdRegistro = Id,
                IdNumeroParte = NuevaAccion.BuscarId(cboModelo.SelectedItem.ToString(), "ObtenerIdModeloPartes"),
                Cantidad = int.Parse(txtCantidad.Text),
                TipoMovimiento = (radSalida.Checked) ? "Salida" : "Entrada",
                Fecha = dtpFechaRegistro.Value
            };
            ColocarIdTipoPersona(nuevoMovimiento);

            if (EstaModificando)
            {

            }
            else
            {
                lgRegistroPartes.AgregarRegistroInventario(nuevoMovimiento);
                MessageBox.Show("Parte agregada al inventario", "REGISTRO EXITOSO", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            LimpiarForm(grpDatosRegistro);
            Mostrar("MostrarRegistroInventarioPartesRicoh");

            //MessageBox.Show("IdNumeroParte" + nuevoMovimiento.IdNumeroParte + " Cantidad:" +nuevoMovimiento.Cantidad + " TipoMovimiento: " 
            //    +nuevoMovimiento.TipoMovimiento + " IdPersona "+ nuevoMovimiento.IdTipoPersona);
        }

        public void ColocarIdTipoPersona(MovimientoParteRicoh nuevoMovimiento)
        {
            if (radSalida.Checked)
            {
                nuevoMovimiento.IdTipoPersona = NuevaAccion.BuscarId(cboTipoPersona.SelectedItem.ToString(), "ObtenerIdCliente");
            }
            else
            {
                nuevoMovimiento.IdTipoPersona = NuevaAccion.BuscarId(cboTipoPersona.SelectedItem.ToString(), "ObtenerIdProveedor");
                if (chkCliente.Checked)
                {
                    nuevoMovimiento.IdTipoPersona = NuevaAccion.BuscarId(cboTipoPersona.SelectedItem.ToString(), "ObtenerIdCliente");
                }
            }
        }

        public void CapturaAInventario()
        {
            //if (!ValidarCamposInventario())
            //    return;
            ParteRicoh nuevoParte = new ParteRicoh()
            {
                IdNumeroParte = Id,
                NumeroParte = txtNumeroParte.Text,
                Cantidad = 0,
                Descripcion = rtxtDescripcion.Text
            };
            if (radUsado.Checked && !EstaModificando)
            {
                nuevoParte.Descripcion += " Usada";
                nuevoParte.NumeroParte += "U";
            }


            if (EstaModificando)
            {
                //Modificar stop procedure de modificar, para que se pueda cambiar la fecha
                if (MessageBox.Show("¿Esta seguro de modificar el registro?", "CONFIRME LA MODIFICACION", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                {
                    MessageBox.Show("!!Modificación cancelada!!", "CANCELADO", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                    LimpiarForm(grpDatosInventario);
                    return;
                }
                lgInventarioPartes.RegistroParteRicoh(nuevoParte, "ModificarParteRicoh");
                MessageBox.Show("Parte modificada correctamente", "MODIFICACION DE REGISTRO", MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
            else
            {
                bool ModeloDuplicado = NuevaAccion.VerificarDuplicados(nuevoParte.NumeroParte, "VerificarNumeroParteRicoh");
                if (ModeloDuplicado)
                {
                    MessageBox.Show("Ingrese un modelo distinto", "EL MODELO YA EXISTE", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                else
                {
                    lgInventarioPartes.RegistroParteRicoh(nuevoParte, "AgregarParteRicoh");
                    MessageBox.Show("Parte agregada al inventario", "REGISTRO EXITOSO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                //lgInventarioPartes.RegistroParteRicoh(nuevoParte, "AgregarParteRicoh");
                //MessageBox.Show("Numero de parte agregada correctamente al inventario", "REGISTRO EXITOSO", MessageBoxButtons.OK, MessageBoxIcon.Information);

            }

            //ESTA LA USO EN MI BASE DE DATOS
            Mostrar("MostrarInventarioPartesRicoh");
        }

        private void btnEliminar_Click(object sender, EventArgs e)
        {

        }

        private void btnCancelar_Click(object sender, EventArgs e)
        {
            LimpiarForm(grpDatosInventario);
            ControlesDesactivados(false);
            EstaModificando = false;
            Id = 0;
        }
        #endregion

        #region Eventos
        private void dtgInventario_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            ControlesDesactivados(true);
            EstaModificando = true;

            Id = int.Parse(dtgInventario.CurrentRow.Cells[0].Value.ToString());
            txtNumeroParte.Text = dtgInventario.CurrentRow.Cells[1].Value.ToString();
            rtxtDescripcion.Text = dtgInventario.CurrentRow.Cells[2].Value.ToString();
        }
        #endregion

        #region MetodosLocales
        private void LimpiarForm(GroupBox grp)
        {
            foreach (Control c in grp.Controls)
            {
                if (c is TextBox || c is RichTextBox)
                {
                    c.Text = "";
                }
            }
            cboModelo.SelectedIndex = 0;
            cboTipoPersona.SelectedIndex = 0;
        }
        #endregion

        private void btnImprimir_Click(object sender, EventArgs e)
        {
            lgInventarioPartes.ImprimirInventario();
        }

        private void grpDatosRegistro_Enter(object sender, EventArgs e)
        {

        }

        private void radSalida_CheckedChanged(object sender, EventArgs e)
        {
            if(radSalida.Checked)
            {
                LlenarClientes();
                chkCliente.Visible = false;
                chkCliente.Checked = false;
            }
            else
            {
                LlenarProveedores();
                chkCliente.Visible = true;
            }
        }

        public void LlenarProveedores()
        {
            lblTipoPersona.Text = "Proveedor:";
            LlenarComboBox(cboTipoPersona, "SeleccionarProveedores", 0);
        }

        public void LlenarClientes()
        {
            lblTipoPersona.Text = "Cliente:";
            LlenarComboBox(cboTipoPersona, "SeleccionarClientesServicios", 0);
        }

        private void chkCliente_CheckedChanged(object sender, EventArgs e)
        {
            if(chkCliente.Checked) 
            {
                LlenarClientes();
            }
            else
            {
                LlenarProveedores();
            }
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            //ReiniciarControles();
            if (tabControl1.SelectedTab == tabInventario)
            {
                Inventario = true;
                Mostrar("MostrarInventarioPartesRicoh");

            }
            else
            {
                Inventario = false;
                Mostrar("MostrarRegistroInventarioPartesRicoh");
            }
        }

        private void AbrirForm(object formNuevo)
        {
            //Declaramos la forma
            Form fh = formNuevo as Form;

            //Mostramos la forma 
            fh.Show();
        }

        private void btnGenerarReporte_Click(object sender, EventArgs e)
        {
            AbrirForm(new ReporteRegistrosInventarioPartesRicoh());
        }
    }
}
