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
    public partial class Equipos : Form
    {
        public Equipos()
        {
            InitializeComponent();
            InicioAplicacion();
        }
        AccionesLógica NuevaAccion = new AccionesLógica();
        CD_Conexion cn = new CD_Conexion();
        LogicaEquipos lgEquipo = new LogicaEquipos();
        LogicaEquiposBodega lgEquipoBodega = new LogicaEquiposBodega();
        //Variable que servira para los  reportes
        string TipoBusqueda = "";
        string Parametro = "";

        //Sabremos cuando estamos añadiendo un nuevo registro o modificando
        bool Modificando = false;
        bool BuscandoFolio = false;
        private bool Buscando = false;
        int Id;

        #region Inicio
        public void InicioAplicacion()
        {
            PropiedadesDtg();

            ControlesDesactivados(false);

            //Se manda el nombre del cbo,el stop procedure que ejecutara, en dado caso de que sea modelos se manda el id de la marca
            LlenarComboBox(cboClientes, "SeleccionarClientesServicios", 0);
            LlenarComboBox(cboTipoRenta, "SeleccionarTipoRenta", 0);
            LlenarComboBox(cboMarcas, "SeleccionarMarca", 0);
            LlenarComboBox(cboModelos, "SeleccionarModelos", 0);

            //Llenamos el datagridview
            Mostrar("MostrarEquipos");

            //Deshabilitamos escritura en combobox
            cboClientes.DropDownStyle = ComboBoxStyle.DropDownList;
            cboTipoRenta.DropDownStyle = ComboBoxStyle.DropDownList;

        }


        public void ControlesDesactivados(bool Desactivado)
        {
            btnEliminar.Enabled = Desactivado;
            btnCancelar.Enabled = Desactivado;
            btnEnviarABodega.Enabled = Desactivado;
            //btnMostrar.Enabled = Desactivado;
        }

        public void PropiedadesDtg()
        {
            //Solo lectura
            dtgEquipos.ReadOnly = true;

            //No agregar renglones
            dtgEquipos.AllowUserToAddRows = false;

            //No borrar renglones
            dtgEquipos.AllowUserToDeleteRows = false;

            //Ajustar automaticamente el ancho de las columnas
            dtgEquipos.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            //dtgServicios.AutoResizeColumns(DataGridViewAutoSizeColumnsMo‌​de.Fill);
            dtgEquipos.AutoResizeColumns();

            dtgEquipos.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
        }

        public void LlenarComboBox(ComboBox cb, string sp, int Marca)
        {
            SqlDataReader dr;
            cb.Items.Clear();
            if (sp == "SeleccionarModelos")
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

        public void Mostrar(string sp)
        {
            //Limpiamos los datos del datagridview
            dtgEquipos.DataSource = null;
            dtgEquipos.Refresh();
            DataTable tabla = new DataTable();
            //Guardamos los registros dependiendo la consulta
            tabla = NuevaAccion.Mostrar(sp);
            //Asignamos los registros que optuvimos al datagridview
            dtgEquipos.DataSource = tabla;
        }
        #endregion


        #region Validaciones
        public bool ValidarCampos()
        {
            bool Validado = true;
            erEquipos.Clear();

            foreach (Control c in grpDatos.Controls)
            {
                if ((c is ComboBox || c is TextBox) && c.Name != "txtReferencia")
                {
                    if (c is ComboBox && ComboBoxEstaVacio((ComboBox)c))
                    {
                        erEquipos.SetError(c, "Campo Obligatorio");
                        Validado = false;
                    }
                    else if (c is TextBox && string.IsNullOrEmpty(c.Text))
                    {
                        erEquipos.SetError(c, "Campo Obligatorio");
                        Validado = false;
                    }
                }

            }
            return Validado;
        }

        public bool ComboBoxEstaVacio(ComboBox c)
        {
            if (c.SelectedItem == null || c.SelectedIndex == 0)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        

        private void txtSerie_KeyPress(object sender, KeyPressEventArgs e)
        {
            Validacion.SoloLetrasYNumeros(e);
        }

        private void txtReferencia_KeyPress(object sender, KeyPressEventArgs e)
        {
            Validacion.SoloLetrasYNumeros(e);
        }

        private void txtPrecio_KeyPress(object sender, KeyPressEventArgs e)
        {
            Validacion.SoloNumeros(e);
            if ((e.KeyChar == '.') && (!txtPrecio.Text.Contains(".")))
            {
                e.Handled = false;
            }
        }

        private void txtFechaPago_KeyPress(object sender, KeyPressEventArgs e)
        {
            Validacion.SoloLetrasYNumeros(e);
        }
        #endregion

        #region Botones
        private void btnGuardar_Click(object sender, EventArgs e)
        {
            bool SerieRepetida = false;
            try
            {
                if (!ValidarCampos())
                    return;

                Equipo nuevoEquipo = new Equipo()
                {
                    IdEquipo = Id,
                    IdCliente = NuevaAccion.BuscarId(cboClientes.SelectedItem.ToString(), "ObtenerIdCliente"),
                    Ubicacion = txtReferencia.Text,
                    IdMarca = NuevaAccion.BuscarId(cboMarcas.SelectedItem.ToString(), "ObtenerIdMarca"),
                    IdModelo = NuevaAccion.BuscarId(cboModelos.SelectedItem.ToString(), "ObtenerIdModelo"),
                    Serie = txtSerie.Text,
                    IdRenta = NuevaAccion.BuscarId(cboTipoRenta.SelectedItem.ToString(), "ObtenerIdTipoRenta"),
                    Precio = double.Parse(txtPrecio.Text.Replace(",", "")),
                    FechaPago = txtFechaPago.Text,
                    Valor = double.Parse(txtPrecioEquipo.Text)
                };

                if (Modificando)
                {
                    //Modificar stop procedure de modificar, para que se pueda cambiar la fecha
                    if (MessageBox.Show("¿Esta seguro de modificar el registro?", "CONFIRME LA MODIFICACION", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                    {
                        MessageBox.Show("!!Modificación cancelada!!", "OPERACION EXITOSA", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        LimpiarForm();
                        return;
                    }
                    lgEquipo.GuardarRegistro(nuevoEquipo, "ModificarEquipo");
                    MessageBox.Show("Equipo modificado correctamente", "OPERACION EXITOSA", MessageBoxButtons.OK, MessageBoxIcon.Information);

                }
                else
                {
                    SerieRepetida = NuevaAccion.VerificarDuplicados(nuevoEquipo.Serie, "VerificarDuplicadoSerieEquipos");
                    if (SerieRepetida)
                    {
                        MessageBox.Show("Ingrese un numero de serie distinto", "SERIE YA EXISTENTE", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    else
                    {
                        lgEquipo.GuardarRegistro(nuevoEquipo, "AgregarEquipo");
                        MessageBox.Show("Equipo agregado correctamente", "OPERACION EXITOSA", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                Mostrar("MostrarEquipos");
                //LimpiarForm();
                ReiniciarControles();


            }
            catch (Exception ex)
            {
                MessageBox.Show("Ocurrio un error" + ex);
            }
        }

        private void btnEliminar_Click(object sender, EventArgs e)
        {
            try
            {
                if (MessageBox.Show("¿Esta seguro de eliminar el registro?", "CONFIRME LA ELIMINACION", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                {
                    MessageBox.Show("!!Eliminacion cancelada!!", "CANCELADO", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                    LimpiarForm();
                    return;
                }
                NuevaAccion.Eliminar(Id, "EliminarEquipo");
                MessageBox.Show("Se ha eliminado el registro correctamente", "ELIMINACION CONFIRMADA", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                Mostrar("MostrarEquipos");
                LimpiarForm();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void btnCancelar_Click(object sender, EventArgs e)
        {
            ReiniciarControles();
        }


        
        #endregion

        #region Eventos
        private void cboMarcas_SelectedIndexChanged(object sender, EventArgs e)
        {
            //En dado caso que se haya seleccionado algo de las marcas
            if (cboMarcas.SelectedItem.ToString() != " " && Buscando == false)
            {
                int IdMarca = NuevaAccion.BuscarId(cboMarcas.SelectedItem.ToString(), "ObtenerIdMarca");
                //Se llenara de acuerdo a la marca que se haya escogido
                LlenarComboBox(cboModelos, "SeleccionarModelos", IdMarca);
            }
        }
        private void dtgEquipos_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            LimpiarForm();
            ControlesDesactivados(true);
            Modificando = true;
            erEquipos.Clear();

            Id = int.Parse(dtgEquipos.CurrentRow.Cells[0].Value.ToString());
            cboClientes.SelectedItem = dtgEquipos.CurrentRow.Cells[1].Value.ToString();
            txtReferencia.Text = dtgEquipos.CurrentRow.Cells[2].Value.ToString();
            cboMarcas.SelectedItem = dtgEquipos.CurrentRow.Cells[3].Value.ToString();
            cboModelos.SelectedItem = dtgEquipos.CurrentRow.Cells[4].Value.ToString();
            txtSerie.Text = dtgEquipos.CurrentRow.Cells[5].Value.ToString();
            cboTipoRenta.SelectedItem = dtgEquipos.CurrentRow.Cells[6].Value.ToString();
            txtPrecio.Text = dtgEquipos.CurrentRow.Cells[7].Value.ToString().Replace("$", "");
            txtFechaPago.Text = dtgEquipos.CurrentRow.Cells[8].Value.ToString();
            txtPrecioEquipo.Text = dtgEquipos.CurrentRow.Cells[9].Value.ToString().Replace("$", "");
        }

        #endregion


        #region MetodosLocales
        public void LimpiarForm()
        {
            foreach (Control c in grpDatos.Controls)
            {
                if (c is TextBox)
                {
                    c.Text = "";
                }
            }
            cboClientes.Focus();

            cboClientes.SelectedIndex = 0;
            cboTipoRenta.SelectedIndex = 0;
            cboModelos.SelectedIndex = 0;
            LlenarComboBox(cboModelos, "SeleccionarModelos", 0);
        }

        public void ReiniciarControles()
        {
            ControlesDesactivados(false);
            LimpiarForm();
            Modificando = false;
        }

        #endregion

        private void grpDatos_Enter(object sender, EventArgs e)
        {

        }



        private void btnEnviarABodega_Click(object sender, EventArgs e)
        {
            string Mensaje;
            bool SerieRepetida;

            if (MessageBox.Show("¿Esta seguro de enviar el equipo a bodega?", "CONFIRMAR ACCION", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                MessageBox.Show("!!Envio cancelado!!", "CANCELADO", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                LimpiarForm();
                return;
            }

            try
            {
                if (!ValidarCampos())
                    return;

                EquipoBodega NuevoEquipo = new EquipoBodega()
                {
                    IdMarca = NuevaAccion.BuscarId(cboMarcas.SelectedItem.ToString(), "ObtenerIdMarca"),
                    IdModelo = NuevaAccion.BuscarId(cboModelos.SelectedItem.ToString(), "ObtenerIdModelo"),
                    Serie = txtSerie.Text,
                    Precio = double.Parse(txtPrecioEquipo.Text),
                    Estado = "Usada",
                    Notas = ""
                };
                SerieRepetida = NuevaAccion.VerificarDuplicados(NuevoEquipo.Serie, "VerificarSerieEquiposBodega");
                if (SerieRepetida)
                {
                    MessageBox.Show("Ingrese un numero de serie distinto", "SERIE YA EXISTENTE", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    LimpiarForm();
                    return;
                }
                Mensaje = lgEquipoBodega.GuardarEquipo(NuevoEquipo, "AgregarEquipoABodega");

                //Deberemos de eliminar el equipo que se envio a bodega
                NuevaAccion.Eliminar(Id, "EliminarEquipo");

                LimpiarForm();
                Mostrar("MostrarEquipos");
            }
            catch (Exception ex)
            {
                MessageBox.Show("No se pudieron guardar los datos por: " + ex);
            }
        }


        private void btnReportes_Click(object sender, EventArgs e)
        {
            ReporteEquipoRenta NuevoReporte = new ReporteEquipoRenta();
            NuevoReporte.Show();
        }
    }
}
