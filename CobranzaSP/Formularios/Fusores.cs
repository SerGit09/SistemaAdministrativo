using CobranzaSP.Lógica;
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
using CobranzaSP.Modelos;
using DocumentFormat.OpenXml.Office2010.Excel;

namespace CobranzaSP.Formularios
{
    public partial class Fusores : Form
    {
        public Fusores()
        {
            InitializeComponent();
            InicioAplicacion();
        }
        AccionesLógica NuevaAccion = new AccionesLógica();
        CD_Conexion cn = new CD_Conexion();
        LogicaFusor AccionFusor = new LogicaFusor();
        bool Modificando = false;
        int Id;

        #region Inicio

        public void InicioAplicacion()
        {
            string[] opcionesBusqueda = { "","Habilitado", "Deshabilitado", "Rango Fecha", "Serie", "Todos" };
            cboBusqueda.Items.AddRange(opcionesBusqueda);

            btnGenerarReporte.Enabled = false;

            string[] Proveedores = { "", "American Stock", "GPARTS (Ricardo Pascoe)", "MultiSid", "Copy Copias", "Mercado Libre"};
            cboProveedores.Items.AddRange(Proveedores);

            string[] NumeroDias = {"", "0", "30", "90", "120", "180" };
            cboDiasGarantía.Items.AddRange(NumeroDias);

            PropiedadesDtgServicios();
            ControlesDesactivadosInicialmente(false);
            //Llenar el combobox de los modelos
            LlenarComboBox(cboModelos, "SeleccionarModelosCartucho");
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

        public void PropiedadesDtgServicios()
        {
            //Solo lectura
            dtgFusores.ReadOnly = true;

            //No agregar renglones
            dtgFusores.AllowUserToAddRows = false;

            //No borrar renglones
            dtgFusores.AllowUserToDeleteRows = false;

            //Ajustar automaticamente el ancho de las columnas
            dtgFusores.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            //dtgServicios.AutoResizeColumns(DataGridViewAutoSizeColumnsMo‌​de.Fill);
            dtgFusores.AutoResizeColumns();
            MostrarDatosFusores();
        }

        private void ControlesDesactivadosInicialmente(bool activado)
        {
            btnCancelar.Enabled = activado;
            btnEliminar.Enabled = activado;
            btnGenerarReporte.Enabled = activado;
        }

        public void MostrarDatosFusores()
        {
            //Limpiamos los datos del datagridview
            dtgFusores.DataSource = null;
            dtgFusores.Refresh();
            DataTable tabla = new DataTable();
            //Guardamos los registros dependiendo la consulta
            tabla = NuevaAccion.Mostrar("MostrarTodosFusores");
            //Asignamos los registros que optuvimos al datagridview
            dtgFusores.DataSource = tabla;
        }

        #endregion

        #region Validaciones
        public bool ValidarCampos()
        {
            bool Validado = true;
            erFusores.Clear();
            if (string.IsNullOrEmpty(txtSerie.Text))
            {
                erFusores.SetError(txtSerie, "Campo obligatorio");
                Validado = false;
            }
            if (string.IsNullOrEmpty(txtSerieSp.Text))
            {
                erFusores.SetError(txtSerieSp, "Campo obligatorio");
                Validado = false;
            }
            if (string.IsNullOrEmpty(txtFactura.Text))
            {
                erFusores.SetError(txtFactura, "Campo obligatorio");
                Validado = false;
            }
            if (string.IsNullOrEmpty(txtCosto.Text))
            {
                erFusores.SetError(txtCosto, "Campo obligatorio");
                Validado = false;
            }
            if (cboDiasGarantía.SelectedItem == null)
            {
                erFusores.SetError(cboDiasGarantía, "Campo obligatorio");
                Validado = false;
            }
            return Validado;
        }

        public bool ValidarCamposReporte()
        {
            bool Validado = true;
            erFusores.Clear();
            if (cboBusqueda.SelectedItem == null)
            {
                erFusores.SetError(cboBusqueda, "Campo obligatorio");
                Validado = false;
            }
            else
            {
                string Parametro = cboBusqueda.SelectedItem.ToString();
                switch (Parametro)
                {
                    case "Serie":
                        if (string.IsNullOrEmpty(txtSerieBusqueda.Text))
                        {
                            erFusores.SetError(txtSerieBusqueda, "Campo obligatorio");
                            Validado = false;
                        }; break;
                }
            }
            return Validado;

        }
        #endregion

        #region Botones
        private void btnGuardar_Click(object sender, EventArgs e)
        {
            bool Repetido = false;
            try
            {
                if (!ValidarCampos())
                {
                    return;
                }

                string Estado = (radActivo.Checked) ? "Activado" : "Desactivado";
                Fusor nuevoFusor = new Fusor()
                {
                    IdFusor = Id,
                    SerieO = txtSerie.Text,
                    SerieS = txtSerieSp.Text,
                    NumeroFactura = txtFactura.Text,
                    Proveedor = cboProveedores.SelectedItem.ToString(),
                    FechaFactura = dtpFechaFactura.Value,
                    Cantidad = double.Parse(txtCosto.Text.Replace(",", "")),
                    DiasGarantia = int.Parse(cboDiasGarantía.SelectedItem.ToString()),
                    Estado = Estado,
                    Modelo = cboModelos.SelectedItem.ToString()
                };

                if (Modificando)
                {
                    if (MessageBox.Show("Desea modificar el registro?", "CONFIRME LA MODIFICACION", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                    {
                        MessageBox.Show("Modificacion cancelada!!", "CANCELADO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        LimpiarForm();
                        return;
                    }
                    AccionFusor.GuardarFusor(nuevoFusor, "ModificarFusor");
                    MessageBox.Show("Fusor modificado correctamente", "OPERACION EXITOSA", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    MostrarDatosFusores();
                }
                else
                {
                    //Verificamos que los numeros de serie no esten duplicados
                    Repetido = NuevaAccion.VerificarDuplicados(nuevoFusor.SerieO, "VerificarSerieExistenteFusor");
                    Repetido = NuevaAccion.VerificarDuplicados(nuevoFusor.SerieS, "VerificarSerieExistenteFusor");
                    if (Repetido)
                    {
                        MessageBox.Show("Ingrese un numero de serie distinto", "SERIE YA EXISTENTE", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    else
                    {
                        AccionFusor.GuardarFusor(nuevoFusor, "AgregarFusor");
                        MessageBox.Show("Fusor agregado correctamente", "OPERACION EXITOSA", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        MostrarDatosFusores();
                    }
                }
                ReiniciarControles();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void btnGenerarExcel_Click(object sender, EventArgs e)
        {
            AccionFusor.GenerarExcel();
        }

        private void btnGenerarReporte_Click(object sender, EventArgs e)
        {
            try
            {
                if (!ValidarCamposReporte())
                    return;
                string Parametro = cboBusqueda.SelectedItem.ToString();
                string Serie = txtSerieBusqueda.Text;
                if (Serie != "")
                {
                    SqlDataReader dr = NuevaAccion.Buscar(txtSerieBusqueda.Text, "BuscarSerieSp");
                    //Si no nos regresa un registro quiere decir que no existe en nuestra base de datos
                    if (!dr.Read())
                    {
                        MessageBox.Show("Serie no encontrada", "SERIE NO EXISTENTE", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                    dr.Close();
                }
                DateTime FechaInicio = dtpFechaInicio.Value;
                DateTime FechaFinal = dtpFechaFinal.Value;
                AccionFusor.ReporteFusores(Parametro, FechaInicio, FechaFinal, Serie);
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
            txtSerie.Focus();
        }

        private void btnEliminar_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Desea eliminar el registro?", "CONFIRME LA ELIMINACION", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                MessageBox.Show("Eliminación cancelada!!", "OPERACION CANCELADA", MessageBoxButtons.OK, MessageBoxIcon.Information);
                LimpiarForm();
                return;
            }

            NuevaAccion.Eliminar(Id, "EliminarFusor");
            MessageBox.Show("Fusor eliminado correctamente", "OPERACION EXITOSA", MessageBoxButtons.OK, MessageBoxIcon.Information);
            MostrarDatosFusores();
            ReiniciarControles();
        }
        #endregion

        #region Eventos
        private void cboBusqueda_SelectedIndexChanged(object sender, EventArgs e)
        {
            btnGenerarReporte.Enabled = true;
            MostrarFechas(false);
            string Parametro = cboBusqueda.SelectedItem.ToString();
            switch (Parametro)
            {
                case "Habilitado": txtSerieBusqueda.Visible = false; break;
                case "Deshabilitada": txtSerieBusqueda.Visible = false; break;
                case "Rango Fecha": MostrarFechas(true); txtSerieBusqueda.Visible = false; break;
                case "Serie": txtSerieBusqueda.Visible = true; txtSerieBusqueda.Focus(); break;
                case "Todos": txtSerieBusqueda.Visible = false; break;
            }
        }

        private void dtgFusores_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            ControlesDesactivadosInicialmente(true);
            LimpiarForm();
            Modificando = true;
            bool activo;
            Id = int.Parse(dtgFusores.CurrentRow.Cells[0].Value.ToString());
            txtSerie.Text = dtgFusores.CurrentRow.Cells[1].Value.ToString();
            txtSerieSp.Text = dtgFusores.CurrentRow.Cells[2].Value.ToString();
            txtFactura.Text = dtgFusores.CurrentRow.Cells[3].Value.ToString();
            cboProveedores.SelectedItem = dtgFusores.CurrentRow.Cells[4].Value.ToString();
            dtpFechaFactura.Value = Convert.ToDateTime(dtgFusores.CurrentRow.Cells[5].Value.ToString());
            txtCosto.Text = dtgFusores.CurrentRow.Cells[7].Value.ToString().Replace("$", "");

            string Estado = dtgFusores.CurrentRow.Cells[10].Value.ToString();
            cboModelos.SelectedItem = dtgFusores.CurrentRow.Cells[11].Value.ToString();
            activo = (Estado == "Activado") ? true : false;

            cboDiasGarantía.SelectedItem = dtgFusores.CurrentRow.Cells[8].Value.ToString();


            if (activo)
            {
                radActivo.Checked = activo;
            }
            else
            {
                radInactivo.Checked = true;
            }
        }
        #endregion

        #region MetodosLocales
        public void MostrarFechas(bool Mostrar)
        {
            lblFechaFinal.Visible = Mostrar;
            lblFechaInicio.Visible = Mostrar;
            dtpFechaInicio.Visible = Mostrar;
            dtpFechaFinal.Visible = Mostrar;
        }

        public void LimpiarForm()
        {
            foreach (Control c in grpDatos.Controls)
            {
                if (c is TextBox)
                {
                    c.Text = "";
                }
            }
            dtpFechaFactura.Value = DateTime.Today;
            dtpFechaInicio.Value = DateTime.Today;
            dtpFechaFinal.Value = DateTime.Today;
            MostrarFechas(false);
            txtSerieBusqueda.Visible = false;
            cboBusqueda.Text = "";
            cboProveedores.SelectedIndex = 0;
            cboModelos.SelectedIndex = 0;
            txtSerie.Focus();
        }

        public void ReiniciarControles()
        {
            LimpiarForm();
            ControlesDesactivadosInicialmente(false);
            Modificando = false;
        }
        #endregion

        private void cboModelos_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
    }
}
