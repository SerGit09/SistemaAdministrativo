using CobranzaSP.Lógica;
using System;
using System.Data.SqlClient;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using CobranzaSP.Modelos;
using System.Windows.Forms.DataVisualization.Charting;
using DocumentFormat.OpenXml.Office2010.Excel;
using DocumentFormat.OpenXml.Vml.Spreadsheet;
using DocumentFormat.OpenXml.Bibliography;
using System.Reflection;

namespace CobranzaSP.Formularios
{
    public partial class Modulos : Form
    {
        public Modulos()
        {
            InitializeComponent();
            InicioAplicacion();
        }
        AccionesLógica nuevaAccion = new AccionesLógica();
        CD_Conexion cn = new CD_Conexion();
        LogicaRegistroModulos lgModulo = new LogicaRegistroModulos();
        string NumeroFolio;
        bool Validado = true;
        int IdRegistroInsumo = 0;
        int IdRegistroParte = 0;

        //Sabremos cuando estamos añadiendo un nuevo registro o modificando
        bool Modificar = false; 
        bool ModificarParte = false;
        bool BuscandoFolio = false;

        #region Inicio
        public void InicioAplicacion()
        {
            PropiedadesDtgServicios(dtgFolioos);
            PropiedadesDtgServicios(dtgPartes);

            string[] Tecnicos = { "", "ALVARO", "MIGUEL", };
            cboPersonaRecibido.Items.AddRange(Tecnicos);


            LlenarComboBox(cboModelo, "SeleccionarNumeroPartes", 0);

            MostrarDatosHoja();
            txtFolioSeleccionado.Enabled = false;

            HabilitarBotonesInsumos(false);
            HabilitarBotonesPartes(false);
        }

        public void HabilitarBotonesInsumos(bool habilitado)
        {
            btnEliminar.Enabled = habilitado;
            btnCancelar.Enabled = habilitado;
        }

        public void HabilitarBotonesPartes(bool habilitado)
        {
            btnEliminarParte.Enabled = habilitado;
            btnCancelarParte.Enabled = habilitado;
        }

        public void LlenarComboBox(ComboBox cb, string sp, int Marca)
        {
            SqlDataReader dr;
            cb.Items.Clear();
            if (Marca != 0 || sp == "SeleccionarModelos")
            {
                dr = nuevaAccion.LlenarComboBoxEspecifico(sp, Marca);
            }
            else
            {
                dr = nuevaAccion.LlenarComboBox(sp);
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

        public void PropiedadesDtgServicios(DataGridView dtg)
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

        public void MostrarDatosHoja()
        {
            //Limpiamos los datos del datagridview
            dtgFolioos.DataSource = null;
            dtgFolioos.Refresh();
            DataTable tabla = new DataTable();
            //Guardamos los registros dependiendo la consulta
            tabla = nuevaAccion.Mostrar("MostrarRegistrosInsumos");
            //Asignamos los registros que optuvimos al datagridview
            dtgFolioos.DataSource = tabla;
            dtgFolioos.Columns["IdRegistro"].Visible = false;
        }

        public void MostrarPartes(int Folio)
        {
            //Limpiamos los datos del datagridview
            dtgPartes.DataSource = null;
            dtgPartes.Refresh();
            DataTable tabla = new DataTable();
            //Guardamos los registros dependiendo la consulta
            tabla = lgModulo.MostrarPartes(Folio);
            //Asignamos los registros que optuvimos al datagridview
            dtgPartes.DataSource = tabla;
            dtgPartes.Columns["IdRegistro"].Visible = false;
        }
        #endregion

        #region Validaciones
        #endregion

        #region Botones
        private void btnGuardar_Click(object sender, EventArgs e)
        {
            try
            {
                //if (!ValidarCamposVacios())
                //    return;

                string Mensaje;
                Modulo nuevoModulo = new Modulo()
                {
                    Folio = int.Parse(txtFolio.Text),
                    IdRegistro = IdRegistroInsumo,
                    Fecha = dtpFecha.Value,
                    Recibido = cboPersonaRecibido.SelectedItem.ToString(),
                    Concepto = txtConcepto.Text
                };


                if (Modificar)
                {
                    if (MessageBox.Show("¿Esta seguro de modificar el registro?", "CONFIRME LA MODIFICACION", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                    {
                        MessageBox.Show("Modificacion cancelada!!", "CANCELADO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        LimpiarForm(grpDatosHoja);
                        return;
                    }
                    Mensaje = lgModulo.RegistroInsumo(nuevoModulo, "ModificarRegistroInsumo");
                    MessageBox.Show(Mensaje, "MODIFICANDO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    MostrarPartes(int.Parse(txtFolioSeleccionado.Text));
                }
                else
                {
                    bool FolioRepetido = lgModulo.VerificarExistenciaFolio(nuevoModulo.Folio, "VerificarFolioInsumo");
                    if (FolioRepetido)
                    {
                        MessageBox.Show("El numero de folio ya existe!!", "DUPLICADO", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        LimpiarForm(grpDatosHoja);
                        MostrarDatosHoja();
                        return;
                    }

                    Mensaje = lgModulo.RegistroInsumo(nuevoModulo, "AgregarRegistroInsumo");
                    MessageBox.Show(Mensaje, "REGISTRO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                LimpiarForm(grpDatosHoja);
                MostrarDatosHoja();
            }
            catch (Exception ex)
            {
                MessageBox.Show("No se pudieron guardar los datos por: " + ex);
            }
        }

        private void btnGuardarParte_Click(object sender, EventArgs e)
        {
            try
            {
                //if (!ValidarCamposVacios())
                //    return;

                string Mensaje;
                Parte nuevaParte = new Parte()
                {
                    Id = IdRegistroParte, 
                    Folio = int.Parse(txtFolioSeleccionado.Text),
                    Cantidad = int.Parse(txtCantidad.Text),
                    Descripcion = rtxtDescripcion.Text,
                    Estado = (radNuevo.Checked) ? "Nuevo" : "Usado",
                    IdNumeroParte = nuevaAccion.BuscarId(cboModelo.SelectedItem.ToString(), "ObtenerIdModeloPartes")
                };


                if (ModificarParte)
                {
                    if (MessageBox.Show("¿Esta seguro de modificar el registro?", "CONFIRME LA MODIFICACION", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                    {
                        MessageBox.Show("Modificacion cancelada!!", "CANCELADO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        LimpiarForm(grpDatosHoja);
                        return;
                    }
                    Mensaje = lgModulo.RegistroParte(nuevaParte, "ModificarRegistroParte");
                    MessageBox.Show(Mensaje, "MODIFICANDO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    Mensaje = lgModulo.RegistroParte(nuevaParte, "AgregarRegistroParte");
                    MessageBox.Show(Mensaje, "REGISTRO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                LimpiarForm(grpDatosPartes);
                MostrarPartes(int.Parse(txtFolioSeleccionado.Text));
            }
            catch (Exception ex)
            {
                MessageBox.Show("No se pudieron guardar los datos por: " + ex);
            }
        }


        private void btnCancelar_Click(object sender, EventArgs e)
        {
            HabilitarBotonesInsumos(false);
            LimpiarForm(grpDatosHoja);
            cboPersonaRecibido.SelectedIndex = 0;

            Reiniciar();
        }


        private void btnEliminarParte_Click(object sender, EventArgs e)
        {
            //Preguntamos si se esta seguro de eliminar el registro 
            if (MessageBox.Show("¿Esta seguro de eliminar el registro?", "CONFIRME LA ELIMINACION", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                MessageBox.Show("Elimacion cancelada!!", "CANCELADO", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                LimpiarForm(grpDatosHoja);
                return;
            }

            try
            {
                nuevaAccion.Eliminar(IdRegistroParte, "EliminarRegistroParte");
                MessageBox.Show("Se elimino el registro", "OPERACION EXITOSA", MessageBoxButtons.OK, MessageBoxIcon.Information);
                LimpiarForm(grpDatosPartes);
                MostrarPartes(int.Parse(txtFolioSeleccionado.Text));
            }
            catch (Exception ex)
            {
                MessageBox.Show("No se pudo eliminar el registro: " + ex, "OCURRIO UN PROBLEMA", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void btnCancelarParte_Click(object sender, EventArgs e)
        {
            HabilitarBotonesPartes(false);
            LimpiarForm(grpDatosPartes);
            ModificarParte = false;
        }

        private void btnEliminar_Click(object sender, EventArgs e)
        {
            //Preguntamos si se esta seguro de eliminar el registro 
            if (MessageBox.Show("¿Esta seguro de eliminar el registro?", "CONFIRME LA ELIMINACION", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                MessageBox.Show("Elimacion cancelada!!", "CANCELADO", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                LimpiarForm(grpDatosHoja);
                return;
            }

            try
            {
                nuevaAccion.Eliminar(IdRegistroInsumo, "EliminarRegistroInsumo");
                MessageBox.Show("Se elimino el registro", "OPERACION EXITOSA", MessageBoxButtons.OK, MessageBoxIcon.Information);
                MostrarDatosHoja();
                LimpiarForm(grpDatosHoja);
                Reiniciar();
            }
            catch (Exception ex)
            {
                MessageBox.Show("No se pudo eliminar el registro: " + ex, "OCURRIO UN PROBLEMA", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void btnReportes_Click(object sender, EventArgs e)
        {
            AbrirForm(new ReporteModulos());
        }
        #endregion

        #region Eventos
        //Metodo que nos ayuda a cargar los datos dependiendo el folio que se seleccione
        private void txtFolioSeleccionado_TextChanged(object sender, EventArgs e)
        {
            int Folio = int.Parse(txtFolioSeleccionado.Text);
            MostrarPartes(Folio);
        }

        private void dtgFolioos_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            IdRegistroInsumo = int.Parse(dtgFolioos.CurrentRow.Cells[0].Value.ToString());
            txtFolioSeleccionado.Text = dtgFolioos.CurrentRow.Cells[1].Value.ToString();
            txtFolio.Text = dtgFolioos.CurrentRow.Cells[1].Value.ToString();
            dtpFecha.Value = DateTime.Parse(dtgFolioos.CurrentRow.Cells[2].Value.ToString());
            cboPersonaRecibido.SelectedItem = dtgFolioos.CurrentRow.Cells[3].Value.ToString();
            txtConcepto.Text = dtgFolioos.CurrentRow.Cells[4].Value.ToString();

            HabilitarBotonesInsumos(true);
            Modificar = true;
            LimpiarForm(grpDatosPartes);
        }

        private void dtgPartes_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            IdRegistroParte = int.Parse(dtgPartes.CurrentRow.Cells[0].Value.ToString());
            cboModelo.SelectedItem = dtgPartes.CurrentRow.Cells[1].Value.ToString();
            txtCantidad.Text = dtgPartes.CurrentRow.Cells[2].Value.ToString();
            rtxtDescripcion.Text = dtgPartes.CurrentRow.Cells[3].Value.ToString();
            string Estado = dtgPartes.CurrentRow.Cells[4].Value.ToString();
            if(Estado == "Nuevo")
            {
                radNuevo.Checked = true;
            }
            else
            {
                radUsado.Checked = true;
            }

            HabilitarBotonesPartes(true);
            ModificarParte = true;
        }

        private void cboModelo_SelectedIndexChanged(object sender, EventArgs e)
        {
            string NumeroParte = cboModelo.SelectedItem.ToString();
            rtxtDescripcion.Text = lgModulo.ObtenerDescripcionParteRicoh(NumeroParte);
        }
        #endregion

        #region MetodosLocales
        public void LimpiarForm(GroupBox grpDatos)
        {
            foreach (Control c in grpDatos.Controls)
            {
                if (c is TextBox || c is RichTextBox)
                {
                    c.Text = "";
                }
            }
            //dtpFecha.Value = DateTime.Now;
            //cboPersonaRecibido.SelectedIndex = 0;
            cboModelo.SelectedIndex = 0;
            Modificar = false;
        }
        private void AbrirForm(object formNuevo)
        {
            //Declaramos la forma
            Form fh = formNuevo as Form;

            //Mostramos la forma 
            fh.Show();
        }

        public void Reiniciar()
        {
            //Limpiamos el DataGridView de las partes
            dtgPartes.DataSource = null;
            dtgPartes.Refresh();
            IdRegistroInsumo = 0;
        }
        #endregion
    }
}
