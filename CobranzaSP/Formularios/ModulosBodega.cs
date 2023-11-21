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
    public partial class ModulosBodega : Form
    {
        public ModulosBodega()
        {
            InitializeComponent();
            InicioAplicacion();
        }
        AccionesLógica NuevaAccion = new AccionesLógica();
        CD_Conexion cn = new CD_Conexion();
        LogicaModulosBodega lgModuloBodega = new LogicaModulosBodega();
        bool EstaModificando = false;
        int Id = 0;

        #region Inicio
        public void InicioAplicacion()
        {
            LlenarComboBox(cboModelo, "SeleccionarModelosInventarioModulos", 0);
            ControlesDesactivados(false);
            PropiedadesDtgModulos();
            MostrarDatosModulos();
        }

        public void ControlesDesactivados(bool Desactivado)
        {
            btnEliminar.Enabled = Desactivado;
            btnCancelar.Enabled = Desactivado;
        }

        public void PropiedadesDtgModulos()
        {
            //Solo lectura
            dtgModulosBodega.ReadOnly = true;
            //No agregar renglones
            dtgModulosBodega.AllowUserToAddRows = false;
            //No borrar renglones
            dtgModulosBodega.AllowUserToDeleteRows = false;
            //Ajustar automaticamente el ancho de las columnas
            dtgModulosBodega.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            //dtgServicios.AutoResizeColumns(DataGridViewAutoSizeColumnsMo‌​de.Fill);
            dtgModulosBodega.AutoResizeColumns();
        }

        public void MostrarDatosModulos()
        {
            //Limpiamos los datos del datagridview
            dtgModulosBodega.DataSource = null;
            dtgModulosBodega.Refresh();
            DataTable tabla = new DataTable();
            //Guardamos los registros dependiendo la consulta
            tabla = NuevaAccion.Mostrar("MostrarModulosBodega");
            //Asignamos los registros que optuvimos al datagridview
            dtgModulosBodega.DataSource = tabla;
            dtgModulosBodega.Columns["Id"].Visible = false;
        }

        public void LlenarComboBox(ComboBox cb, string sp, int Marca)
        {
            SqlDataReader dr;
            cb.Items.Clear();
            if (Marca != 0)
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

        #endregion

        #region Validaciones
        public bool ValidarClave()
        {
            bool Validado = true;
            erModulo.Clear();

            foreach (Control c in grpDatosModulo.Controls)
            {
                if (c is TextBox)
                {
                    if (string.IsNullOrEmpty(c.Text))
                    {
                        erModulo.SetError(c, "Campo Obligatorio");
                        Validado = false;
                    }
                }
                if (c is ComboBox)
                {
                    if (((ComboBox)c).SelectedIndex == 0)
                    {
                        erModulo.SetError(c, "Campo Obligatorio");
                        Validado = false;
                    }
                }
            }

            return Validado;
        }
        #endregion

        #region Botones
        private void btnGuardar_Click(object sender, EventArgs e)
        {
            if (!ValidarClave())
                return;
            try
            {
                string Resultado;
                ModuloBodega NuevoModulo = new ModuloBodega()
                {
                    Id = Id,
                    IdModelo = NuevaAccion.BuscarId(cboModelo.SelectedItem.ToString(), "ObtenerIdModeloModulo"),
                    Clave = txtClave.Text
                };
                NuevoModulo.IdModulo = lgModuloBodega.BuscarIdModulo(cboModulos.SelectedItem.ToString(), NuevoModulo.IdModelo);
                if (EstaModificando)
                {
                    if (MessageBox.Show("¿Esta seguro de modificar el registro?", "CONFIRME LA MODIFICACION", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                    {
                        MessageBox.Show("Modificacion cancelada!!", "CANCELADO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }
                    Resultado = lgModuloBodega.RegistrarModulo(NuevoModulo, "ModificarModuloAInventario");
                    MessageBox.Show(Resultado, "REGISTRO INVENTARIO MODULOS", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else//Agregamos al catalogo
                {
                    bool ClaveRepetida = NuevaAccion.VerificarDuplicados(NuevoModulo.Clave, "VerificarClaveDuplicadaBodega");
                    if (ClaveRepetida)
                    {
                        MessageBox.Show("¡¡EL NUMERO DE FOLIO YA EXISTE!!", "DUPLICADO", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    Resultado = lgModuloBodega.RegistrarModulo(NuevoModulo, "AgregarModuloBodega");
                    MessageBox.Show(Resultado, "REGISTRO INVENTARIO MODULOS", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                MostrarDatosModulos();
                Reiniciar();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ocurrio un error " + ex.Message);
            }
        }

        private void btnEliminar_Click(object sender, EventArgs e)
        {
            //Preguntamos si se esta seguro de eliminar el registro 
            if (MessageBox.Show("¿Esta seguro de eliminar el registro?", "CONFIRME LA ELIMINACION", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                MessageBox.Show("Elimacion cancelada!!", "CANCELADO", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                LimpiarForm();
                return;
            }

            try
            {
                NuevaAccion.Eliminar(Id, "EliminarModuloABodega");
                MessageBox.Show("SE HA ACTUALIZADO EL INVENTARIO", "CLAVE ELIMINADA", MessageBoxButtons.OK, MessageBoxIcon.Information);
                MostrarDatosModulos();
                Reiniciar();
            }
            catch (Exception ex)
            {
                MessageBox.Show("No se pudo eliminar la clave: " + ex, "OCURRIO UN PROBLEMA", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void ModulosBodega_Load(object sender, EventArgs e)
        {

        }

        private void btnCancelar_Click(object sender, EventArgs e)
        {
            Reiniciar();
        }
        #endregion

        #region Eventos
        private void cboModelo_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cboModelo.SelectedItem.ToString() != " ")
            {
                int IdModelo = NuevaAccion.BuscarId(cboModelo.SelectedItem.ToString(), "ObtenerIdModeloModulo");
                LlenarComboBox(cboModulos, "SeleccionarModulos", IdModelo);
            }
        }
        private void dtgModulosBodega_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void dtgModulosBodega_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            EstaModificando = true;
            ControlesDesactivados(true);
            Id = int.Parse(dtgModulosBodega.CurrentRow.Cells[0].Value.ToString());
            cboModelo.SelectedItem = dtgModulosBodega.CurrentRow.Cells[1].Value.ToString();
            cboModulos.SelectedItem = dtgModulosBodega.CurrentRow.Cells[2].Value.ToString();
            txtClave.Text = dtgModulosBodega.CurrentRow.Cells[3].Value.ToString();
        }
        #endregion

        #region MetodosLocales
        private void LimpiarForm()
        {
            foreach (Control c in grpDatosModulo.Controls)
            {
                if (c is TextBox)
                {
                    c.Text = "";
                }
            }
            cboModelo.SelectedIndex = 0;
            cboModulos.SelectedIndex = 0;
        }

        public void Reiniciar()
        {
            EstaModificando = false;
            ControlesDesactivados(false);
            LimpiarForm();
            Id = 0;
        }
        #endregion

    }
}
