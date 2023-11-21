using CobranzaSP.Lógica;
using CobranzaSP.Modelos;
using DocumentFormat.OpenXml.Vml.Spreadsheet;
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
    public partial class InventarioModulos : Form
    {
        public InventarioModulos()
        {
            InitializeComponent();
            InicioAplicacion();
        }
        bool Inventario = true;
        bool EstaModificando = false;
        int Id = 0;
        AccionesLógica NuevaAccion = new AccionesLógica(); 
        CD_Conexion cn = new CD_Conexion();
        LogicaInventarioModulos lgInventarioModulo = new LogicaInventarioModulos();
        LogicaModulosBodega lgModuloBodega = new LogicaModulosBodega();


        #region Inicio
        public void InicioAplicacion()
        {
            LlenarComboBox(cboModelo, "SeleccionarModelosInventarioModulos", 0);

            MostrarDatosModulos();
            ControlesDesactivados(false);
            PropiedadesDtgModulos();
        }

        public void ControlesDesactivados(bool Desactivado)
        {
            btnEliminar.Enabled = Desactivado;
            btnCancelar.Enabled = Desactivado;
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

        public void PropiedadesDtgModulos()
        {
            //Solo lectura
            dtgModulos.ReadOnly = true;
            //No agregar renglones
            dtgModulos.AllowUserToAddRows = false;
            //No borrar renglones
            dtgModulos.AllowUserToDeleteRows = false;
            //Ajustar automaticamente el ancho de las columnas
            dtgModulos.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            //dtgServicios.AutoResizeColumns(DataGridViewAutoSizeColumnsMo‌​de.Fill);
            dtgModulos.AutoResizeColumns();
        }

        public void MostrarDatosModulos()
        {
            //Limpiamos los datos del datagridview
            dtgModulos.DataSource = null;
            dtgModulos.Refresh();
            DataTable tabla = new DataTable();
            //Guardamos los registros dependiendo la consulta
            tabla = NuevaAccion.Mostrar("MostrarModulosBodega");
            //Asignamos los registros que optuvimos al datagridview
            dtgModulos.DataSource = tabla;
            dtgModulos.Columns["Id"].Visible = false;
        }
        #endregion

        #region Validaciones
        public bool ValidarClave()
        {
            bool Validado = true;
            erInventarioModulos.Clear();

            foreach (Control c in grpDatosInventario.Controls)
            {
                if (c is TextBox)
                {
                    if (string.IsNullOrEmpty(c.Text))
                    {
                        erInventarioModulos.SetError(c, "Campo Obligatorio");
                        Validado = false;
                    }
                }
                if (c is ComboBox)
                {
                    if (((ComboBox)c).SelectedIndex == 0)
                    {
                        erInventarioModulos.SetError(c, "Campo Obligatorio");
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
            if (Inventario)
            {
                SeccionCatalogo();
            }
            else
            {
                //Seccion de registro de inventario
            }
        }

        #region Secciones
        public void SeccionCatalogo()
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
                    Resultado = lgModuloBodega.RegistrarModulo(NuevoModulo, "ModificarModuloBodega");
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
        #endregion

        private void btnCancelar_Click(object sender, EventArgs e)
        {
            Reiniciar();
        }


        public void Reiniciar()
        {
            EstaModificando = false;
            ControlesDesactivados(false);
            LimpiarForm(grpDatosRegistro);
            LimpiarForm(grpDatosInventario);
            Id = 0;
            cboModelo.SelectedIndex = 0;
            cboModulos.SelectedIndex = 0;
        }

        #endregion




        #region Eventos
        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tabControl1.SelectedTab == tabInventario)
            {
                Inventario = true;
                //Mostrar el inventario
            }
            else
            {
                Inventario = false;
                //Mostrar movimientos de inventario
            }
        }

        private void dtgModulos_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            EstaModificando = true;
            ControlesDesactivados(true);
            SeleccionarDatos();
        }

        public void SeleccionarDatos()
        {
            if (Inventario)
            {
                LlenarDatosInventario();
            }
            else
            {

            }
        }

        public void LlenarDatosInventario()
        {
            EstaModificando = true;
            ControlesDesactivados(true);
            Id = int.Parse(dtgModulos.CurrentRow.Cells[0].Value.ToString());
            cboModelo.SelectedItem = dtgModulos.CurrentRow.Cells[1].Value.ToString();
            cboModulos.SelectedItem = dtgModulos.CurrentRow.Cells[2].Value.ToString();
            txtClave.Text = dtgModulos.CurrentRow.Cells[3].Value.ToString();
        }
        #endregion

        #region MetodosLocales
        private void LimpiarForm(GroupBox grp)
        {
            foreach (Control c in grp.Controls)
            {
                if (c is TextBox)
                {
                    c.Text = "";
                }
            }
        }
        #endregion

        private void btnImprimir_Click(object sender, EventArgs e)
        {
            lgInventarioModulo.ImprimirInventario();
        }

        private void btnEliminar_Click(object sender, EventArgs e)
        {
            if (Inventario)
            {
                EliminarModuloCatalogo();
            }
        }

        public void EliminarModuloCatalogo()
        {
            //Preguntamos si se esta seguro de eliminar el registro 
            if (MessageBox.Show("¿Esta seguro de eliminar el registro?", "CONFIRME LA ELIMINACION", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                MessageBox.Show("Elimacion cancelada!!", "CANCELADO", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                LimpiarForm(grpDatosInventario);
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

        private void cboModelo_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cboModelo.SelectedItem.ToString() != " ")
            {
                int IdModelo = NuevaAccion.BuscarId(cboModelo.SelectedItem.ToString(), "ObtenerIdModeloModulo");
                LlenarComboBox(cboModulos, "SeleccionarModulos", IdModelo);
            }
        }
    }
}
