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
    public partial class DatosEquipoVendido : Form
    {
        public DatosEquipoVendido(ReporteEquipo EquipoVendido, int IdEquipo, EquiposBodega equipo)
        {
            InitializeComponent();
            InicioAplicacion();
            cboMarcas.SelectedItem = EquipoVendido.Marca;
            cboModelos.SelectedItem = EquipoVendido.Modelo;
            txtSerie.Text = EquipoVendido.Serie;
            txtPrecioEquipo.Text = EquipoVendido.Precio.ToString();
            this.IdEquipo = IdEquipo;
            this.equipo = equipo;
        }
        AccionesLógica NuevaAccion = new AccionesLógica();
        CD_Conexion cn = new CD_Conexion();
        LogicaEquipos lgEquipo = new LogicaEquipos();
        int IdEquipo = 0;
        EquiposBodega equipo = new EquiposBodega();


        #region Inicio
        public void InicioAplicacion()
        {
            LlenarComboBox(cboMarcas, "SeleccionarMarca", 0);
            LlenarComboBox(cboModelos, "SeleccionarModelos", 0);
            LlenarComboBox(cboClientes, "SeleccionarClientesServicios", 0);


            //Deshabilitamos escritura en combobox
            cboMarcas.DropDownStyle = ComboBoxStyle.DropDownList;
            cboModelos.DropDownStyle = ComboBoxStyle.DropDownList;

            txtSerie.Enabled = false;
            cboMarcas.Enabled = false;
            cboModelos.Enabled = false;
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
        #endregion

        #region Validaciones
        public bool ValidarCampos()
        {
            bool Validado = true;
            erEquipos.Clear();

            foreach (Control c in grpDatos.Controls)
            {
                if (c is TextBox || c is ComboBox)
                {
                    if (string.IsNullOrEmpty(c.Text) || c.Text == " ")
                    {
                        erEquipos.SetError(c, "Campo Obligatorio");
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
            bool SerieRepetida = false;
            if (!ValidarCampos())
                return;
            try
            {
                Equipo nuevoEquipo = new Equipo()
                {
                    IdCliente = NuevaAccion.BuscarId(cboClientes.SelectedItem.ToString(), "ObtenerIdCliente"),
                    IdMarca = NuevaAccion.BuscarId(cboMarcas.SelectedItem.ToString(), "ObtenerIdMarca"),
                    IdModelo = NuevaAccion.BuscarId(cboModelos.SelectedItem.ToString(), "ObtenerIdModelo"),
                    Serie = txtSerie.Text,
                    Precio = double.Parse(txtPrecioEquipo.Text.Replace(",", "")),
                    FechaVenta = dtpFecha.Value,
                };
                lgEquipo.GuardarEquipoVendido(nuevoEquipo);
                NuevaAccion.Eliminar(IdEquipo, "EliminarEquipoBodega");
                MessageBox.Show("Equipo vendido", "OPERACION EXITOSA", MessageBoxButtons.OK, MessageBoxIcon.Information);

                equipo.LimpiarForm();
                equipo.MostrarDatosEquipos();
                equipo.Refresh();
                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ah ocurrido un error " + ex.Message);
            }
        }
        #endregion

        private void btnCerrar_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
