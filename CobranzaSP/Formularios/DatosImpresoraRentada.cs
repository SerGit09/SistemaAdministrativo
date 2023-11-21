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
    public partial class DatosImpresoraRentada : Form
    {
        private EquiposBodega EquipoBodega;
        public DatosImpresoraRentada(EquiposBodega EquipoBodega, ReporteEquipo nuevoEquipo, int Id)
        {
            InitializeComponent();
            InicioAplicacion();
            this.EquipoBodega = EquipoBodega;
            idEquipo = Id;
            //Cargamos los datos que teniamos desde bodega
            txtPrecioEquipo.Text = nuevoEquipo.Precio.ToString();
            cboMarcas.SelectedItem = nuevoEquipo.Marca;
            cboModelos.SelectedItem = nuevoEquipo.Modelo;
            txtSerie.Text = nuevoEquipo.Serie.ToString();
        }
        AccionesLógica NuevaAccion = new AccionesLógica();
        CD_Conexion cn = new CD_Conexion();
        LogicaEquipos lgEquipo = new LogicaEquipos();
        int idEquipo;


        #region Inicio
        public void InicioAplicacion()
        {
            LlenarComboBox(cboClientes, "SeleccionarClientesServicios", 0);
            LlenarComboBox(cboTipoRenta, "SeleccionarTipoRenta", 0);
            LlenarComboBox(cboMarcas, "SeleccionarMarca", 0);
            LlenarComboBox(cboModelos, "SeleccionarModelos", 0);

            //Deshabilitamos escritura en combobox
            cboClientes.DropDownStyle = ComboBoxStyle.DropDownList;
            cboTipoRenta.DropDownStyle = ComboBoxStyle.DropDownList;
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
        #endregion

        #region Botones

        private void btnCerrar_Click(object sender, EventArgs e)
        {
            this.Close();
        }

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
                    Ubicacion = txtReferencia.Text,
                    IdMarca = NuevaAccion.BuscarId(cboMarcas.SelectedItem.ToString(), "ObtenerIdMarca"),
                    IdModelo = NuevaAccion.BuscarId(cboModelos.SelectedItem.ToString(), "ObtenerIdModelo"),
                    Serie = txtSerie.Text,
                    IdRenta = NuevaAccion.BuscarId(cboTipoRenta.SelectedItem.ToString(), "ObtenerIdTipoRenta"),
                    Precio = double.Parse(txtPrecio.Text.Replace(",", "")),
                    FechaPago = txtFechaPago.Text,
                    Valor = double.Parse(txtPrecioEquipo.Text)
                };
                SerieRepetida = NuevaAccion.VerificarDuplicados(nuevoEquipo.Serie, "VerificarDuplicadoSerieEquipos");
                if (SerieRepetida)
                {
                    MessageBox.Show("Ingrese un numero de serie distinto", "SERIE YA EXISTENTE", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    lgEquipo.GuardarRegistro(nuevoEquipo, "AgregarEquipo");
                    NuevaAccion.Eliminar(idEquipo, "EliminarEquipoBodega");
                    MessageBox.Show("Equipo agregado correctamente", "OPERACION EXITOSA", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                //Limpiamos entonces el equipo que seleccionamos en bodega y que ahora se ira a equipos en renta
                EquipoBodega.LimpiarForm();
                EquipoBodega.MostrarDatosEquipos();
                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ah ocurrido un error " + ex.Message);
            }
        }

        #endregion
    }
}
