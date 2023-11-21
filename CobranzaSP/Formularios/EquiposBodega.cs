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
    public partial class EquiposBodega : Form
    {
        public EquiposBodega()
        {
            InitializeComponent();
            InicioAplicacion();
        }
        AccionesLógica nuevaAccion = new AccionesLógica();
        CD_Conexion cn = new CD_Conexion();
        LogicaEquiposBodega lgEquipoBodega = new LogicaEquiposBodega();
        EquipoBodega equipo = new EquipoBodega();
        ReporteEquipo nuevoEquipoRentado;
        int IdEquipo = 0;
        bool EstaModificando = false;
        bool BuscandoFolio = false;
        string TipoBusqueda = "";
        #region Inicio

        public void InicioAplicacion()
        {
            ControlesDesactivadosInicialmente(false);

            LlenarComboBox(cboMarcas, "SeleccionarMarca", 0);
            
            LlenarComboBox(cboModelos, "SeleccionarModelos", 0);

            PropiedadesDtgEquipos();

            MostrarDatosEquipos();

            //Deshabilitamos escritura en combobox
            cboOpcionMostrar.DropDownStyle = ComboBoxStyle.DropDownList;
            cboBusqueda.DropDownStyle = ComboBoxStyle.DropDownList;

            LlenarCboOpcionesMostrar();
            rdTodasLasMarcas.Checked = true;
            rdTodosLosModelos.Checked = true;

            btnMostrar.Enabled = false;
        }

        public void LlenarCboOpcionesMostrar()
        {
            string[] Opciones = { "", "Marca", "Modelo" };
            cboOpcionMostrar.Items.AddRange(Opciones);
            cboOpcionMostrar.SelectedIndex = 0;
        }

        private void ControlesDesactivadosInicialmente(bool activado)
        {
            btnCancelar.Enabled = activado;
            btnEliminar.Enabled = activado;
            btnRentada.Enabled = activado;
            btnVentaEquipo.Enabled = activado;
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

        public void PropiedadesDtgEquipos()
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
        }

        public void MostrarDatosEquipos()
        {
            //Limpiamos los datos del datagridview
            dtgEquipos.DataSource = null;
            dtgEquipos.Refresh();
            DataTable tabla = new DataTable();
            //Guardamos los registros dependiendo la consulta
            tabla = nuevaAccion.Mostrar("MostrarEquiposBodega");
            //Asignamos los registros que optuvimos al datagridview
            dtgEquipos.DataSource = tabla;
            dtgEquipos.Columns["IdEquipo"].Visible = false;
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

        public bool ValidarCamposReporte()
        {
            bool Validado = true;
            erEquipos.Clear();
            if (string.IsNullOrWhiteSpace(cboOpcionMostrar.SelectedItem.ToString()))
            {
                erEquipos.SetError(cboOpcionMostrar, "Eliga una opcion");
                Validado = false;
            }
            else
            {
                if (string.IsNullOrWhiteSpace(cboBusqueda.SelectedItem.ToString()))
                {
                    erEquipos.SetError(cboBusqueda, "Eliga una opcion");
                    Validado = false;
                }
            }
            return Validado;
        }

        public bool ValidarMarcaSeleccionada()
        {
            bool Validado = true;
            erEquipos.Clear();

            if (string.IsNullOrWhiteSpace(cboBusqueda.SelectedItem.ToString()))
            {
                erEquipos.SetError(cboBusqueda, "Eliga una marca");
                Validado = false;
            }
            return Validado;
        }

        public bool ValidarModeloSeleccionado()
        {
            bool Validado = true;
            erEquipos.Clear();

            if (string.IsNullOrWhiteSpace(cboModelo.SelectedItem.ToString()))
            {
                erEquipos.SetError(cboModelo, "Eliga una modelo");
                Validado = false;
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
        private void btnGuardar_Click(object sender, EventArgs e)
        {
            string Mensaje;
            bool SerieRepetida;
            try
            {
                if (!ValidarCampos())
                    return;

                EquipoBodega NuevoEquipo = new EquipoBodega()
                {
                    IdEquipo = IdEquipo,
                    IdMarca = nuevaAccion.BuscarId(cboMarcas.SelectedItem.ToString(), "ObtenerIdMarca"),
                    IdModelo = nuevaAccion.BuscarId(cboModelos.SelectedItem.ToString(), "ObtenerIdModelo"),
                    Serie = txtSerie.Text,
                    Precio = double.Parse(txtPrecio.Text),
                    Notas = rtxtNotas.Text
                };
                NuevoEquipo.Estado = (rdNueva.Checked) ?"Nueva":"Usada";
                if (EstaModificando)
                {
                    if (MessageBox.Show("¿Esta seguro de modificar el registro?", "CONFIRME LA MODIFICACION", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                    {
                        MessageBox.Show("Modificacion cancelada!!", "CANCELADO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        LimpiarForm();
                        return;
                    }
                    Mensaje = lgEquipoBodega.GuardarEquipo(NuevoEquipo, "ModificarEquipoBodega");
                    ReiniciarControles();
                }
                else
                {
                    SerieRepetida = nuevaAccion.VerificarDuplicados(NuevoEquipo.Serie, "VerificarSerieEquiposBodega");
                    if (SerieRepetida)
                    {
                        MessageBox.Show("Ingrese un numero de serie distinto", "SERIE YA EXISTENTE", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        LimpiarForm();
                        return;
                    }
                    Mensaje = lgEquipoBodega.GuardarEquipo(NuevoEquipo, "AgregarEquipoABodega");
                }
                MessageBox.Show(Mensaje, "REGISTRO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                LimpiarForm();
                MostrarDatosEquipos();

            }
            catch (Exception ex)
            {
                MessageBox.Show("No se pudieron guardar los datos por: " + ex);
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
                nuevaAccion.Eliminar(IdEquipo, "EliminarEquipoBodega");
                MessageBox.Show("Se elimino el registro", "OPERACION EXITOSA", MessageBoxButtons.OK, MessageBoxIcon.Information);
                MostrarDatosEquipos();
                LimpiarForm();
                ReiniciarControles();
            }
            catch (Exception ex)
            {
                MessageBox.Show("No se pudo eliminar el registro: " + ex, "OCURRIO UN PROBLEMA", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void btnCancelar_Click(object sender, EventArgs e)
        {
            LimpiarForm();
            ReiniciarControles();
        }

        private void btnMostrar_Click(object sender, EventArgs e)
        {
            string Parametro = "";
            //PENDIENTE EL PDF AUN
            bool DatosEncontrados = true;


            if (rdTodasLasMarcas.Checked)
            {
                //Generamos el reporte completo de todas las marcas
                TipoBusqueda = "Todas";
                DatosEncontrados = lgEquipoBodega.ObtenerDatosTodasLasMarcas(TipoBusqueda,Parametro);
            }
            else
            {
                //Generaremos por una marca en especifico
                if (!ValidarMarcaSeleccionada())
                    return;
                Parametro = cboBusqueda.SelectedItem.ToString();

                if (rdTodosLosModelos.Checked)
                {
                    //Generamos reporte por la marca seleccionada
                    int IdMarca = nuevaAccion.BuscarId(Parametro, "ObtenerIdMarca");
                    DatosEncontrados = lgEquipoBodega.ObtenerDatosMarcaElegida(TipoBusqueda, Parametro, IdMarca);
                }
                else
                {
                    if (!ValidarModeloSeleccionado())
                        return;
                    TipoBusqueda = "Modelo";
                    Parametro = cboModelo.SelectedItem.ToString();
                    int IdModelo = nuevaAccion.BuscarId(Parametro, "ObtenerIdModelo");
                    DatosEncontrados = lgEquipoBodega.ObtenerDatosModeloElegida(TipoBusqueda, Parametro, IdModelo);
                }
            }

            if (!DatosEncontrados)
            {
                MessageBox.Show("¡NO SE ENCONTRARON EQUIPOS EN BODEGA!", "EQUIPOS NO ENCONTRADOS", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void ReiniciarControles()
        {
            ControlesDesactivadosInicialmente(false);
            EstaModificando = false;
        }
        #endregion

        #region Eventos
        private void cboMarcas_SelectedIndexChanged(object sender, EventArgs e)
        {
            //En dado caso que se haya seleccionado algo de las marcas y mientras no estemos buscando un registro en especifico
            if (cboMarcas.SelectedItem.ToString() != " " && BuscandoFolio == false)
            {
                int IdMarca = nuevaAccion.BuscarId(cboMarcas.SelectedItem.ToString(), "ObtenerIdMarca");
                LlenarComboBox(cboModelos, "SeleccionarModelos", IdMarca);
            }
        }
        #endregion

        #region MetodosLocales
        public void LimpiarForm()
        {
            foreach (Control c in grpDatos.Controls)
            {
                if (c is TextBox || c is RichTextBox)
                {
                    c.Text = "";
                }
            }
            cboMarcas.SelectedIndex = 0;
            cboModelos.SelectedIndex = 0;
        }
        #endregion

        private void dtgEquipos_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            LimpiarForm();
            //Una vez que se escoga alguna fila podremos activar los botones para poder modificar y eliminar
            btnGuardar.Enabled = true;
            ControlesDesactivadosInicialmente(true);
            EstaModificando = true;

            //Asignacion a los controles
            IdEquipo = int.Parse(dtgEquipos.CurrentRow.Cells[0].Value.ToString());
            cboMarcas.SelectedItem = dtgEquipos.CurrentRow.Cells[1].Value.ToString();
            cboModelos.SelectedItem = dtgEquipos.CurrentRow.Cells[2].Value.ToString();
            txtSerie.Text = dtgEquipos.CurrentRow.Cells[3].Value.ToString();
            txtPrecio.Text = dtgEquipos.CurrentRow.Cells[4].Value.ToString().Replace("$", "");
            string Estado = dtgEquipos.CurrentRow.Cells[5].Value.ToString();


            nuevoEquipoRentado = new ReporteEquipo()
            {
                Marca = dtgEquipos.CurrentRow.Cells[1].Value.ToString(),
                Modelo = dtgEquipos.CurrentRow.Cells[2].Value.ToString(),
                Serie = txtSerie.Text,
                Precio = double.Parse(txtPrecio.Text)
            };

            if (Estado == "Nueva")
            {
                rdNueva.Checked= true;
            }
            else
            {
                rdUsada.Checked = true;
            }

            rtxtNotas.Text = dtgEquipos.CurrentRow.Cells[6].Value.ToString();
        }

        private void cboOpcionMostrar_SelectedIndexChanged(object sender, EventArgs e)
        {
            btnMostrar.Enabled = true;
            TipoBusqueda = cboOpcionMostrar.SelectedItem.ToString();
            switch (TipoBusqueda)
            {
                case "Marca":
                    MostrarOpcionesMarca();
                    break;
                case "Modelo": LlenarComboBox(cboBusqueda, "SeleccionarModelos", 0); break;
                default: MostrarComboBoxBusqueda(false); break;
            }
        }

        public void MostrarOpcionesMarca()
        {
            rdUnaMarca.Visible = true;
            rdTodasLasMarcas.Visible = true;
            rdTodasLasMarcas.Checked = true;
        }

        public void MostrarOpcionesModelo()
        {
            grpModelo.Visible = true;
        }

        public void MostrarComboBoxBusqueda(bool mostrar)
        {
            cboBusqueda.Visible = mostrar;
        }

        private void rdUnaMarca_CheckedChanged(object sender, EventArgs e)
        {
            if (rdUnaMarca.Checked)
            {
                cboBusqueda.Visible = true;
                LlenarComboBox(cboBusqueda, "SeleccionarMarca", 0);
                rdUnaMarca.Checked = true;
                
            }
            else
            {
                grpModelo.Visible = false;
                cboBusqueda.Visible = false;
            }
        }

        private void cboBusqueda_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cboBusqueda.SelectedItem.ToString() != " ")
            {
                MostrarOpcionesModelo();
                int IdMarca = nuevaAccion.BuscarId(cboBusqueda.SelectedItem.ToString(), "ObtenerIdMarca");
                LlenarComboBox(cboModelo, "SeleccionarModelos", IdMarca);
            }
        }

        private void radUnModelo_CheckedChanged(object sender, EventArgs e)
        {
            if (radUnModelo.Checked)
            {
                cboModelo.Visible = true;
            }
            else
            {
                cboModelo.Visible = false;
            }
        }

        private void txtPrecio_KeyPress(object sender, KeyPressEventArgs e)
        {

            if (e.KeyChar == '.')
            {
                // Verificar si ya hay un punto en el texto
                if (txtPrecio.Text.Contains("."))
                {
                    // Ignorar el ingreso adicional de puntos
                    e.Handled = true;
                }
            }
            else
            {
                // Permitir el ingreso de otras teclas
                Validacion.SoloNumeros(e);
            }
        }

        private void btnRentada_Click(object sender, EventArgs e)
        {
            DatosImpresoraRentada forma = new DatosImpresoraRentada(this,nuevoEquipoRentado,IdEquipo);
            forma.Show();
        }

        private void btnVentaEquipo_Click(object sender, EventArgs e)
        {
            EquiposBodega equipo = new EquiposBodega();
            DatosEquipoVendido forma = new DatosEquipoVendido(nuevoEquipoRentado, IdEquipo, equipo);
            forma.Show();
        }
    }
}
