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
    public partial class Servicios : Form
    {
        public Servicios()
        {
            InitializeComponent();
            InicioAplicacion();
        }
        AccionesLógica nuevaAccion = new AccionesLógica();
        CD_Conexion cn = new CD_Conexion();
        LogicaServicio lgServicio = new LogicaServicio();
        LogicaRegistroModulos lgModulo = new LogicaRegistroModulos();
        string NumeroFolio;
        bool Validado = true;

        //Sabremos cuando estamos añadiendo un nuevo registro o modificando
        bool Modificar = false;
        bool BuscandoFolio = false;

        #region Inicio
        public void InicioAplicacion()
        {
            //Asignamos propiedades iniciales a nuestro datagriwview
            PropiedadesDtgServicios();

            //Desactivamos controles que no ocuparemos al inicio
            ControlesDesactivadosInicialmente(false);

            //Agregamos opciones a los combobox

            LlenarComboBox(cboClientes, "SeleccionarClientesServicios", 0);
            LlenarComboBox(cboMarca, "SeleccionarMarca", 0);
            LlenarComboBox(cboModelos, "SeleccionarModelos", 0);
            LlenarComboBox(cboFusor, "SeleccionarFusores", 0);
            LlenarComboBox(cboFusorRetirado, "SeleccionarFusores", 0);

            //Denegar escritura en combobox
            cboMarca.DropDownStyle = ComboBoxStyle.DropDownList;
            //cboMostrar.DropDownStyle = ComboBoxStyle.DropDownList;
            cboClientes.DropDownStyle = ComboBoxStyle.DropDownList;
            cboModelos.DropDownStyle = ComboBoxStyle.DropDownList;

            //Mostramos los registros que tengamos en nuestra base de datos de los reportes de los servicos
            MostrarDatosServicios();

            dtpFecha.MaxDate = DateTime.Today;

            string[] Tecnicos = { "", "ALVARO", "MIGUEL", };
            cboTecnico.Items.AddRange(Tecnicos);

            string[] OpcionesMostrar = { "", "Ultima Semana", "Ultimo Mes", "Mes Pasado", "Este año", "Todos" };
            cboMostrar.Items.AddRange(OpcionesMostrar);
        }


        //Metodo que ayuda a llenar los combobox dependiendo el stop procedure que se ejecute
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

        private void ControlesDesactivadosInicialmente(bool activado)
        {
            btnCancelar.Enabled = activado;
            btnEliminar.Enabled = activado;
        }

        //Opciones combobox Mostrar

        public void PropiedadesDtgServicios()
        {
            //Solo lectura
            dtgServicios.ReadOnly = true;
            //No agregar renglones
            dtgServicios.AllowUserToAddRows = false;
            //No borrar renglones
            dtgServicios.AllowUserToDeleteRows = false;
            //Ajustar automaticamente el ancho de las columnas
            dtgServicios.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            //dtgServicios.AutoResizeColumns(DataGridViewAutoSizeColumnsMo‌​de.Fill);
            dtgServicios.AutoResizeColumns();
        }

        public void MostrarDatosServicios()
        {
            //Limpiamos los datos del datagridview
            dtgServicios.DataSource = null;
            dtgServicios.Refresh();
            DataTable tabla = new DataTable();
            //Guardamos los registros dependiendo la consulta
            tabla = nuevaAccion.Mostrar("SeleccionarTodosLosServicios");
            //Asignamos los registros que optuvimos al datagridview
            dtgServicios.DataSource = tabla;
        }
        #endregion

        #region Validaciones
        public bool ValidarCamposVacios()
        {
            erServicios.Clear();

            if (txtNumeroFolio.Text == "")
            {
                erServicios.SetError(txtNumeroFolio, "Ingrese número de folio");
                Validado = false;
            }
            if (txtContador.Text == "")
            {
                erServicios.SetError(txtContador, "Ingrese contador");
                Validado = false;
            }
            bool Fusor = chkFusor.Checked;
            foreach (Control c in grpDatos.Controls)
            {
                if (c is ComboBox || c is RichTextBox || c is TextBox)
                {
                    switch (c.Name)
                    {
                        case "cboFusor": ValidarCampoVacio(c, chkFusor.Checked); break;
                        case "cboFusorRetirado": ValidarCampoVacio(c, chkFusor.Checked); break;
                        case "cboNumeroSerie": ValidarCampoVacio(c, !chkSerie.Checked); break;
                        case "txtNumeroSerie": ValidarCampoVacio(c, chkSerie.Checked); break;
                        case "txtModelo": ValidarCampoVacio(c, checkBox1.Checked); break;
                        case "cboModelos": ValidarCampoVacio(c, !checkBox1.Checked); break;
                        default:
                            ValidarCampoVacio(c, true); break;
                    }

                }
            }
            return Validado;
        }


        public void ValidarCampoVacio(Control c, bool Habilitado)
        {
            if ((c.Text == "" || c.Text == " ") && Habilitado == true)
            {
                erServicios.SetError(c, "Campo Obligatorio");
                Validado = false;
            }
        }

        public bool ValidarCampoBusqueda()
        {
            bool Validado = true;
            erServicios.Clear();

            if (txtBusqueda.Text == "")
            {
                erServicios.SetError(txtBusqueda, "Campo obligatorio");
                Validado = false;
            }
            return Validado;
        }

        private void txtNumeroFolio_KeyPress(object sender, KeyPressEventArgs e)
        {
            Validacion.SoloLetrasYNumeros(e);
        }

        private void txtNumeroSerie_KeyPress(object sender, KeyPressEventArgs e)
        {
            Validacion.SoloLetrasYNumeros(e);
        }

        private void txtModelo_KeyPress(object sender, KeyPressEventArgs e)
        {
            Validacion.SoloLetrasYNumeros(e);
        }

        private void txtContador_KeyPress(object sender, KeyPressEventArgs e)
        {
            Validacion.SoloNumeros(e);
        }
        #endregion

        #region Botones

        public void VerificarUsoFusor(Servicio servicio)
        {
            if (!chkFusor.Checked)
            {
                servicio.Fusor = " ";
                servicio.FusorSaliente = " ";
                return;
            }

            string NombreCliente = cboClientes.SelectedItem.ToString();

            //Restaremos del inventario el fusor que se utilizo en base a su modelo

            servicio.Fusor = cboFusor.SelectedItem.ToString();
            servicio.FusorSaliente = cboFusorRetirado.SelectedItem.ToString();
            if (Modificar && servicio.Fusor != "S/N")
            {
                return;
            }
            LogicaRegistro AccionRegistro = new LogicaRegistro();
            //Con ayuda de la clave del fusor, podemos obtener a traves de su modelo el idcartucho 
            int IdCartucho = nuevaAccion.BuscarId(servicio.Fusor, "ObtenerModeloFusor");

            RegistroInventario registroFusor = new RegistroInventario()
            {
                Cliente = cboClientes.SelectedItem.ToString(),
                IdMarca = lgServicio.ObtenerMarcaFusor(IdCartucho),
                IdCartucho = IdCartucho,
                Fecha = servicio.Fecha,
                NumeroSerie = servicio.Fusor
            };

            //Definimos la cantidad de entrada y salida dependiendo si el cliente es "Speed Toner"
            registroFusor.CantidadSalida = (NombreCliente == "SPEED TONER") ? 0 : 1; // La cantidad de salida se establece en 1 si el nombre del cliente no es "SPEED TONER", de lo contrario se mantiene en 0
            registroFusor.CantidadEntrada = (NombreCliente == "SPEED TONER") ? 1 : 0; // La cantidad de entrada se establece en 1 si el nombre del cliente es "SPEED TONER", de lo contrario se mantiene en 0

            string Mensaje = AccionRegistro.AgregarRegistroInventario(registroFusor);
            MessageBox.Show(Mensaje, "REGISTRO INVENTARIO", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }


        public void VerificarNuevoModelo(Servicio servicio)
        {
            if (checkBox1.Checked)
            {
                lgServicio.AñadirModelo(txtModelo.Text, servicio.IdMarca);
                servicio.IdModelo = nuevaAccion.BuscarId(txtModelo.Text, "ObtenerIdModelo");
            }
            else
                servicio.IdModelo = nuevaAccion.BuscarId(cboModelos.SelectedItem.ToString(), "ObtenerIdModelo");
        }

        private void btnGuardar_Click(object sender, EventArgs e)
        {
            try
            {
                if (!ValidarCamposVacios())
                    return;

                string Mensaje;
                Servicio nuevoServicio = new Servicio()
                {
                    //NumeroFolio = NumeroFolio,
                    NumeroFolio = txtNumeroFolio.Text,
                    IdCliente = nuevaAccion.BuscarId(cboClientes.SelectedItem.ToString(), "ObtenerIdCliente"),
                    IdMarca = nuevaAccion.BuscarId(cboMarca.SelectedItem.ToString(), "ObtenerIdMarca"),
                    Contador = int.Parse(txtContador.Text.Replace(",", "")),
                    Fecha = dtpFecha.Value,
                    Tecnico = cboTecnico.SelectedItem.ToString(),
                    ServicioRealizado = rtxtServicio.Text,
                    ReporteFallo = rtxtFallas.Text
                };
                VerificarNuevoModelo(nuevoServicio);
                
                //Verifiicamos si es una serie existente o una nueva que debe de ser capturada por el usuario
                nuevoServicio.Serie = (chkSerie.Checked) ? txtNumeroSerie.Text : cboNumeroSerie.SelectedItem.ToString();

                if (Modificar)
                {
                    if (MessageBox.Show("¿Esta seguro de modificar el registro?", "CONFIRME LA MODIFICACION", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                    {
                        MessageBox.Show("Modificacion cancelada!!", "CANCELADO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        LimpiarForm();
                        return;
                    }
                    VerificarUsoFusor(nuevoServicio);
                    Mensaje = lgServicio.RegistroServicio(nuevoServicio, "ModificarServicio");
                    MessageBox.Show(Mensaje, "MODIFICANDO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    bool FolioRepetido = nuevaAccion.VerificarDuplicados(nuevoServicio.NumeroFolio, "VerificarFolioExistente");
                    if (FolioRepetido)
                    {
                        MessageBox.Show("El numero de folio ya existe!!", "DUPLICADO", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    else
                    {
                        VerificarUsoFusor(nuevoServicio);
                        Mensaje = lgServicio.RegistroServicio(nuevoServicio, "AgregarServicio");
                        MessageBox.Show(Mensaje, "REGISTRO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                LimpiarForm();
                cboNumeroSerie.SelectedIndex = 0;
                MostrarDatosServicios();
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
                string NumeroFolio = txtNumeroFolio.Text;
                lgServicio.EliminarRegistro(NumeroFolio, "EliminarServicio");
                MessageBox.Show("Se elimino el registro", "OPERACION EXITOSA", MessageBoxButtons.OK, MessageBoxIcon.Information);
                MostrarDatosServicios();
                LimpiarForm();
            }
            catch (Exception ex)
            {
                MessageBox.Show("No se pudo eliminar el registro: " + ex, "OCURRIO UN PROBLEMA", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void btnCancelar_Click(object sender, EventArgs e)
        {
            //Regresaremos a como era al iniciar el programa
            ControlesDesactivadosInicialmente(false);
            btnGuardar.Enabled = true;
            LimpiarForm();
            Modificar = false;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            AbrirForm(new ReporteServicio());
        }

        private void btnBusqueda_Click(object sender, EventArgs e)
        {
            if (!ValidarCampoBusqueda())
            {
                return;
            }
            Modificar = true;
            string str = txtBusqueda.Text;
            SqlDataReader dr;

            dr = nuevaAccion.Buscar(txtBusqueda.Text, "BuscarServicioFolio");
            BuscandoFolio = true;
            btnCancelar.Enabled = true;

            if (dr.Read())
            {
                txtNumeroFolio.Text = (dr[0].ToString());
                cboClientes.SelectedItem = (dr[1].ToString());
                cboMarca.SelectedItem = (dr[2].ToString());
                cboModelos.SelectedItem = (dr[3].ToString());
                if (cboNumeroSerie.Text != dr[4].ToString())
                {
                    chkSerie.Checked = true;
                    txtNumeroSerie.Text = dr[4].ToString();
                }
                else
                    cboNumeroSerie.SelectedItem = dr[4].ToString();
                txtContador.Text = (dr[5].ToString());
                DateTime FechaRegistro = Convert.ToDateTime((dr[6].ToString()));
                dtpFecha.Value = FechaRegistro;
                cboTecnico.SelectedItem = (dr[7].ToString());


                string Fusor = dr[8].ToString();
                if (Fusor != "" && Fusor != " ")
                {
                    chkFusor.Checked = true;
                    cboFusor.SelectedItem = dr[8].ToString();
                    cboFusorRetirado.SelectedItem = dr[9].ToString();
                }
                rtxtServicio.Text = (dr[10].ToString());
                rtxtFallas.Text = (dr[11].ToString());
                //Agregamos las opciones dependiendo los registros que nos devolvieron
            }
            else
            {
                MessageBox.Show("El número de folio no esta registrado en la base de datos", "REGISTRO NO ENCONTRADO", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

            dr.Close();
            cn.CerrarConexion();
            txtBusqueda.Text = "";
            BuscandoFolio = false;

        }
        #endregion

        #region Eventos
        private void chkSerie_CheckedChanged(object sender, EventArgs e)
        {
            cboNumeroSerie.Visible = !chkSerie.Checked;
            txtNumeroSerie.Visible = chkSerie.Checked;

            if (chkSerie.Checked)
                txtNumeroSerie.Focus();
            else
                cboNumeroSerie.Focus();
        }

        private void chkFusor_CheckedChanged(object sender, EventArgs e)
        {
            if (chkFusor.Checked)
                MostrarCamposFusor(true);
            else
                MostrarCamposFusor(false);
        }

        private void cboClientes_DropDownClosed(object sender, EventArgs e)
        {
            LlenarSeries();
        }

        public void LlenarSeries()
        {
            //bool Lleno = true;
            if (cboClientes.SelectedItem.ToString() != " " && BuscandoFolio == false)
            {
                int IdCliente = nuevaAccion.BuscarId(cboClientes.SelectedItem.ToString(), "ObtenerIdCliente");
                LlenarComboBox(cboNumeroSerie, "SeleccionarNumeroSerieEquipo", IdCliente);
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            txtModelo.Visible = checkBox1.Checked;
            cboModelos.Visible = !checkBox1.Checked;
        }

        private void dtgServicios_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            LimpiarForm();
            //Una vez que se escoga alguna fila podremos activar los botones para poder modificar y eliminar
            btnGuardar.Enabled = true;
            ControlesDesactivadosInicialmente(true);
            Modificar = true;

            //Asignacion a los controles
            txtNumeroFolio.Text = dtgServicios.CurrentRow.Cells[0].Value.ToString();
            NumeroFolio = dtgServicios.CurrentRow.Cells[0].Value.ToString();
            cboClientes.SelectedItem = dtgServicios.CurrentRow.Cells[1].Value.ToString();

            cboMarca.SelectedItem = dtgServicios.CurrentRow.Cells[2].Value.ToString();
            cboModelos.SelectedItem = dtgServicios.CurrentRow.Cells[3].Value.ToString();
            cboNumeroSerie.SelectedItem = dtgServicios.CurrentRow.Cells[4].Value.ToString();

            if (cboNumeroSerie.Text != dtgServicios.CurrentRow.Cells[4].Value.ToString())
            {
                chkSerie.Checked = true;
                txtNumeroSerie.Text = dtgServicios.CurrentRow.Cells[4].Value.ToString();
            }
            string Fusor = dtgServicios.CurrentRow.Cells[8].Value.ToString();
            if (Fusor != "" && Fusor != " ")
            {
                chkFusor.Checked = true;
                cboFusor.SelectedItem = dtgServicios.CurrentRow.Cells[8].Value.ToString();
                cboFusorRetirado.SelectedItem = dtgServicios.CurrentRow.Cells[9].Value.ToString();
            }

            txtContador.Text = dtgServicios.CurrentRow.Cells[5].Value.ToString();
            DateTime FechaRegistro = Convert.ToDateTime(dtgServicios.CurrentRow.Cells[6].Value.ToString());
            dtpFecha.Value = FechaRegistro;
            cboTecnico.SelectedItem = dtgServicios.CurrentRow.Cells[7].Value.ToString();

            rtxtServicio.Text = dtgServicios.CurrentRow.Cells[10].Value.ToString();
            rtxtFallas.Text = dtgServicios.CurrentRow.Cells[11].Value.ToString();
        }

        private void cboClientes_SelectedIndexChanged(object sender, EventArgs e)
        {
            LlenarSeries();
        }

        private void cboClientes_SelectionChangeCommitted(object sender, EventArgs e)
        {

        }

        private void cboMarca_SelectedIndexChanged(object sender, EventArgs e)
        {
            //En dado caso que se haya seleccionado algo de las marcas y mientras no estemos buscando un registro en especifico
            if (cboMarca.SelectedItem.ToString() != " " && BuscandoFolio == false)
            {
                int IdMarca = nuevaAccion.BuscarId(cboMarca.SelectedItem.ToString(), "ObtenerIdMarca");
                LlenarComboBox(cboModelos, "SeleccionarModelos", IdMarca);
            }
        }

        private void cboMostrar_SelectedIndexChanged(object sender, EventArgs e)
        {
            DataTable tabla = new DataTable();
            //Dependiendo la opcion se enviaran diferentes stop procedures al metodo mostrar
            switch (cboMostrar.SelectedItem.ToString())
            {
                case "Ultima Semana": tabla = nuevaAccion.Mostrar("MostrarUltimaSemana"); break;
                case "Mes Pasado": tabla = nuevaAccion.Mostrar("MostrarMesPasado"); break;
                case "Ultimo Mes": tabla = nuevaAccion.Mostrar("MostrarUltimoMes"); break;
                case "Este año": tabla = nuevaAccion.Mostrar("MostrarAñoActual"); break;
                case "Todos": tabla = nuevaAccion.Mostrar("SeleccionarTodosLosServicios"); break;
            }
            //Asignamos los registros a nuestro datagridview
            dtgServicios.DataSource = tabla;
        }
        #endregion

        #region MetodosLocales
        private void LimpiarForm()
        {
            foreach (Control c in grpDatos.Controls)
            {
                if (c is TextBox || c is RichTextBox)
                {
                    c.Text = "";
                }
            }

            cboFusor.SelectedIndex = 0;
            cboFusorRetirado.SelectedIndex = 0;
            cboClientes.SelectedIndex = 0;
            cboMarca.SelectedIndex = 0;
            cboTecnico.SelectedIndex = 0;

            dtpFecha.Value = DateTime.Today;
            //Y volvemos a llenar todo el combobox con todos los modelos
            LlenarComboBox(cboModelos, "SeleccionarModelos", 0);
            chkSerie.Checked = false;
            chkFusor.Checked = false;
            checkBox1.Checked = false;
            txtNumeroFolio.Focus();
        }
        public void MostrarCamposFusor(bool Mostrar)
        {
            lblFusorInstalado.Visible = Mostrar;
            lblFusorRetirado.Visible = Mostrar;
            cboFusor.Visible = Mostrar;
            cboFusorRetirado.Visible = Mostrar;
        }



        private void AbrirForm(object formNuevo)
        {
            //Declaramos la forma
            Form fh = formNuevo as Form;

            //Mostramos la forma 
            fh.Show();
        }



        #endregion

        private void btnreportes2_Click(object sender, EventArgs e)
        {
            AbrirForm(new ServicioReportes());
        }
    }
}
