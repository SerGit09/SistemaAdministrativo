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
using DocumentFormat.OpenXml.Office2010.Excel;

namespace CobranzaSP.Formularios
{
    public partial class Servicios_Ricoh : Form
    {
        public Servicios_Ricoh()
        {
            InitializeComponent();
            InicioAplicacion();
        }
        AccionesLógica nuevaAccion = new AccionesLógica();
        CD_Conexion cn = new CD_Conexion();
        LogicaServiciosRicoh lgServicioRicoh = new LogicaServiciosRicoh();
        LogicaServicio lgServicio = new LogicaServicio();
        LogicaRegistroModulos lgModulo = new LogicaRegistroModulos();
        LogicaPartesUsadas lgPartesUsadas = new LogicaPartesUsadas();
        LogicaModulosCliente lgModuloCliente = new LogicaModulosCliente();

        //Objeto para guardar los datos cuando buscamos un folio
        Modulo_Cliente ModuloSeleccionado;

        #region Variables
        //Generales
        string Serie = "";

        //Validaciones
        bool Validado = true;
        bool ValidadoPartes = true;

        //Modificaciones
        bool Modificar = false;
        bool ModificarParte = false;
        bool ModificarModulo = false;
        bool EstaModuloHabilitado = false;
        bool RetirarModulo = false;

        bool ModulosLlenos = false;
        bool BuscandoFolio = false;

        //Ids para modificar o eliminar
        int IdParte = 0;
        int IdServicio = 0;
        int Id = 0;
        #endregion

        #region Inicio

        public void InicioAplicacion()
        {
            PropiedadesDtgPartesRicoh();
            LlenarComboBox(cboClientes, "SeleccionarClientesServicios",0);
            LlenarComboBox(cboModelos, "SeleccionarModelosRicoh",0);

            LlenarComboBox(cboDescripciones, "SeleccionarDescripcionesPartesRicoh", 0);

            string[] Tecnicos = { "", "ALVARO", "MIGUEL", };
            cboTecnico.Items.AddRange(Tecnicos);
            HabilitarBotonesModulos(false);
            ControlesDesactivadosInicialmente(false);
            EstablecerBotonesPartesHabilitadas(false);
            txtClaveParte.Enabled = false;
            btnAgregar.Enabled = false;

            //dtgModulos.Columns["Id"].Visible = false;
            MostrarControlesBusquedaFolio(false);
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

        private void ControlesDesactivadosInicialmente(bool activado)
        {
            btnCancelar.Enabled = activado;
            btnEliminar.Enabled = activado;
        }

        private void EstablecerBotonesPartesHabilitadas(bool activado)
        {
            btnEliminarParte.Enabled = activado;
            btnCancelarParte.Enabled = activado;
        }

        public void PropiedadesDtgPartesRicoh()
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

        public void MostrarPartesUsadas(string NumeroFolio)
        {
            //Limpiamos los datos del datagridview
            dtgModulos.DataSource = null;
            dtgModulos.Refresh();
            DataTable tabla = new DataTable();
            //Guardamos los registros dependiendo la consulta
            tabla = lgPartesUsadas.MostrarPartes(NumeroFolio);
            //Asignamos los registros que optuvimos al datagridview
            dtgModulos.DataSource = tabla;
            dtgModulos.Columns["Id"].Visible = false;
            dtgModulos.Columns["Modelo"].Visible = false;
            dtgModulos.Columns["Serie"].Visible = false;
            
        }

        

        #endregion

        #region Validaciones
        private void txtNumeroFolio_KeyPress(object sender, KeyPressEventArgs e)
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

        private void txtCantidad_KeyPress(object sender, KeyPressEventArgs e)
        {
            Validacion.SoloNumeros(e);
        }
        private void txtBusqueda_KeyPress(object sender, KeyPressEventArgs e)
        {
            Validacion.SoloLetrasYNumeros(e);
        }

        public bool ValidarCamposVaciosServicios()
        {
            erServicios.Clear();
            Validado = true;

            foreach (Control c in grpDatos.Controls)
            {
                if (c is ComboBox || c is RichTextBox || c is TextBox)
                {
                    switch (c.Name)
                    {
                        case "cboNumeroSerie": ValidarCampoVacio(c, false); break;
                        case "txtModelo": ValidarCampoVacio(c, checkBox1.Checked); break;
                        case "cboModelos": ValidarCampoVacio(c, !checkBox1.Checked); break;
                        default:
                            ValidarCampoVacio(c, true); break;
                    }

                }
            }
            return Validado;
        }

        public bool ValidarCampoVacioBusqueda()
        {
            erServicios.Clear();
            Validado = true;

            if(txtBusqueda.Text == "")
            {
                erServicios.SetError(txtBusqueda, "Ingrese numero de folio");
                Validado = false;
            }
            return Validado;
        }

        public bool ValidarCamposVaciosPartesUsadas()
        {
            erServicios.Clear();

            foreach (Control c in grpDatosPartes.Controls)
            {
                if (c is ComboBox || c is RichTextBox || c is TextBox)
                {
                    if(c.Text == "" || c.Text == " ")
                    {
                        erServicios.SetError(c, "Campo obligatorio");
                        ValidadoPartes = false;
                    }
                }
            }
            return Validado;
        }
        //Metodo que ayuda a saber cuando el usuario presiona para agregar un modelo nuevo o si uso un dato de los combobox
        public void ValidarCampoVacio(Control c, bool Habilitado)
        {
            if ((c.Text == "" || c.Text == " ") && Habilitado == true)
            {
                erServicios.SetError(c, "Campo Obligatorio");
                Validado = false;
            }
        }
        #endregion

        #region Botones
        private void btnAgregarReporte_Click(object sender, EventArgs e)
        {
            BuscandoFolio = false;
            LogicaModulosCliente.ClavesObservaciones.Clear();
            grpDatos.Visible = true;
            txtNumeroFolio.Focus();
            MostrarControlesBusquedaFolio(false);
            Limpiar();
            Modificar = false;
        }

        public void MostrarControlesBusquedaFolio(bool mostrar)
        {
            lblNumeroFolio.Visible = mostrar;
            txtBusqueda.Visible = mostrar;
            btnBusqueda.Visible = mostrar;
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
                if (!ValidarCamposVaciosServicios())
                    return;

                string Mensaje;
                Servicio nuevoServicio = new Servicio()
                {
                    NumeroFolio = txtNumeroFolio.Text,
                    IdCliente = nuevaAccion.BuscarId(cboClientes.SelectedItem.ToString(), "ObtenerIdCliente"),
                    //IdModelo = nuevaAccion.BuscarId(cboModelos.SelectedItem.ToString(), "ObtenerIdModelo"),
                    Contador = int.Parse(txtContador.Text.Replace(",", "")),
                    Modelo = cboModelos.SelectedItem.ToString(),
                    IdMarca = 6,
                    Fecha = dtpFecha.Value,
                    Serie = cboNumeroSerie.SelectedItem.ToString(),
                    Tecnico = cboTecnico.SelectedItem.ToString(),
                    ServicioRealizado = rtxtServicio.Text,
                    ReporteFallo = rtxtFallas.Text,
                    Fusor = "",
                    FusorSaliente = ""
                };
                VerificarNuevoModelo(nuevoServicio);


                if (Modificar)
                {
                    if (MessageBox.Show("¿Esta seguro de modificar el registro?", "CONFIRME LA MODIFICACION", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                    {
                        MessageBox.Show("Modificacion cancelada!!", "CANCELADO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        dtpFecha.Value = DateTime.Now;
                        LimpiarForm(grpDatos);
                        return;
                    }
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
                        //Actualizamos los contadores de cada uno de los modulos
                        if (ModulosLlenos)
                        {
                            lgServicioRicoh.ActualizarModulosEquipo(nuevoServicio);
                        }

                        //Se tendran que tener un nuevo registro de esos modulos

                        //Una vez actualizados agregamos el reporte
                        Mensaje = lgServicio.RegistroServicio(nuevoServicio, "AgregarServicio");
                        MessageBox.Show(Mensaje, "REGISTRO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        tabControl1.Visible = false;
                    }
                }
                LimpiarForm(grpDatos);
                dtpFecha.Value = DateTime.Now;
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
                LimpiarForm(grpDatos);
                return;
            }

            try
            {
                string NumeroFolio = txtNumeroFolio.Text;
                lgServicio.EliminarRegistro(NumeroFolio, "EliminarServicio");
                MessageBox.Show("Se elimino el registro", "OPERACION EXITOSA", MessageBoxButtons.OK, MessageBoxIcon.Information);
                LimpiarForm(grpDatos);
            }
            catch (Exception ex)
            {
                MessageBox.Show("No se pudo eliminar el registro: " + ex, "OCURRIO UN PROBLEMA", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void btnCancelar_Click(object sender, EventArgs e)
        {
            ControlesDesactivadosInicialmente(false);
            LimpiarForm(grpDatos);
            Modificar = false;
            tabControl1.Visible = false;
            dtpFecha.Value = DateTime.Now;

            //Limpiamos el DataGridView de los modulos
            dtgModulos.DataSource = null;
            dtgModulos.Refresh();
        }

        private void btnBusqueda_Click(object sender, EventArgs e)
        {
            BuscandoFolio = true;
            Servicio ServicioBuscado = new Servicio();
            if (!ValidarCampoVacioBusqueda())
            {
                return;
            }
            Modificar = true;
            string str = txtBusqueda.Text;
            SqlDataReader dr;
            

            dr = nuevaAccion.Buscar(txtBusqueda.Text, "BuscarServicioFolio");
            //BuscandoFolio = true;
            btnCancelar.Enabled = true;

            if (dr.Read())
            {
                grpDatos.Visible = true;
                ServicioBuscado = new Servicio()
                {
                    //IdServicio = int.Parse(dr[0].ToString()),
                    NumeroFolio = (dr[0].ToString()),
                    Cliente = dr[1].ToString(),
                    Modelo = dr[3].ToString(),
                    Serie = dr[4].ToString(),
                    Contador = int.Parse(dr[5].ToString()),
                    Fecha = Convert.ToDateTime((dr[6].ToString())),
                    Tecnico = (dr[7].ToString()),
                    ServicioRealizado = dr[10].ToString(),
                    ReporteFallo = dr[11].ToString()
                };
                
                //Agregamos las opciones dependiendo los registros que nos devolvieron
            }
            else
            {
                MessageBox.Show("El número de folio no esta registrado en la base de datos", "REGISTRO NO ENCONTRADO", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                dr.Close();
                cn.CerrarConexion();
                return;
            }
            dr.Close();
            cn.CerrarConexion();

            //Llenamos los campos con los datos que guardamos en el objeto
            //IdServicio = ServicioBuscado.IdServicio;
            txtNumeroFolio.Text = ServicioBuscado.NumeroFolio;
            cboClientes.SelectedItem = ServicioBuscado.Cliente;
            cboModelos.SelectedItem = ServicioBuscado.Modelo;
            cboNumeroSerie.SelectedItem = ServicioBuscado.Serie;
            txtContador.Text = ServicioBuscado.Contador.ToString();
            dtpFecha.Value = ServicioBuscado.Fecha;
            cboTecnico.SelectedItem = ServicioBuscado.Tecnico;
            rtxtServicio.Text = ServicioBuscado.ServicioRealizado;
            rtxtFallas.Text = ServicioBuscado.ReporteFallo;


            txtBusqueda.Text = "";
            //MostrarPartesUsadas(txtNumeroFolio.Text);
        }

        private void btnBuscarReporte_Click(object sender, EventArgs e)
        {
            grpDatos.Visible = false;
            //Si no hay campos vacios, quiere decir que el formulario esta lleno
            Limpiar();
            tabControl1.Visible = false;
            //LimpiarForm(grpDatos);
            MostrarControlesBusquedaFolio(true);
            txtBusqueda.Focus();
        }

        #region Partes

        private void btnGuardarParte_Click(object sender, EventArgs e)
        {
            try
            {
                if (!ValidarCamposVaciosPartesUsadas())
                    return;

                string Mensaje;
                Parte nuevaParte = new Parte()
                {
                    Id = IdParte,
                    NumeroFolio = txtNumeroFolio.Text,
                    Cantidad = int.Parse(txtCantidad.Text),
                    Descripcion = cboDescripciones.SelectedItem.ToString(),
                    Estado = (radNuevo.Checked) ? "Nuevo" : "Usado"
                    //IdNumeroParte = nuevaAccion.BuscarId(cboDescripciones.SelectedItem.ToString(), "ObtenerIdModeloPartes")
                };


                if (ModificarParte)
                {
                    if (MessageBox.Show("¿Esta seguro de modificar el registro?", "CONFIRME LA MODIFICACION", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                    {
                        MessageBox.Show("Modificacion cancelada!!", "CANCELADO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        LimpiarForm(grpDatosPartes);
                        return;
                    }
                    Mensaje = lgPartesUsadas.RegistroParte(nuevaParte, "ModificarParteUsadaServicio");
                    MessageBox.Show(Mensaje, "MODIFICANDO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    Mensaje = lgPartesUsadas.RegistroParte(nuevaParte, "AgregarParteUsadaServicio");
                    MessageBox.Show(Mensaje, "REGISTRO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                LimpiarForm(grpDatosPartes);
                MostrarPartesUsadas(txtNumeroFolio.Text);
            }
            catch (Exception ex)
            {
                MessageBox.Show("No se pudieron guardar los datos por: " + ex);
            }
        }

        private void btnEliminarParte_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("¿Esta seguro de eliminar el registro?", "CONFIRME LA ELIMINACION", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                MessageBox.Show("Elimacion cancelada!!", "CANCELADO", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                LimpiarForm(grpDatosPartes);
                return;
            }

            try
            {
                string Mensaje;
                Parte nuevaParte = new Parte()
                {
                    Id = IdParte,
                    NumeroFolio = txtNumeroFolio.Text,
                    Cantidad = int.Parse(txtCantidad.Text),
                    Descripcion = cboDescripciones.SelectedItem.ToString(),
                    Estado = (radNuevo.Checked) ? "Nuevo" : "Usado",
                    IdNumeroParte = nuevaAccion.BuscarId(cboDescripciones.SelectedItem.ToString(), "ObtenerIdModeloPartes")
                };
                lgPartesUsadas.EliminarRegistroParteUsada(nuevaParte);
                MessageBox.Show("Resgistro eliminado correctamente, se ha actualizado el inventario de partes Ricoh", "REGISTRO ELIMINADO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                LimpiarForm(grpDatosPartes);
                MostrarPartesUsadas(txtNumeroFolio.Text);
            }
            catch (Exception ex)
            {
                MessageBox.Show("No se pudo eliminar e registro: " + ex);
            }
        }

        private void btnCancelarParte_Click(object sender, EventArgs e)
        {
            LimpiarForm(grpDatosPartes);
            EstablecerBotonesPartesHabilitadas(false);
            ModificarParte = false;
            cboDescripciones.SelectedIndex = 0;
        }
        #endregion
        #endregion

        #region Eventos
        private void cboModelo_SelectedIndexChanged(object sender, EventArgs e)
        {
            string Descripcion = cboDescripciones.SelectedItem.ToString();
            txtClaveParte.Text = lgServicioRicoh.ObtenerClaveParteRicoh(Descripcion);
        }

        private void cboClientes_SelectedIndexChanged(object sender, EventArgs e)
        {
            LlenarSeriesClienteSeleccionado();
        }
        private void dtgPartes_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            LimpiarForm(grpDatosPartes);
            EstablecerBotonesPartesHabilitadas(true);
            ModificarParte = true;

            IdParte = int.Parse(dtgModulos.CurrentRow.Cells[0].Value.ToString());
            cboDescripciones.SelectedItem = dtgModulos.CurrentRow.Cells[1].Value.ToString();
            txtCantidad.Text = dtgModulos.CurrentRow.Cells[2].Value.ToString();
            //rtxtDescripcion.Text = dtgPartes.CurrentRow.Cells[3].Value.ToString();
            string Estado = dtgModulos.CurrentRow.Cells[4].Value.ToString();
            if (Estado == "Nuevo")
            {
                radNuevo.Checked = true;
            }
            else
            {
                radUsado.Checked = true;
            }

            Modificar = true;
        }

        private void cboNumeroSerie_SelectedIndexChanged(object sender, EventArgs e)
        {
            ModulosLlenos = false;
            if (cboNumeroSerie.SelectedItem.ToString() != " ")
            {
                
                Serie = cboNumeroSerie.SelectedItem.ToString();
                
                MostrarModulosCliente(Serie,true);

                //Llenamos el modelo dependiendo del numero de Serie
                cboModelos.SelectedItem = lgServicioRicoh.ObtenerModeloEquipo(Serie);
                tabControl1.Visible = true;
                
                btnAgregar.Enabled = true;//Habilitamos para poder agregar un modulo
                EstaModuloHabilitado = true;
                return;
            }
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
                else if(c is ComboBox)
                {
                    ComboBox comboBox = (ComboBox)c;
                    comboBox.SelectedIndex = 0;
                }
            }
        }

        public void LlenarSeriesClienteSeleccionado()
        {
            if (cboClientes.SelectedItem.ToString() != " ")
            {
                int IdCliente = nuevaAccion.BuscarId(cboClientes.SelectedItem.ToString(), "ObtenerIdCliente");
                

                string Cliente = cboClientes.SelectedItem.ToString();
                if (Cliente == "SPEED TONER")
                {
                    LlenarComboBox(cboNumeroSerie, "SeleccionarSeriesEnBodegaRicoh", 0);
                }
                else
                {
                    LlenarComboBox(cboNumeroSerie, "SeleccionarNumeroSerieEquiposRicoh", IdCliente);
                }
            }
        }
        #endregion

        private void txtNumeroFolio_TextChanged(object sender, EventArgs e)
        {

        }
        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            
        }
        #region Modulos
        //Boton que nos abrira una forma para colocar un modulo nuevo
        private void btnAgregar_Click(object sender, EventArgs e)
        {
            int contador = int.Parse(txtContador.Text);
            ModuloNuevo modulo = new ModuloNuevo(this, cboModelos.SelectedItem.ToString(), Serie, txtNumeroFolio.Text,contador, ModificarModulo);
            modulo.Show();
        }

        //Abrira una forma, donde tendremos ya cargados los datos
        private void btnModificarModulo_Click(object sender, EventArgs e)
        {
            int Contador = int.Parse(txtContador.Text);
            string Modelo = cboModelos.SelectedItem.ToString();
            string NumeroFolio = txtNumeroFolio.Text;
            ModuloNuevo modulo = new ModuloNuevo(this, Modelo, ModuloSeleccionado, Serie, NumeroFolio,Contador, ModificarModulo,RetirarModulo,BuscandoFolio);
            modulo.Show();
        }

        private void btnCance_Click(object sender, EventArgs e)
        {
            ReiniciarCapturaModulo();
        }

        private void btnRetirarModulo_Click(object sender, EventArgs e)
        {
            ModuloNuevo modulo;
            int contador = int.Parse(txtContador.Text);
            RetirarModulo = true;
            if (ModificarModulo)
            {
                //Alguna variable para cargar
                modulo = new ModuloNuevo(this, cboModelos.SelectedItem.ToString(), ModuloSeleccionado, Serie, txtNumeroFolio.Text, contador, ModificarModulo, RetirarModulo,BuscandoFolio);
            }
            else
            {
                modulo = new ModuloNuevo(this, cboModelos.SelectedItem.ToString(), Serie, txtNumeroFolio.Text, contador, ModificarModulo, RetirarModulo);
            }

            modulo.Show();
        }

        public void ReiniciarCapturaModulo()
        {
            HabilitarBotonesModulos(false);
            ModificarModulo = false;
            btnAgregar.Enabled = true;
        }

        public void HabilitarBotonesModulos(bool habilitar)
        {
            btnModificarModulo.Enabled = habilitar;
            btnEliminarModulo.Enabled = habilitar;
        }

        private void dtgModulos_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            HabilitarBotonesModulos(true);
            btnAgregar.Enabled = false;
            ModificarModulo = true;

            ModuloSeleccionado = new Modulo_Cliente()
            {
                Id = int.Parse(dtgModulos.CurrentRow.Cells[0].Value.ToString()),
                //Modelo = dtgModulos.CurrentRow.Cells[1].Value.ToString(),
                Modulo = dtgModulos.CurrentRow.Cells[1].Value.ToString(),
                //Serie = dtgModulos.CurrentRow.Cells[3].Value.ToString(),
                Clave = dtgModulos.CurrentRow.Cells[2].Value.ToString(),
                //Paginas = int.Parse(txtContador.Text),
                Estado = dtgModulos.CurrentRow.Cells[4].Value.ToString(),
                Observacion = dtgModulos.CurrentRow.Cells[3].Value.ToString()
            };
        }

        public void MostrarModulosCliente(string NumeroSerie, bool LimpiarLista)
        {
            string FolioReporte = txtNumeroFolio.Text;
            //Limpiamos los datos del datagridview
            dtgModulos.DataSource = null;
            dtgModulos.Refresh();
            DataTable tabla = new DataTable();
            //Guardamos los registros dependiendo la consulta
            if (BuscandoFolio)
            {
                tabla = lgModuloCliente.MostrarModulosCliente(NumeroSerie,FolioReporte);
            }
            else
            {
                FolioReporte = "";
                tabla = lgModuloCliente.MostrarModulosCliente(NumeroSerie, FolioReporte);
            }


            //Asignamos los registros que optuvimos al datagridview
            dtgModulos.DataSource = tabla;
            
            dtgModulos.Columns["Id"].Visible = false;
            //dtgModulos.Columns["Modelo"].Visible = false;
            //dtgModulos.Columns["Contador"].Visible = false;
            //dtgModulos.Columns["Serie"].Visible = false;
            dtgModulos.Columns["Estado"].Visible = false;

            //AH PRUEBA PARA LO DEL COMENTARIO DE OBSERVACIONES
            //Si no estamos buscando un numero de folio de reporte
            if (!BuscandoFolio)
            {
                foreach (DataGridViewRow row in dtgModulos.Rows)
                {
                    string Clave = row.Cells["Clave"].Value.ToString();
                    //Solamente borraremos las observaciones de claves a las que nos se les haya asignado apenas una observación
                    if (!LogicaModulosCliente.ClavesObservaciones.Contains(Clave))
                    {
                        row.Cells["Observacion"].Value = "";  // Establecer el valor de la celda en blanco
                    }
                }
            }

            ObtenerDatosDeColumnas(LimpiarLista);
        }

        private void ObtenerDatosDeColumnas(bool LimpiarLista)
        {
            //Dependiendo el valor que mandemos decidiremos si limpiar o no la lista
            if (!LimpiarLista)
                return;

            if (ModulosLlenos)
                return;

            LogicaModulosCliente.ModulosEquipo.Clear();
            // Iterar a través de cada fila del DataGridView

            foreach (DataGridViewRow fila in dtgModulos.Rows)
            {
                // Asegurarse de que la fila no sea la fila de encabezado
                if (!fila.IsNewRow)
                {
                    //Quiere decir que al menos tenemos un registro
                    ModulosLlenos = true;
                    
                    //string Clave = fila.Cells[4].Value.ToString();
                    //LogicaModulosCliente.Claves.Add(Clave);

                    //Asignamos el numero de folio que tiene la fila
                    string FolioTabla = fila.Cells[5].Value.ToString();
                    //Preguntamos si el folio es diferente al que estamos capturando
                    if (txtNumeroFolio.Text != FolioTabla)
                    {
                        //En dado caso de que sea diferente eso quiere decir que ese modulo ya estaba por lotanto debe de actualizarse
                        ModuloEquipo NuevoModulo = new ModuloEquipo()
                        {
                            Modulo = fila.Cells[1].Value.ToString(),
                            Clave = fila.Cells[2].Value.ToString()
                        };
                        LogicaModulosCliente.ModulosEquipo.Add(NuevoModulo);
                    }
                }
            }
        }
        #endregion

        private void grpDatos_Enter(object sender, EventArgs e)
        {

        }
        public void Limpiar()
        {
            bool FormularioLleno = ValidarCamposVaciosServicios();
            erServicios.Clear();
            if (FormularioLleno)
            {
                //Por lo tanto limpiamos lo que teniamos seleccionado
                tabControl1.Visible = false;
                erServicios.Clear();
                LimpiarForm(grpDatos);
                cboClientes.SelectedItem = 0;
                cboNumeroSerie.SelectedItem = 0;
            }
            
        }
        private void btnReportes_Click(object sender, EventArgs e)
        {
            ReporteServiciosRicoh forma = new ReporteServiciosRicoh();
            forma.Show();
        }

        private void grpDatosModulo_Enter(object sender, EventArgs e)
        {

        }
    }
}
