using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data;
using System.Data.SqlClient;
using CobranzaSP.Lógica;
using CobranzaSP.Modelos;

namespace CobranzaSP.Formularios
{
    public partial class Inventario : Form
    {
        public Inventario()
        {
            InitializeComponent();
            InicioAplicacion();
        }
        AccionesLógica NuevaAccion = new AccionesLógica();
        CD_Conexion cn = new CD_Conexion();
        LogicaInventario AccionInventario = new LogicaInventario();
        LogicaRegistro AccionRegistro = new LogicaRegistro();
        //Objeto que nos servirá para poder tener capturado en dado caso de una modificacion
        RegistroInventario RegistroAnterior;
        bool inventario = true;
        bool EstaModificando = false;
        int Id = 0;
        bool BuscandoRegistro = false;

        private void tabRegistros_Click(object sender, EventArgs e)
        {

        }

        #region Inicio
        public void InicioAplicacion()
        {
            PropiedadesDtg();

            //Metodo que controla los controles que seran utiles al iniciar la aplicacion
            ControlesDesactivados(false);

            //Agregamos opciones que estan en la base de datos
            LlenarComboBox(cboClientes, "SeleccionarClientesServicios", 0);
            
            //LlenarComboBox(cboModelos, "SeleccionarCartuchos", 0);
            LlenarComboBox(cboMarca, "SeleccionarMarca", 0);
            LlenarComboBox(cboMarcas, "SeleccionarMarca", 0);
            LlenarComboBox(cboModelos, "SeleccionarCartuchos", 0);
            LlenarComboBox(cboFusor, "SeleccionarFusores", 0);

            //Denegar escritura en combobox
            cboModelos.DropDownStyle = ComboBoxStyle.DropDownList;
            cboClientes.DropDownStyle = ComboBoxStyle.DropDownList;
            cboMarcaSeleccionada.DropDownStyle = ComboBoxStyle.DropDownList;

            Mostrar("MostrarInventario");

            //Configuramos los DateTimePicker
            //dtpFechaInicio.MaxDate = DateTime.Today;
            //dtpFechaFinal.MaxDate = DateTime.Today;
            dtpFecha.MaxDate = DateTime.Now;
            dtpFechaRegistro.MaxDate = DateTime.Now;
            LlenarCboMarcaSeleccionada();
            //LlenarCboFechaRegistros();
            radEntrada.Checked = true;
        }

        public void LlenarCboMarcaSeleccionada()
        {
            string[] Marcas = { "", "Todos", "Hp", "Brother" ,"Canon","Xerox", "Ricoh", "Samsung"};
            cboMarcaSeleccionada.Items.AddRange(Marcas);
        }
        public void ControlesDesactivados(bool Desactivado)
        {
            btnEliminar.Enabled = Desactivado;
            btnCancelar.Enabled = Desactivado;
        }

        public void PropiedadesDtg()
        {
            //Solo lectura
            dtgCartuchos.ReadOnly = true;

            //No agregar renglones
            dtgCartuchos.AllowUserToAddRows = false;

            //No borrar renglones
            dtgCartuchos.AllowUserToDeleteRows = false;

            //Ajustar automaticamente el ancho de las columnas
            dtgCartuchos.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            //dtgServicios.AutoResizeColumns(DataGridViewAutoSizeColumnsMo‌​de.Fill);
            dtgCartuchos.AutoResizeColumns();

            dtgCartuchos.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
        }

        public void Mostrar(string sp)
        {
            //Limpiamos los datos del datagridview
            dtgCartuchos.DataSource = null;
            dtgCartuchos.Refresh();
            DataTable tabla = new DataTable();
            //Guardamos los registros dependiendo la consulta
            tabla = NuevaAccion.Mostrar(sp);
            //Asignamos los registros que optuvimos al datagridview
            dtgCartuchos.DataSource = tabla;
            if (inventario)
            {
                dtgCartuchos.Columns["Id"].Visible = false;
            }
            else
            {
                dtgCartuchos.Columns["IdRegistro"].Visible = false;
            }

        }

        public void LlenarComboBox(ComboBox cb, string sp, int indice)
        {
            SqlDataReader dr;
            cb.Items.Clear();
            if (sp == "SeleccionarModelos" || sp == "SeleccionarCartuchos")
            {
                dr = NuevaAccion.LlenarComboBoxModelos(sp, indice);
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
        private bool ValidarCamposInventario()
        {
            bool Validado = true;
            erInventario.Clear();

            if (string.IsNullOrEmpty(txtModelo.Text))
            {
                erInventario.SetError(txtModelo, "Ingrese el modelo");
                Validado = false;
            }
            if (string.IsNullOrEmpty(txtOficina.Text))
            {
                erInventario.SetError(txtOficina, "Ingrese cantidad en oficina");
                Validado = false;
            }
            if (cboMarca.SelectedIndex == 0)
            {
                erInventario.SetError(cboMarca, "Seleccione una marca");
                Validado = false;
            }
            return Validado;
        }

        private bool ValidarRegistros()
        {
            bool Validado = true;
            erInventario.Clear();

            if (string.IsNullOrEmpty(txtCantidad.Text))
            {
                erInventario.SetError(txtCantidad, "Ingrese una cantidad");
                Validado = false;
            }
            if (cboModelos.Text == " ")
            {
                erInventario.SetError(cboModelos, "Seleccione un modelo");
                Validado = false;
            }
            if (cboMarcas.Text == " ")
            {
                erInventario.SetError(cboMarcas, "Seleccione una marca");
                Validado = false;
            }
            if (cboClientes.SelectedIndex == 0)
            {
                erInventario.SetError(cboClientes, "Seleccione cliente");
                Validado = false;
            }
            return Validado;
        }

        private void txtModelo_KeyPress(object sender, KeyPressEventArgs e)
        {
            Validacion.SoloLetrasYNumeros(e);
        }

        private void txtOficina_KeyPress(object sender, KeyPressEventArgs e)
        {
            Validacion.SoloNumeros(e);
        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            Validacion.SoloLetrasYNumeros(e);
        }

        private void txtCantidad_KeyPress(object sender, KeyPressEventArgs e)
        {
            Validacion.SoloNumeros(e);
        }

        #endregion

        #region Botones

        public void SeccionInventario()
        {
            if (!ValidarCamposInventario())
                return;
            InventarioDatos nuevoCartucho = new InventarioDatos()
            {
                Id = Id,
                Modelo = txtModelo.Text,
                IdMarca = NuevaAccion.BuscarId(cboMarca.SelectedItem.ToString(), "ObtenerIdMarca"),
                CantidadOficina = int.Parse(txtOficina.Text),
                Fecha = dtpFecha.Value,
                Precio = double.Parse(txtPrecio.Text)
            };
            if (EstaModificando)
            {
                //Modificar stop procedure de modificar, para que se pueda cambiar la fecha
                if (MessageBox.Show("¿Esta seguro de modificar el registro?", "CONFIRME LA MODIFICACION", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                {
                    MessageBox.Show("!!Modificación cancelada!!", "CANCELADO", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                    LimpiarForm(grpDatosInventario);
                    return;
                }
                AccionInventario.RegistrarInventario(nuevoCartucho, "ModificarInventario");
                MessageBox.Show("Registro modificado correctamente", "MODIFICACION DE REGISTRO", MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
            else
            {
                bool ModeloDuplicado = NuevaAccion.VerificarDuplicados(nuevoCartucho.Modelo, "VerificarModeloExistente");
                if (ModeloDuplicado)
                {
                    MessageBox.Show("Ingrese un modelo distinto", "EL MODELO YA EXISTE", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                else
                {
                    AccionInventario.RegistrarInventario(nuevoCartucho, "AñadirInventario");
                    MessageBox.Show("Toner agregado al inventario", "REGISTRO EXITOSO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }

            }

            //ESTA LA USO EN MI BASE DE DATOS
            Mostrar("MostrarInventario");

            LimpiarForm(grpDatosInventario);

        }

        private void btnGuardar_Click(object sender, EventArgs e)
        {
            try
            {
                //Entrara en dado caso que nos encontremos en el inventario
                if (inventario)
                {
                    SeccionInventario();
                }
                else//Estamos en los registros
                {
                    SeccionRegistro();
                    //SeccionRegistroPrueba();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ocurrio un error" + ex);
            }
        }

        public void SeccionRegistro()
        {
            if (!ValidarRegistros())
                return;

            int IdMarcaEncontrada = NuevaAccion.BuscarId(cboMarcas.SelectedItem.ToString(), "ObtenerIdMarca");
            RegistroInventario nuevoRegistro = new RegistroInventario()
            {
                Id = Id,
                Cliente = cboClientes.SelectedItem.ToString(),
                IdMarca = IdMarcaEncontrada,
                IdCartucho = NuevaAccion.BuscarIdCartucho(cboModelos.SelectedItem.ToString(), IdMarcaEncontrada, "ObtenerIdCartucho"),
                Fecha = dtpFechaRegistro.Value
            };
            ColocarCantidadesMovimiento(nuevoRegistro);

            //Verificamos que la clave del fusor este en los fusores o si esta vacio
            if (ComprobarNumeroSerie())
            {
                nuevoRegistro.NumeroSerie = cboFusor.SelectedItem.ToString();
            }


            if (EstaModificando)
            {
                if (MessageBox.Show("¿Esta seguro de modificar el registro?", "CONFIRME LA MODIFICACION", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                {
                    MessageBox.Show("!!Modificación cancelada!!", "CANCELADO", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                    LimpiarForm(grpDatosRegistro);
                    return;
                }

                //En dado caso de que sean diferentes quiere decir que cambio el modelo o la marca, por lo que se tiene que modificar
                //tambien ese registro anterior que se tenia
                bool CantidadesModificadas = RegistroAnterior.IdMarca != nuevoRegistro.IdMarca || RegistroAnterior.CantidadSalida != nuevoRegistro.CantidadSalida || RegistroAnterior.CantidadEntrada != nuevoRegistro.CantidadEntrada || RegistroAnterior.IdCartucho != nuevoRegistro.IdCartucho;
                if (CantidadesModificadas)
                {
                    if (nuevoRegistro.CantidadSalida > 0)//Comprobamos si se trata de una salida
                    {
                        if (AccionRegistro.ModificarRegistroInventario(nuevoRegistro, false))//Comprobamos si podemos realizar los cambios
                        {
                            //Ahora deberemos de modificar el registro anterior
                            AccionRegistro.ModificarInventario(RegistroAnterior, "ModificarTonerInventario");
                            MessageBox.Show("Se ha actualizado el registro", "MODIFICACION EXITOSA", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        else
                        {
                            MessageBox.Show("La cantidad de salida excede la cantidad disponible en oficina", "FALLO DE MODIFICACION", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                    else//Si no, se trata de una entrada
                    {
                        //Comprobamos que podamos quitarle las entradas al cartucho anterior y si podemos hacer nuevo registro
                        if (AccionRegistro.ComprobarCantidadInventario(RegistroAnterior) && AccionRegistro.ModificarRegistroInventario(nuevoRegistro, false))
                        {
                            AccionRegistro.ModificarInventario(RegistroAnterior, "ModificarTonerInventario");
                        }
                        else
                        {
                            MessageBox.Show("La cantidad de entrada excede la cantidad disponible en oficina del registro anterior", "FALLO DE MODIFICACION", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                }
                else
                {
                    //Si no hay modificaciones acerca de cantidades, marcas o modelos, solamente se modificara el registro en sí
                    if (AccionRegistro.ModificarRegistroInventario(nuevoRegistro, true))
                    {
                        MessageBox.Show("Registro modificado correctamente", "MODIFICACION EXITOSA", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        MessageBox.Show("Algo salió mal al realizar la modificación", "ERROR EN MODIFICACION", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }

            }
            else
            {
                string Mensaje = AccionRegistro.AgregarRegistroInventario(nuevoRegistro);
                MessageBox.Show(Mensaje, "REGISTRO INVENTARIO", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

            Mostrar("VerRegistroInventario");
            LimpiarForm(grpDatosRegistro);
        }

        public void ColocarCantidadesMovimiento(RegistroInventario nuevoRegistro)
        {
            nuevoRegistro.CantidadEntrada = 0;
            nuevoRegistro.CantidadSalida = 0;
            nuevoRegistro.CantidadGarantia = 0;
            if (radEntrada.Checked)
            {
                nuevoRegistro.CantidadEntrada = int.Parse(txtCantidad.Text);
            }
            else if (radSalida.Checked)
            {
                nuevoRegistro.CantidadSalida = int.Parse(txtCantidad.Text);
            }
            else
            {
                nuevoRegistro.CantidadGarantia = int.Parse(txtCantidad.Text);
            }
        }

        private void btnEliminar_Click(object sender, EventArgs e)
        {
            try
            {
                int IdMarca = NuevaAccion.BuscarId(cboMarcas.SelectedItem.ToString(), "ObtenerIdMarca");
                //int IdMarca = NuevaAccion.BuscarId(cboMarcas.SelectedItem.ToString(), "ObtenerIdMarca");
                if (MessageBox.Show("¿Esta seguro de eliminar el registro?", "CONFIRMAR ELIMINACION", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                {
                    MessageBox.Show("!!Eliminación cancelada!!", "CANCELADO", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                    LimpiarForm(grpDatosRegistro);
                    return;
                }
                if (inventario)
                {
                    NuevaAccion.Eliminar(Id, "EliminarCartuchoInventario");
                    MessageBox.Show("Registro eliminado correctamente", "ELIMINACION EXITOSA", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    Mostrar("MostrarInventario");
                    LimpiarForm(grpDatosInventario);
                }
                else
                {
                    RegistroInventario eliminarRegistro = new RegistroInventario()
                    {
                        Id = Id,
                        IdMarca = IdMarca,
                        IdCartucho = NuevaAccion.BuscarIdCartucho(cboModelos.SelectedItem.ToString(), IdMarca, "ObtenerIdCartucho"),
                        //IdCartucho = NuevaAccion.BuscarIdCartucho(cboModelosP.SelectedItem.ToString(), IdMarca, "ObtenerIdCartucho"),
                        //CantidadSalida = int.Parse(txtCantidadSalida.Text),
                        //CantidadEntrada = int.Parse(txtCantidadEntrada.Text),
                        //CantidadGarantia = int.Parse(txtCantidad.Text),
                        NumeroSerie = cboFusor.SelectedItem.ToString()
                    };
                    ColocarCantidadesMovimiento(eliminarRegistro);
                    //Primero verifiquemos que la cantidad a restar en el anterior registro no exceda a las que cuenta
                    if (AccionRegistro.ComprobarCantidadInventario(eliminarRegistro))
                    {
                        AccionRegistro.ModificarInventario(eliminarRegistro, "EliminarRegistroInventario");
                    }
                    else
                    {
                        MessageBox.Show("La cantidad que se quiere restar de entrada excede la cantidad del toner", "ERROR AL ELIMINAR", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        LimpiarForm(grpDatosRegistro);
                        return;
                    }

                    MessageBox.Show("Registro eliminado correctamente", "ELIMINAR REGISTRO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    Mostrar("VerRegistroInventario");
                    ReiniciarControles();
                }
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
        
        //Evento que nos ayudara a mostrar en el combobox la informacion del inventario o de los registros dependiendo cual seleccionemos
        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            ReiniciarControles();
            if (tabControl1.SelectedTab == tabInventario)
            {
                inventario = true;
                Mostrar("MostrarInventario");
                
            }
            else
            {
                inventario = false;
                Mostrar("VerRegistroInventario");
            }
        }

        public void SeleccionarDatosRegistros()
        {
            if (inventario)
            {
                LlenarCamposInventario();
            }
            else
            {
                LlenarCamposRegistroInventario();
            }
        }

        public void LlenarCamposRegistroInventario()
        {
            int IdMarcaSeleccionada = NuevaAccion.BuscarId(dtgCartuchos.CurrentRow.Cells[1].Value.ToString(), "ObtenerIdMarca");
            Id = int.Parse(dtgCartuchos.CurrentRow.Cells[0].Value.ToString());

            cboMarcas.SelectedItem = dtgCartuchos.CurrentRow.Cells[1].Value.ToString();
            cboModelos.SelectedItem = dtgCartuchos.CurrentRow.Cells[2].Value.ToString();
            //Aqui se requiere capturar de nuevo el modelo y la marca en dado caso de que se vaya cambiar el modelo o la marca
            //Esto para tener capturado la marca y el modelo anterior para realizar la respectiva modificacion en el inventario
            RegistroAnterior = new RegistroInventario()
            {
                IdMarca = IdMarcaSeleccionada,
                IdCartucho = NuevaAccion.BuscarIdCartucho(dtgCartuchos.CurrentRow.Cells[2].Value.ToString(), IdMarcaSeleccionada, "ObtenerIdCartucho"),
                CantidadSalida = int.Parse(dtgCartuchos.CurrentRow.Cells[3].Value.ToString()),
                CantidadEntrada = int.Parse(dtgCartuchos.CurrentRow.Cells[4].Value.ToString()),
                CantidadGarantia = int.Parse(dtgCartuchos.CurrentRow.Cells[5].Value.ToString())
            };
            SeleccionarRadioButtonTipoMovimiento(RegistroAnterior);


            cboClientes.SelectedItem = dtgCartuchos.CurrentRow.Cells[6].Value.ToString();
            dtpFechaRegistro.Value = DateTime.Parse(dtgCartuchos.CurrentRow.Cells[7].Value.ToString());

            //Comprobamos que tenga un numero de serie
            if (dtgCartuchos.CurrentRow.Cells[8].Value.ToString() != " ")
            {
                MostrarCapturaSerie(true);
                cboFusor.SelectedItem = dtgCartuchos.CurrentRow.Cells[8].Value.ToString();
                RegistroAnterior.NumeroSerie = dtgCartuchos.CurrentRow.Cells[8].Value.ToString();
            }
        }

        public void SeleccionarRadioButtonTipoMovimiento(RegistroInventario RegistroSeleccionado)
        {
            if (RegistroSeleccionado.CantidadSalida > 0)
            {
                radSalida.Checked = true;
                txtCantidad.Text = RegistroSeleccionado.CantidadSalida.ToString();
            }
            else if (RegistroSeleccionado.CantidadEntrada > 0)
            {
                radEntrada.Checked = true;
                txtCantidad.Text = RegistroSeleccionado.CantidadEntrada.ToString();
            }
            else
            {
                radGarantia.Checked = true;
                txtCantidad.Text = RegistroSeleccionado.CantidadGarantia.ToString();
            }
        }

        public void LlenarCamposInventario()
        {
            Id = int.Parse(dtgCartuchos.CurrentRow.Cells[0].Value.ToString());
            cboMarca.SelectedItem = dtgCartuchos.CurrentRow.Cells[1].Value.ToString();
            txtModelo.Text = dtgCartuchos.CurrentRow.Cells[2].Value.ToString();
            txtOficina.Text = dtgCartuchos.CurrentRow.Cells[3].Value.ToString();
            dtpFecha.Value = DateTime.Parse(dtgCartuchos.CurrentRow.Cells[4].Value.ToString());
            txtPrecio.Text = dtgCartuchos.CurrentRow.Cells[5].Value.ToString().Replace("$", "").Replace(",", "");
        }

        public void MostrarCapturaSerie(bool Mostrar)
        {
            lblSerieP.Visible = Mostrar;
            cboFusor.Visible = Mostrar;
        }

        private void dtgCartuchos_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            EstaModificando = true;
            ControlesDesactivados(true);
            SeleccionarDatosRegistros();
        }

        private void cboMarcaP_SelectedIndexChanged(object sender, EventArgs e)
        {
            //En dado caso que se haya seleccionado algo de las marcas y mientras no estemos buscando un registro en especifico
            if (cboMarcas.SelectedItem.ToString() != " " && BuscandoRegistro == false)
            {
                int IdMarca = NuevaAccion.BuscarId(cboMarcas.SelectedItem.ToString(), "ObtenerIdMarca");
                LlenarComboBox(cboModelos, "SeleccionarCartuchos", IdMarca);
            }
        }

        private void cboModelosP_SelectedIndexChanged(object sender, EventArgs e)
        {
            string Cartucho = cboModelos.SelectedItem.ToString();
            var prefijosValidos = new HashSet<string> { "RM1-", "D01SE", "DR", "RM2-" };
            //Para saber si se trata de un fusor
            if (prefijosValidos.Any(prefijo => Cartucho.StartsWith(prefijo)))
            {
                MostrarCapturaSerie(true);
                //Si entra en la condicion quiere decir que se trata e un drum
                if (Cartucho.StartsWith("DR"))
                {
                    LlenarComboBox(cboFusor, "SeleccionarSeriesBrother", 0);
                    return;
                }
                LlenarComboBox(cboFusor, "SeleccionarFusores", 0);
            }
            else
            {
                MostrarCapturaSerie(false);
            }
        }

        #endregion

        #region MetodosLocales
        private void AbrirForm(object formNuevo)
        {
            //Declaramos la forma
            Form fh = formNuevo as Form;

            //Mostramos la forma 
            fh.Show();
        }

        private void LimpiarForm(GroupBox grp)
        {
            foreach (Control c in grp.Controls)
            {
                if (c is TextBox)
                {
                    c.Text = "";
                }
            }
            txtModelo.Focus();

            cboClientes.SelectedIndex = 0;
            cboModelos.SelectedIndex = 0;
            cboMarcas.SelectedIndex = 0;
            cboMarca.SelectedIndex = 0;
            cboMarcaSeleccionada.SelectedIndex = 0;
            cboFusor.SelectedIndex = 0;
            dtpFechaRegistro.Value = DateTime.Today;
        }

        public void ReiniciarControles()
        {
            EstaModificando = false;
            ControlesDesactivados(false);
            LimpiarForm(grpDatosRegistro);
            LimpiarForm(grpDatosInventario);
            Id = 0;
            dtpFechaRegistro.Value = DateTime.Today;
        }

        public bool ComprobarNumeroSerie()
        {
            bool DrumUtilizado = false;
            if (cboFusor.SelectedIndex > 0)
            {
                DrumUtilizado = true;
            }
            return DrumUtilizado;
        }

        

        #endregion

        #region Formas
        private void btnAbrirReportes_Click(object sender, EventArgs e)
        {
            AbrirForm(new ReporteRegistros());
        }
        private void btnAbrirReporteExistencias_Click(object sender, EventArgs e)
        {
            AbrirForm(new ReporteInventario());
        }

        private void btnImprimir_Click(object sender, EventArgs e)
        {
            AccionInventario.ImprimirInventario();
        }
        #endregion


        private void btnGarantias_Click(object sender, EventArgs e)
        {
            AbrirForm(new Garantias());
        }
    }
}
