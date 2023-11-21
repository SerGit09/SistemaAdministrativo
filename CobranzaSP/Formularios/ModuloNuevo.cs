using CobranzaSP.Lógica;
using CobranzaSP.Modelos;
using DocumentFormat.OpenXml.Office2010.Excel;
using DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CobranzaSP.Formularios
{
    public partial class ModuloNuevo : Form
    {
        private Servicios_Ricoh servicioRicoh;
        FuncionesFormularios funcion = new FuncionesFormularios();
        AccionesLógica NuevaAccion = new AccionesLógica();
        LogicaModulosCliente lgModuloCliente = new LogicaModulosCliente();
        Modulo_Cliente NuevoModulo = new Modulo_Cliente();

        bool EstaModificando;
        bool ModuloSeleccionado;
        bool ModulosLlenos = true;
        int Id = 0;
        string NumeroReporte = "";
        string Modelo = "";
        int Contador = 0;
        bool CambioModulo = false;
        bool BuscandoFolio = false;
        string Estado = "";

        private string Serie;
        //Variables cuando se cambie el modulo en algun reporte
        private string ClaveAnterior;

        #region Constructores
        //Constructor al querer agregar nuevo modulo
        public ModuloNuevo(Servicios_Ricoh servicioRicoh,string Modelo,string Serie, string Folio,  int Contador,bool EstaModificando)
        {
            InitializeComponent();
            //Asignamos la instancia que ya tenemos de los servicios ricoh, esto para poder manipular funciones en tiempo real
            this.servicioRicoh = servicioRicoh;

            this.Modelo = DeterminarModeloModulo(Modelo);
            this.Contador = Contador;
            this.Serie = Serie;
            this.EstaModificando = EstaModificando;
            this.NumeroReporte = Folio;
            InicioAplicacion();
            txtClaveRetirada.Enabled = EstaModificando;
            this.StartPosition = FormStartPosition.CenterScreen;
        }

        //Constructor al querer modificar modulo de un cliente
        public ModuloNuevo(Servicios_Ricoh servicioRicoh, string Modelo, Modulo_Cliente NuevoModulo,string Serie, string Folio,int Contador, bool EstaModificando, bool RetirarModulo, bool BuscandoFolio)
        {
            InitializeComponent();
            //Asignamos la instancia que ya tenemos de los servicios ricoh, esto para poder manipular funciones en tiempo real
            this.servicioRicoh = servicioRicoh;

            //Agregamos los parametros a nuestras variables que no seran de utilidad
            this.Serie = Serie;
            this.EstaModificando = EstaModificando;
            this.NumeroReporte = Folio;
            this.NuevoModulo = NuevoModulo;
            this.Contador = Contador;
            this.Estado = NuevoModulo.Estado;
            this.Modelo = DeterminarModeloModulo(Modelo);
            InicioAplicacion();
            this.StartPosition = FormStartPosition.CenterScreen;
            LlenarDatosModulo();

            //Habrá un cambio de modulo
            CambioModulo = RetirarModulo;
            ModuloSeleccionado = true;
            VerificarRetiroModulo(RetirarModulo);
            txtClaveRetirada.Enabled = false;

            //Deshabilitamos cambiar de modelo y modulo
            //cboModelo.Enabled = false;
            cboModulos.Enabled = false;
            this.BuscandoFolio = BuscandoFolio;
            //chkCambioClave.Visible = true;
            //cboModelo.SelectedItem= this.Modelo;
        }

        //Constructor para cuando se quiera instalar y retirar algun modulo sin tener que seleccionar previamente uno
        public ModuloNuevo(Servicios_Ricoh servicioRicoh, string Modelo, string Serie, string Folio, int Contador, bool EstaModificando, bool RetirarModulo)
        {
            InitializeComponent();
            //Asignamos la instancia que ya tenemos de los servicios ricoh, esto para poder manipular funciones en tiempo real
            this.servicioRicoh = servicioRicoh;

            this.Modelo = DeterminarModeloModulo(Modelo);
            this.Serie = Serie;
            this.EstaModificando = EstaModificando;
            this.NumeroReporte = Folio;
            this.Contador = Contador;
            InicioAplicacion();
            this.StartPosition = FormStartPosition.CenterScreen;
            //LlenarDatosModulo();
            CambioModulo = RetirarModulo;
            VerificarRetiroModulo(RetirarModulo);
            txtClaveRetirada.Enabled = !EstaModificando;
            ModuloSeleccionado = EstaModificando;

            //Habilitamos para poder seleccionar el tipo de modulo
            cboModulos.Enabled = true;

        }

        public void VerificarRetiroModulo(bool RetiroDeModulo)
        {
            if (RetiroDeModulo)
            {
                MostrarControlesCambioClave(true);
                if (ModuloSeleccionado)
                {
                    //En base al modulo que hayamos seleccionado, obtendremos las claves disponibles en bodega
                    NuevoModulo.IdModelo = NuevaAccion.BuscarId(this.Modelo, "ObtenerIdModeloModulo");
                    int IdModulo = lgModuloCliente.BuscarIdModulo(cboModulos.SelectedItem.ToString(), NuevoModulo.IdModelo);
                    //int IdModulo = NuevaAccion.BuscarId(cboModulos.SelectedItem.ToString(), "ObtenerIdModuloCliente");
                    funcion.LlenarComboBox(cboClaves, "SeleccionarModuloBodega", IdModulo);
                }
                

                //Si estamos retirando modulos quiere decir que no estamos modificando, estamos por agregar 2 nuevos registros a nuestra tabla
                EstaModificando = false;
            }
        }

        public string DeterminarModeloModulo(string Modelo)
        {
            var prefijosValidos = new HashSet<string> { "MP-400", "MP-500","90"};
            //Para saber si se trata de un fusor
            if (prefijosValidos.Any(prefijo => Modelo.StartsWith(prefijo)))
            {
                return "4002";
            }

            return "5054";
        }

        public void LlenarDatosModulo()
        {
            Id = NuevoModulo.Id;
            //cboModelo.SelectedItem = NuevoModulo.Modelo;
            cboModulos.SelectedItem = NuevoModulo.Modulo;
            txtClave.Text = NuevoModulo.Clave;
            //txtPaginas.Text = NuevoModulo.Paginas.ToString();
            ClaveAnterior = NuevoModulo.Clave;
            rtxtObsrevacion.Text = NuevoModulo.Observacion;
        }
        #endregion

        #region Inicio
        public void InicioAplicacion() 
        {
            //funcion.LlenarComboBox(cboModelo, "SeleccionarModelosInventarioModulos", 0);

            //txtPaginas.Enabled = (EstaModificando) ? false:true;
            

            int IdModelo = NuevaAccion.BuscarId(Modelo, "ObtenerIdModeloModulo");
            funcion.LlenarComboBox(cboModulos, "SeleccionarModulos", IdModelo);

            //cboModelo.Enabled = false;
        }
        #endregion

        #region PanelSuperior
        [DllImport("user32.DLL", EntryPoint = "ReleaseCapture")]
        private extern static void ReleaseCapture();
        [DllImport("user32.DLL", EntryPoint = "SendMessage")]

        private extern static void SendMessage(System.IntPtr hWnd, int wMsg, int wParam, int lparam);


        private void panelSuperior_MouseDown_1(object sender, MouseEventArgs e)
        {
            ReleaseCapture();
            SendMessage(this.Handle, 0x112, 0xf012, 0);
        }
        #endregion

        #region Validaciones
        public bool ValidarCamposVaciosModulos()
        {
            bool validado = true;
            erModulo.Clear();
            foreach (Control c in grpModulo.Controls)
            {
                if (c is TextBox textBox && String.IsNullOrWhiteSpace(textBox.Text))
                {
                    if(c.Name == "txtClaveRetirada" || c.Name == "rtxtObservacion")
                    {
                        if (CambioModulo)
                        {
                            erModulo.SetError(c, "CampoObligatorio");
                            validado = false;
                        }
                    }
                    else
                    {
                        erModulo.SetError(c, "Campo obligatorio");
                        validado = false;
                    }
                }
                if (c is ComboBox combo && combo.SelectedIndex == 0)
                {
                    if (combo.Name == "cboClaves")
                    {
                        if (CambioModulo)
                        {
                            erModulo.SetError(c, "Seleccione una clave");
                            validado = false;
                        }
                    }
                    else
                    {
                        erModulo.SetError(c, "Campo obligatorio");
                        validado = false;
                    }
                }
            }

            return validado;
        }
        #endregion
        #region Botones
        private void btnCerrar_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        #region AgregarModificar
        private void btnAgregarModulo_Click(object sender, EventArgs e)
        {
            if (!ValidarCamposVaciosModulos())
                return;
            
            try
            {
                string Resultado;
                Modulo_Cliente NuevoModulo = new Modulo_Cliente()
                {
                    Id = Id,
                    IdModelo = NuevaAccion.BuscarId(this.Modelo, "ObtenerIdModeloModulo"),
                    Reporte = NumeroReporte,
                    Serie = Serie,
                    Paginas = Contador,
                    Clave = txtClave.Text,
                    ClaveAnterior = ClaveAnterior,
                    Observacion = rtxtObsrevacion.Text
                };
                ColocarClaveModulo(NuevoModulo);
                ColocarClaveAnterior(NuevoModulo);
                NuevoModulo.IdModulo = lgModuloCliente.BuscarIdModulo(cboModulos.SelectedItem.ToString(), NuevoModulo.IdModelo);
                if (EstaModificando)
                {
                    if (MessageBox.Show("¿Esta seguro de modificar el registro?", "CONFIRME LA MODIFICACION", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                    {
                        MessageBox.Show("Modificacion cancelada!!", "CANCELADO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }
                    NuevoModulo.Estado = Estado;

                    //CHECAR
                    if(NuevoModulo.Observacion != "" && BuscandoFolio == false)
                    {
                        //Por lo que deberemos de agregar un nuevo registro a nuestra tabla, en este caso una actualzacion, ya que se añadio una observacion en dicho reporte que se esta agregando
                        NuevoModulo.Estado = "ACTUALIZADO";
                        NuevoModulo.Id = 0;
                        lgModuloCliente.RegistrarModulo(NuevoModulo, "AgregarEstadoModuloEquipo");
                        EliminarClaveDeLista(NuevoModulo.Clave);
                        LogicaModulosCliente.ClavesObservaciones.Add(NuevoModulo.Clave);
                    }
                    else//Eso quiere decir que se añadio una nueva observacion
                    {
                        Resultado = lgModuloCliente.RegistrarModulo(NuevoModulo, "ModificarEstadoModuloEquipo");
                        MessageBox.Show(Resultado, "REGISTRO INVENTARIO MODULOS", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }

                    //ASI ES COMO ESTABA ANTES
                    //Resultado = lgModuloCliente.RegistrarModulo(NuevoModulo, "ModificarEstadoModuloEquipo");
                    //MessageBox.Show(Resultado, "REGISTRO INVENTARIO MODULOS", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    servicioRicoh.MostrarModulosCliente(Serie,false);
                }
                else//Agregamos al catalogo
                {
                    //Comprobamos si no hubo retiro de modulo
                    if (ComprobarCambioModulo(NuevoModulo))
                    {
                        //Agregamos el modulo que se esta retirando
                        ModificarHistorialModulos(NuevoModulo);
                        NuevoModulo.Id = 0;
                    }
                    else
                    {
                        //verificamos que no tenga un modulo ya colocado
                        //bool ModuloAsignado = lgModuloCliente.VerificarModuloAsignado(NuevoModulo.Serie, NuevoModulo.IdModulo);
                        //if (ModuloAsignado)
                        //{
                        //    MessageBox.Show("¡¡ESTE EQUIPO YA CUENTA CON UN MODULO ASIGNADO!!", "MODULO ASIGNADO", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        //    return;
                        //}

                        ////Validamos que dicha clave que asiganamos no este en otro equipo
                        bool ClaveAsignada = NuevaAccion.VerificarDuplicados(NuevoModulo.Clave, "ValidarClaveModulo");
                        if (ClaveAsignada)
                        {
                            MessageBox.Show("CLAVE ASIGNADA EN EL EQUIPO " + lgModuloCliente.ObtenerSerieClaveRepetida(NuevoModulo.Clave), "!CLAVE YA EXISTE EN OTRO EQUIPO!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }
                    }

                    NuevoModulo.Estado = "INSTALADO";
                    //Agregamos la clave a la lista
                    LogicaModulosCliente.ClavesObservaciones.Add(NuevoModulo.Clave);
                    Resultado = lgModuloCliente.RegistrarModulo(NuevoModulo, "AgregarEstadoModuloEquipo");
                    MessageBox.Show(Resultado, "REGISTRO MODULO EN EQUIPO DE CLIENTE", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    if(CambioModulo)
                    {
                        servicioRicoh.MostrarModulosCliente(Serie, false);
                    }
                    else
                    {
                        servicioRicoh.MostrarModulosCliente(Serie, true);
                    }
                }
                //Actualizamos en el formulario de servicios ricoh
                
                servicioRicoh.ReiniciarCapturaModulo();
                //Reiniciar();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ocurrio un error " + ex.Message);
            }
            //MessageBox.Show(Serie);
            this.Close();
        }

        //Metodo que nos permitira saber cuando se retire o actualice un modulo, o en dado cado de que se instale uno
        public void ColocarClaveModulo(Modulo_Cliente NuevoModulo)
        {
            if (ModuloSeleccionado)
            {
                if (EstaModificando)
                {
                    NuevoModulo.Clave = txtClave.Text;
                    return;
                }
                NuevoModulo.Clave = cboClaves.SelectedItem.ToString();
            }
            else
            {
                NuevoModulo.Clave = txtClave.Text;
            }
        }

        public void ColocarClaveAnterior(Modulo_Cliente NuevoModulo)
        {
            if (!ModuloSeleccionado)
            {
                NuevoModulo.ClaveAnterior = txtClaveRetirada.Text;
                ClaveAnterior = txtClaveRetirada.Text;
            }
        }

        
        #endregion
        #endregion

        #region RetiroDeModulo
        public void MostrarControlesCambioClave(bool mostrar)
        {
            cboClaves.Visible = ModuloSeleccionado;
            txtClave.Visible = !ModuloSeleccionado;
            txtClaveRetirada.Visible = mostrar;
            lblAnteriorClave.Visible = mostrar;
            CambioModulo = mostrar;
            txtClaveRetirada.Text = NuevoModulo.Clave;
        }

        public void ModificarHistorialModulos(Modulo_Cliente NuevoModulo)
        {
            if (CambioModulo)
            {
                //Guardamos en el historial el modulo que se retiro
                lgModuloCliente.RetirarModulo(NuevoModulo, ModuloSeleccionado);
                //Deberemos de agregar el nuevo modulo al hisotrial como instalado
                //lgModuloCliente.InstalarModulo(NuevoModulo);
                NuevoModulo.Estado = "INSTALADO";
            }
        }

        public bool ComprobarCambioModulo(Modulo_Cliente NuevoModulo)
        {
            //Posiblemente algunas cosas esten de más
            if (ClaveAnterior != NuevoModulo.Clave && CambioModulo)
            {
                MessageBox.Show("Se instalo " + NuevoModulo.Clave + " y se retiro " + ClaveAnterior);
                NuevoModulo.ClaveAnterior = ClaveAnterior;
                EliminarClaveDeLista(ClaveAnterior);
                CambioModulo = true;
                return true;
            }
            else
            {
                return false;
            }
        }

        public void EliminarClaveDeLista(string Clave)
        {
            ModuloEquipo ModuloAEliminar = LogicaModulosCliente.ModulosEquipo.FirstOrDefault(Modulo => Modulo.Clave == Clave);

            if (ModuloAEliminar != null)
            {
                LogicaModulosCliente.ModulosEquipo.Remove(ModuloAEliminar);
            }
        }
        #endregion



        #region Eventos
        private void cboModulos_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void ModuloNuevo_Load(object sender, EventArgs e)
        {
            //this.Location = new Point(300, 300);

        }

        
        #endregion

    }
}
