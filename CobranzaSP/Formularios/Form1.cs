using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using CobranzaSP.Formularios;

namespace CobranzaSP
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            SubMenus();
        }

        #region SubMenus
        public void SubMenus()
        {
            panelSubMenuCobranza.Visible = false;
            SubMenuEquipos.Visible = false;
            SubMenuServicios.Visible = false;
            SubMenuInventario.Visible = false;
            SubMenuModulos.Visible = false;
        }

        public void OcultarSubMenu()
        {
            if (panelSubMenuCobranza.Visible == true)
                panelSubMenuCobranza.Visible = false;
            if (SubMenuEquipos.Visible == true)
                SubMenuEquipos.Visible = false;
            if (SubMenuServicios.Visible == true)
                SubMenuServicios.Visible = false;
            if (SubMenuInventario.Visible == true)
                SubMenuInventario.Visible = false;
        }

        public void MostrarSubMenu(Panel SubMenu)
        {
            if (SubMenu.Visible == false)
            {
                OcultarSubMenu();
                SubMenu.Visible = true;
            }
            else
                SubMenu.Visible = false;
        }
        #endregion

        #region BotonesControlInterfaz
        private void btnCerrar_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void btnRestaurar_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Normal;
            btnMaximizar.Visible = true;
            btnRestaurar.Visible = false;
        }

        private void btnMinimizar_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }

        private void btnMaximizar_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Maximized;
            btnMaximizar.Visible = false;
            btnRestaurar.Visible = true;
        }
        #endregion


        #region Paneles
        [DllImport("user32.DLL", EntryPoint = "ReleaseCapture")]
        private extern static void ReleaseCapture();
        [DllImport("user32.DLL", EntryPoint = "SendMessage")]

        private extern static void SendMessage(System.IntPtr hWnd, int wMsg, int wParam, int lparam);
        private void panelSuperior_MouseDown(object sender, MouseEventArgs e)
        {
            ReleaseCapture();
            SendMessage(this.Handle, 0x112, 0xf012, 0);
        }

        private void panelCentral_Paint(object sender, PaintEventArgs e)
        {

        }

        private void panelSuperior_Paint(object sender, PaintEventArgs e)
        {

        }
        #endregion


        private void AbrirForm(object formNuevo)
        {
            //Si el contenedor central cuenta con formulario ya en el
            if (this.panelCentral.Controls.Count > 0)
            {
                //Lo removeremos para poder agregar el que se eligio
                this.panelCentral.Controls.RemoveAt(0);
            }
            Form fh = formNuevo as Form;
            fh.TopLevel = false; //Decimos que no es un formulario de nivel superior
            fh.Dock = DockStyle.Fill;//Contenedor que estara en el centro
            this.panelCentral.Controls.Add(fh);//Añadimos la forma al panel
            this.panelCentral.Tag = fh;
            //Mostramos la interfaz que se haya seleccionado
            fh.Show();
        }

        #region AbrirInterfaces
        private void btnCobranza_Click(object sender, EventArgs e)
        {
            OcultarSubMenu();
            AbrirForm(new ICobranza());
        }

        private void btnRemisiones_Click(object sender, EventArgs e)
        {
            OcultarSubMenu();
            AbrirForm(new IRemisiones());
        }

        private void btnCobrado_Click(object sender, EventArgs e)
        {
            OcultarSubMenu();
            AbrirForm(new CuentasPagadas());
        }

        private void btnGraficas_Click(object sender, EventArgs e)
        {
            OcultarSubMenu();
            AbrirForm(new Graficas());
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            MostrarSubMenu(panelSubMenuCobranza);
        }

        private void btnServicios_Click(object sender, EventArgs e)
        {
            MostrarSubMenu(SubMenuServicios);
        }


        private void btnCompras_Click(object sender, EventArgs e)
        {
            AbrirForm(new Compras());
        }
        #endregion


        private void btnServiciosEquipos_Click(object sender, EventArgs e)
        {
            OcultarSubMenu();
            AbrirForm(new Servicios());
        }

        private void btnClientes_Click(object sender, EventArgs e)
        {
            SubMenus();
            AbrirForm(new Contactos());
        }

        private void btnInventario_Click(object sender, EventArgs e)
        {
            MostrarSubMenu(SubMenuInventario);
        }

        private void btnFusores_Click(object sender, EventArgs e)
        {
            SubMenus();
            AbrirForm(new Fusores());
        }

        private void btnEquipos_Click(object sender, EventArgs e)
        {
            MostrarSubMenu(SubMenuEquipos);
        }

        private void btnBodega_Click(object sender, EventArgs e)
        {
            OcultarSubMenu();
            AbrirForm(new EquiposBodega());
        }

        private void btnRentaEquipo_Click(object sender, EventArgs e)
        {
            OcultarSubMenu();
            AbrirForm(new Equipos());
        }

        private void btnEquiposVendidos_Click(object sender, EventArgs e)
        {
            OcultarSubMenu();
            AbrirForm(new EquiposVendidos());
        }

        private void btnProveedores_Click(object sender, EventArgs e)
        {
            SubMenus();
        }

        private void btnServiciosRicoh_Click(object sender, EventArgs e)
        {
            OcultarSubMenu();
            AbrirForm(new Servicios_Ricoh());
        }

        private void btnInventarioToners_Click(object sender, EventArgs e)
        {
            OcultarSubMenu();
            AbrirForm(new Inventario());
        }

        private void button2_Click(object sender, EventArgs e)
        {
            OcultarSubMenu();
            AbrirForm(new InventarioPartesRicoh());
        }

        private void btnModulos_Click(object sender, EventArgs e)
        {
            OcultarSubMenu();
            AbrirForm(new InventarioModulos());
        }

        private void button4_Click(object sender, EventArgs e)
        {
            MostrarSubMenu(SubMenuModulos);
        }

        private void btnModuloBodega_Click(object sender, EventArgs e)
        {
            AbrirForm(new ModulosBodega());
        }
    }
}
