using CobranzaSP.Lógica;
using CobranzaSP.Modelos;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CobranzaSP.Formularios
{
    public partial class ReporteRegistrosInventarioPartesRicoh : Form
    {
        public ReporteRegistrosInventarioPartesRicoh()
        {
            InitializeComponent();
        }
        AccionesLógica NuevaAccion = new AccionesLógica();
        CD_Conexion cn = new CD_Conexion();
        LogicaRegistroInventarioRicoh NuevoRegistro = new LogicaRegistroInventarioRicoh();

        private void btnGenerar_Click(object sender, EventArgs e)
        {
            string Parametro = "";
            //if (cboBusqueda.SelectedIndex == -1)
            //    Parametro = "";
            //else
            //    Parametro = cboBusqueda.SelectedItem.ToString();

            //if (TipoBusqueda == "Modelo")
            //    Parametro = cboMarca.SelectedItem.ToString() +" " + Parametro;

            DateTime FechaInicio = dtpFechaInicio.Value;
            DateTime FechaFinal = dtpFechaFinal.Value;

            Reporte nuevoReporte = new Reporte()
            {
                FechaInicio = dtpFechaInicio.Value,
                FechaFinal = dtpFechaFinal.Value,
                ParametroBusqueda = Parametro
            };

            NuevoRegistro.ObtenerDatosReporte(nuevoReporte);
            //MessageBox.Show(NuevoRegistro.ObtenerDatosReporte(nuevoReporte), "REPORTE REGISTRO", MessageBoxButtons.OK, MessageBoxIcon.Information);

            //if (TipoBusqueda == "Modelo")
            //{
            //    string Marca = cboMarca.SelectedItem.ToString();
            //    string Cliente = "";
            //    if (chkCliente.Checked)
            //    {
            //        Cliente = cboClienteEspecifico.SelectedItem.ToString();
            //    }
            //    //En esta parte podemos comenzar con el filtro de cartuchos
            //    //MessageBox.Show(AccionRegistro.ReporteRegistroCartucho(FechaInicio, FechaFinal, TipoBusqueda, Parametro, Marca, Cliente), "REPORTE REGISTRO", MessageBoxButtons.OK, MessageBoxIcon.Information);

            //}
            //else
            //{
            //    //MessageBox.Show(AccionRegistro.ReporteRegistrosInventario(FechaInicio, FechaFinal, TipoBusqueda, Parametro), "REPORTE REGISTRO", MessageBoxButtons.OK, MessageBoxIcon.Information);
            //}
            //cboBusqueda.SelectedIndex = 0;
        }
    }
}
