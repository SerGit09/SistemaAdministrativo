using CobranzaSP.Lógica;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using CobranzaSP.Modelos;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;
using System.Data.SqlClient;
using System.Globalization;

namespace CobranzaSP.Formularios
{
    public partial class Graficas : Form
    {
        public Graficas()
        {
            InitializeComponent();
            InicioAplicacion();
            //Grafica();
        }
        private CD_Conexion conexion = new CD_Conexion();
        SqlCommand comando = new SqlCommand();
        ArrayList Totales = new ArrayList();
        ArrayList TipoPago = new ArrayList();
        GraficaLogica accionGrafica = new GraficaLogica();
        //Series serie = new Series();

        string[] Meses = new string[12] { "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre" };

        #region Inicio

        public void InicioAplicacion()
        {
            TipoPago.Add("FACTURA");
            TipoPago.Add("REMISION");
            TipoPago.Add("MOSTRADOR");
            LlenarDtg(dtgMesesCobranza);
            LlenarDtg(dtgTotalFacturado);

            //Para que el datetimepicker muestre solamente el año
            dtpAño.Format = DateTimePickerFormat.Custom;
            dtpAño.CustomFormat = "yyyy";
            dtpAño.ShowUpDown = true;
            //AgregarMesesGrafica();
            LlenarTablaFacturado();
            LLenarTablaCobrado();
        }

        public void LlenarTablaFacturado()
        {
            DataTable tabla = new DataTable();
            string[] Cantidades = new string[12]; 
            //Guardamos los registros dependiendo la consulta
            tabla = accionGrafica.ObtenerTablaResultados("GraficaResultados");
            dtgTotalFacturado.DataSource = tabla;
            int contador = 0;
            chGrafica.Palette = ChartColorPalette.Pastel;
            foreach (DataRow row in tabla.Rows) 
            {
                Cantidades[contador] = row[4].ToString();
                contador++;
            }


            for (int i = 0; i < Meses.Length; i++)
            {
                //Titulos
                Series serie = chGrafica.Series.Add(Meses[i]);

                //Ajustamos el tamaño de las columnas
                serie["PixelPointWidth"] = "350";

                //Agregamos loa datos
                serie.Points.AddXY(Meses[i], Cantidades[i]);
            }
            
        }

        public void LLenarTablaCobrado()
        {
            DataTable tabla = new DataTable();
            //Guardamos los registros dependiendo la consulta
            tabla = accionGrafica.ObtenerTablaCobradoMeses("TotalCobrado");
            dtgMesesCobranza.DataSource = tabla;
        }

        public void LlenarDtg(DataGridView dtg) 
        {
            dtg.ReadOnly = true;

            //No agregar renglones
            dtg.AllowUserToAddRows = false;

            //No borrar renglones
            dtg.AllowUserToDeleteRows = false;

            //Ajustar automaticamente el ancho de las columnas
            dtg.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            dtg.AutoResizeColumns(DataGridViewAutoSizeColumnsMo‌​de.Fill);

            // Ajustar automáticamente el ancho de las columnas
            
            //dtg.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;

            // Ajustar automáticamente la altura de las filas
            dtg.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;

            dtg.AutoResizeColumns();
            dtg.AutoResizeRows(DataGridViewAutoSizeRowsMode.AllCells);


        }
        #endregion


        public void Grafica()
        {
            int Año = int.Parse(dtpAño.Value.Year.ToString());
            //Año = 2022;

            for (int i = 0; i < Meses.Length; i++)
            {
                //Debo evitar que se llene 2 veces
                //serie = chGrafica.Series.Add(Meses[i]);

                double totalMes = accionGrafica.ObtenerTotalMes(Año, i + 1, "TotalCobradoMeses");

                //serie.Label = totalMes.ToString();
                //serie.Points.Add(totalMes);

                //chGrafica.Series[0].Points.AddXY("s",totalMes);
                string Total = totalMes.ToString("C", CultureInfo.CurrentCulture);
                dtgMesesCobranza.Rows.Add(Meses[i],0,0,0, Total);
            }

            for (int i = 0; i < Meses.Length; i++)
            {
                double totalMes = accionGrafica.ObtenerTotalMes(Año,i + 1, "TotalFacturadoMeses");
                //serie.Label = totalMes.ToString();

                string Total = totalMes.ToString("C", CultureInfo.CurrentCulture);
                dtgTotalFacturado.Rows.Add(Meses[i],0,0,0, Total);
            }
        }

        private void btnMostrar_Click(object sender, EventArgs e)
        {
            dtgMesesCobranza.Rows.Clear();
            dtgTotalFacturado.Rows.Clear();
        }
    }
}
