using System.Data.SqlClient;
using System.Data;

using iTextSharp.text;
using iTextSharp.text.pdf;
using System.Diagnostics;
using SpreadsheetLight;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
//using Excel = Microsoft.Office.Interop.Excel;

using OfficeOpenXml;
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.XmlAccess;
using Color = System.Drawing.Color;
using CobranzaSP.Modelos;
using System.IO;
using System;
using CobranzaSP.Formularios;
using System.Windows.Forms;

namespace CobranzaSP.Lógica
{
    internal class LogicaFusor
    {
        private CD_Conexion conexion = new CD_Conexion();
        private AccionesLógica NuevaAccion = new AccionesLógica();
        SqlCommand comando = new SqlCommand();
        bool Modificando = false;


        public void GuardarFusor(Fusor NuevoFusor, string sp)
        {
            comando.Connection = conexion.AbrirConexion();
            comando.CommandText = sp;
            comando.CommandType = CommandType.StoredProcedure;

            if (sp == "ModificarFusor")
            {
                comando.Parameters.AddWithValue("@Id", NuevoFusor.IdFusor);
                Modificando = true;
            }

            comando.Parameters.AddWithValue("@NumeroSerieO", NuevoFusor.SerieO);
            comando.Parameters.AddWithValue("@NumeroSerieS", NuevoFusor.SerieS);
            comando.Parameters.AddWithValue("@NumeroFactura", NuevoFusor.NumeroFactura);
            comando.Parameters.AddWithValue("@Proveedor", NuevoFusor.Proveedor);
            comando.Parameters.AddWithValue("@FechaFactura", NuevoFusor.FechaFactura);
            comando.Parameters.AddWithValue("@Costo", NuevoFusor.Cantidad);
            comando.Parameters.AddWithValue("@DiasGarantia", NuevoFusor.DiasGarantia);
            comando.Parameters.AddWithValue("@Estado", NuevoFusor.Estado);
            comando.Parameters.AddWithValue("@Modelo", NuevoFusor.Modelo);

            comando.ExecuteNonQuery();

            comando.Parameters.Clear();

            AgregarFusorAInventario(NuevoFusor);
            conexion.CerrarConexion();
        }

        public void AgregarFusorAInventario(Fusor NuevoFusor)
        {

            if (Modificando)
                return;
            //Restaremos del inventario el fusor que se utilizo en base a su modelo
            LogicaRegistro AccionRegistro = new LogicaRegistro();

            //Con ayuda de la clave del fusor, podemos obtener a traves de su modelo el idcartucho 
            int IdCartucho = NuevaAccion.BuscarId(NuevoFusor.Modelo, "BuscarIdCartucho");
            LogicaServicio logicaServicio = new LogicaServicio();
            RegistroInventario registroFusor = new RegistroInventario()
            {
                Cliente = NuevoFusor.Proveedor,
                IdMarca = logicaServicio.ObtenerMarcaFusor(IdCartucho),
                IdCartucho = IdCartucho,
                CantidadSalida = 0,
                CantidadEntrada = 1,
                Fecha = DateTime.Today,
                NumeroSerie = NuevoFusor.SerieS
            };
            string Mensaje = AccionRegistro.AgregarRegistroInventario(registroFusor);
            MessageBox.Show(Mensaje, "REGISTRO INVENTARIO", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        public SqlDataReader Mostrar(string sp)
        {
            SqlDataReader leer;
            comando.Connection = conexion.AbrirConexion();
            comando.CommandText = sp;
            comando.CommandType = CommandType.StoredProcedure;
            leer = comando.ExecuteReader();
            //conexion.CerrarConexion();
            return leer;
        }

        #region GeneracionExcel
        public void GenerarExcel()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            ExcelPackage ExcelPkg = new ExcelPackage();
            ExcelWorksheet wsSheet1 = ExcelPkg.Workbook.Worksheets.Add("Pagina 1");

            //Nos traemos los datos
            SqlDataReader fusores = Mostrar("MostrarTodosFusores");

            //Definimos los encabezados de cada columna de la tabla
            AgregarATablaExcel(wsSheet1, 1, 1, "Numero de serie", 14, true);
            AgregarATablaExcel(wsSheet1, 2, 1, "Numero de serie SpeedToner", 14, true);
            AgregarATablaExcel(wsSheet1, 3, 1, "Factura", 14, true); 
            AgregarATablaExcel(wsSheet1, 4, 1, "Proveedor", 14, true);
            AgregarATablaExcel(wsSheet1, 5, 1, "Fecha factura", 14, true);
            AgregarATablaExcel(wsSheet1, 6, 1, "Dias restantes", 14, true);
            AgregarATablaExcel(wsSheet1, 7, 1, "Precio", 14, true);
            AgregarATablaExcel(wsSheet1, 8, 1, "Dias Garantía", 14, true);
            AgregarATablaExcel(wsSheet1, 9, 1, "Garantía", 14, true);
            AgregarATablaExcel(wsSheet1, 10, 1, "Estado", 14, true);
            AgregarATablaExcel(wsSheet1, 11, 1, "Modelo", 14, true);

            //Insertamos los datos
            int fila = 2;
            while (fusores.Read())
            {
                DateTime FechaFactura = Convert.ToDateTime(fusores[5].ToString());

                AgregarATablaExcel(wsSheet1, 1, fila, fusores[1].ToString(), 12, false);
                AgregarATablaExcel(wsSheet1, 2, fila, fusores[2].ToString(), 12, false);
                AgregarATablaExcel(wsSheet1, 3, fila, fusores[3].ToString(), 12, false);
                AgregarATablaExcel(wsSheet1, 4, fila, fusores[4].ToString(), 12, false);
                AgregarATablaExcel(wsSheet1, 5, fila, FechaFactura.ToString("dd/MM/yyyy"), 12, false);
                AgregarATablaExcel(wsSheet1, 6, fila, fusores[6].ToString(), 12, false);
                AgregarATablaExcel(wsSheet1, 7, fila, fusores[7].ToString(), 12, false);
                AgregarATablaExcel(wsSheet1, 8, fila, fusores[8].ToString(), 12, false);
                AgregarATablaExcel(wsSheet1, 9, fila, fusores[9].ToString(), 12, false);
                AgregarATablaExcel(wsSheet1, 10, fila, fusores[10].ToString(), 12, false);
                AgregarATablaExcel(wsSheet1, 11, fila, fusores[11].ToString(), 12, false);
                fila++;
            }

            //AJUSTES AL DOCUMENTO AL FINAL
            wsSheet1.Row(4).Height = 30;
            //Ajustamos la celda al tamaño de su contenido
            wsSheet1.Cells[wsSheet1.Dimension.Address].AutoFitColumns();
            wsSheet1.Protection.IsProtected = false;
            wsSheet1.Protection.AllowSelectLockedCells = false;
            ExcelPkg.SaveAs(new FileInfo(@"\\DESKTOP-D3OQEJR\Archivos Compartidos\Reportes\Fusores\" + "ExcelFusores" + DateTime.Now.ToString("ddMMyyyyHHmmss") + ".xlsx"));
            conexion.CerrarConexion();
        }

        public void AgregarATablaExcel(ExcelWorksheet wsSheet1, int columna, int fila, string Valor, int TamañoLetra, bool negritas)
        {
            using (ExcelRange Rng = wsSheet1.Cells[fila, columna])
            {
                Rng.Value = Valor;
                Rng.Style.Font.Size = TamañoLetra;
                Rng.Style.Font.Bold = negritas;
            }
        }
        #endregion

        #region Pdf
        public void ReporteFusores(string Parametro, DateTime FechaInicio, DateTime FechaFinal, string Serie)
        {
            SqlDataReader leer;
            comando.Connection = conexion.AbrirConexion();
            comando.CommandText = "ReporteFusores";
            comando.CommandType = CommandType.StoredProcedure;
            comando.Parameters.AddWithValue("@Garantia", Parametro);
            comando.Parameters.AddWithValue("@FechaInicio", FechaInicio);
            comando.Parameters.AddWithValue("@FechaFinal", FechaFinal);
            comando.Parameters.AddWithValue("@Serie", Serie);
            leer = comando.ExecuteReader();//Guardamos la informacion obtenida, para posteriormente ir mostrandola en el pdf
            GenerarReporteFusores(leer, Parametro, FechaInicio, FechaFinal, Serie);
            leer.Close();
            comando.Parameters.Clear();
        }
        public void GenerarReporteFusores(SqlDataReader leer, string Parametro, DateTime FechaInicio, DateTime FechaFinal, string Serie)
        {
            string NombreArchivo = @"\\DESKTOP-D3OQEJR\Archivos Compartidos\Reportes\Fusores\" + "ReporteFusores" + DateTime.Now.ToString("ddMMyyyyHHmmss") + ".pdf";
            //string NombreArchivo = @"C:\Users\DELL PC\Documents\Base de datos\" + "ReporteFusores" + DateTime.Now.ToString("ddMMyyyyHHmmss") + ".pdf";
            //Lista de las series para evitar que se repitan las series en algunas consultas
            FileStream fs = new FileStream(NombreArchivo, FileMode.Create);
            Document document = new Document(PageSize.LETTER);
            document.SetMargins(25f, 25f, 25f, 25f);
            //Colocamos el pdf en horizontal
            document.SetPageSize(iTextSharp.text.PageSize.LETTER.Rotate());
            PdfWriter pw = PdfWriter.GetInstance(document, fs);

            //Variable que servira para que dependiendo lo que eliga el usuario, una columna abarque dos espacios, ya que una columna en especifico no aparecera
            int colspanSerie = 1;

            //Instanciamos para que podamos obtener los indices de cada hoja
            var pe = new Pdf();
            pw.PageEvent = pe;
            pe.ColocarFormatoSuperior = true;

            document.Open();

            Paragraph titulo = new Paragraph("REPORTE FUSORES GARANTIA " + Parametro.ToUpper() + " " + Serie.ToUpper(), pe.FuenteTitulo);
            titulo.Alignment = Element.ALIGN_CENTER;
            document.Add(titulo);

            document.Add(new Chunk());
            switch (Parametro)
            {
                case "Habilitado":
                    colspanSerie = 1
                        ; break;
                case "Deshabilitada": colspanSerie = 2; break;
                case "Rango Fecha":
                    Paragraph RangoFechas = new Paragraph("Fecha Inicial: " + string.Format("{0:d}", FechaInicio) + "    Fecha Final: " + string.Format("{0:d}", FechaFinal), pe.FuenteFecha) { Alignment = Element.ALIGN_CENTER };
                    document.Add(RangoFechas);
                    document.Add(new Chunk());
                    ; break;
                case "Serie": colspanSerie = 2; break;
            }

            PdfPTable Fusores = new PdfPTable(7);
            Fusores.WidthPercentage = 100;

            PdfPCell cltSerie = new PdfPCell(new Phrase("Serie", pe.FuenteParrafoBold)) { BorderWidth = .5f, Colspan = colspanSerie };
            Fusores.AddCell(cltSerie);
            if (Parametro != "Serie")
            {
                PdfPCell cltSerieSp = new PdfPCell(new Phrase("Serie Sp", pe.FuenteParrafoBold)) { BorderWidth = .5f, Colspan = 1 };
                Fusores.AddCell(cltSerieSp);
            }
            PdfPCell cltFactura = new PdfPCell(new Phrase("#Factura", pe.FuenteParrafoBold)) { BorderWidth = .5f, Colspan = 1 };
            PdfPCell cltFechaFactura = new PdfPCell(new Phrase("FechaFactura", pe.FuenteParrafoBold)) { BorderWidth = .5f, Colspan = 1 };
            PdfPCell cltProveedor = new PdfPCell(new Phrase("Proveedor", pe.FuenteParrafoBold)) { BorderWidth = .5f, Colspan = 1 };
            PdfPCell cltPrecio = new PdfPCell(new Phrase("Precio", pe.FuenteParrafoBold)) { BorderWidth = .5f, Colspan = 1 };


            Fusores.AddCell(cltFactura);
            Fusores.AddCell(cltFechaFactura);
            Fusores.AddCell(cltProveedor);
            Fusores.AddCell(cltPrecio);
            //Solo se añadira la columna dias restantes en dado caso de que sea diferente de fusores deshabilitados
            if (Parametro != "Deshabilitada")
            {
                PdfPCell cltDiasRestantes = new PdfPCell(new Phrase("Dias Restantes", pe.FuenteParrafoBold)) { BorderWidth = .5f, Colspan = 1 };
                Fusores.AddCell(cltDiasRestantes);
            }

            while (leer.Read())
            {
                PdfPCell clSerie = new PdfPCell(new Phrase(leer[1].ToString(), pe.FuenteParrafo)) { BorderWidth = .5f, Colspan = colspanSerie };
                Fusores.AddCell(clSerie);
                //Solo se mostrara en dado caso de que no estemos generando reporte por el numero de serie
                if (Parametro != "Serie")
                {
                    PdfPCell clSerieSp = new PdfPCell(new Phrase(leer[2].ToString(), pe.FuenteParrafo)) { BorderWidth = .5f, Colspan = 1 };
                    Fusores.AddCell(clSerieSp);
                }

                PdfPCell clFactura = new PdfPCell(new Phrase(leer[3].ToString(), pe.FuenteParrafo)) { BorderWidth = .5f, Colspan = 1 };
                PdfPCell clProveedor = new PdfPCell(new Phrase(leer[5].ToString(), pe.FuenteParrafo)) { BorderWidth = .5f, Colspan = 1 };
                PdfPCell clFechaFactura = new PdfPCell(new Phrase(string.Format("{0:d}", leer[4]), pe.FuenteParrafo)) { BorderWidth = .5f, Colspan = 1 };
                PdfPCell clPrecio = new PdfPCell(new Phrase(leer[7].ToString(), pe.FuenteParrafo)) { BorderWidth = .5f, Colspan = 1 };

                Fusores.AddCell(clFactura);
                Fusores.AddCell(clFechaFactura);
                Fusores.AddCell(clProveedor);
                Fusores.AddCell(clPrecio);
                //Solo se mostraran los dias restantes en dado caso de que no estemos generando reporte por fusores deshabilitados
                if (Parametro != "Deshabilitada")
                {
                    PdfPCell clDiasRestantes = new PdfPCell(new Phrase(leer[6].ToString(), pe.FuenteParrafo)) { BorderWidth = .5f };
                    Fusores.AddCell(clDiasRestantes);
                }
            }
            document.Add(Fusores);
            leer.Close();
            document.Close();
            //Abrimos el pdf 
            var p = new Process();
            p.StartInfo = new ProcessStartInfo(NombreArchivo)
            {
                UseShellExecute = true
            };
            p.Start();
        }
        #endregion
    }
}
