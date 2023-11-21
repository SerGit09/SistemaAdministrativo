using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.Data;
using CobranzaSP.Lógica;
using CobranzaSP.Modelos;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.Xml.Linq;
using System.Drawing;
using System.Diagnostics;
using CobranzaSP.Formularios;
using System.IO;

namespace CobranzaSP.Lógica
{
    internal class LogicaInventarioModulos
    {
        private CD_Conexion conexion = new CD_Conexion();
        SqlCommand comando = new SqlCommand();
        SqlDataReader Inventario;
        SortedSet<string> lstModulos = new SortedSet<string>();
        PdfPTable tblModulo;
        PdfPTable tblClaves;


        public void ObtenerDatosInventario()
        {
            comando.Connection = conexion.AbrirConexion();
            comando.CommandText = "ReporteInventarioModulos";
            comando.CommandType = CommandType.StoredProcedure;
            Inventario = comando.ExecuteReader();
        }

        public void ImprimirInventario()
        {
            //Primero obtenemos de la base de datos nuestro inventario 
            ObtenerDatosInventario();

            //string NombreArchivo = @"C:\Users\Cobranza\Documents\Reportes\" + "ReporteServicio" + DateTime.Now.ToString("ddMMyyyyHHmmss") + ".pdf";
            //string NombreArchivo = @"C:\Users\DELL PC\Documents\Base de datos\" + "Inventario" + DateTime.Now.ToString("dd-MM-yyyy") + ".pdf";
            string NombreArchivo = @"\\Desktop-de0cg86\archivos compartidos\Reportes\Inventario\" + "Reporte" + DateTime.Now.ToString("ddMMyyyyHHmmss") + ".pdf";
            FileStream fs = new FileStream(NombreArchivo, FileMode.Create);
            Document document = new Document(PageSize.LETTER);
            document.SetMargins(25f, 25f, 25f, 25f);
            document.SetPageSize(iTextSharp.text.PageSize.LETTER);
            PdfWriter pw = PdfWriter.GetInstance(document, fs);

            var pe = new Pdf();
            pe.ColocarFormatoSuperior = true;
            pw.PageEvent = pe;

            document.Open();

            Paragraph titulo = new Paragraph("EXISTENCIAS INVENTARIO DE MODULOS", pe.FuenteTitulo18);
            titulo.Alignment = Element.ALIGN_CENTER;
            document.Add(titulo);

            //Colocamos la imagen para gasolinas de los carros
            //iTextSharp.text.Image Logo = iTextSharp.text.Image.GetInstance(Properties.Resources.TanquesGasolina, System.Drawing.Imaging.ImageFormat.Png);
            //Logo.ScaleToFit(350, 250);
            //Logo.Alignment = iTextSharp.text.Image.UNDERLYING;
            //Logo.SetAbsolutePosition(document.LeftMargin, document.Top - 65);
            //document.Add(Logo);

            Paragraph Fecha = new Paragraph("Fecha: " + DateTime.Now.ToString("dd/MM/yyyy"), pe.FuenteFechaGrande) { Alignment = Element.ALIGN_RIGHT };
            document.Add(Fecha);

            Paragraph Hora = new Paragraph("Hora: " + DateTime.Now.ToString("hh:mm:ss tt"), pe.FuenteFechaGrande) { Alignment = Element.ALIGN_RIGHT };
            document.Add(Hora);

            tblModulo = new PdfPTable(3);
            tblClaves = new PdfPTable(1);
            //PdfPTable tblInventario = new PdfPTable(5);

            document.Add(new Paragraph("\n"));

            while (Inventario.Read())
            {
                ReporteInventarioModulo NuevoReporte = new ReporteInventarioModulo()
                {
                    Modelo = Inventario[0].ToString(),
                    Modulo = Inventario[1].ToString(),
                    Cantidad = int.Parse(Inventario[2].ToString()),
                    Clave = Inventario[3].ToString(),
                };
                string Modulo = NuevoReporte.Modelo + " " + NuevoReporte.Modulo;
                if (!lstModulos.Contains(Modulo))
                {
                    document.Add(tblClaves);
                    iTextSharp.text.pdf.draw.LineSeparator LineaSeparacion = new iTextSharp.text.pdf.draw.LineSeparator() { Offset = 2f };
                    document.Add(new Chunk(LineaSeparacion));
                    //Limpiamos la tabla
                    
                    tblClaves = new PdfPTable(1);
                    //Lo guardamos en la lista y lo mostramos en el documento
                    lstModulos.Add(Modulo);
                    

                    if(NuevoReporte.Cantidad > 0)
                    {
                        CrearTablaModulo(NuevoReporte);
                        //Agregamos la tabla al documento
                        document.Add(tblModulo);

                        Paragraph lineBreak = new Paragraph("\n");
                        document.Add(lineBreak);
                        //Agregamos las claves
                        CrearTablaClaves(NuevoReporte);
                    }
                    
                    //Limpiamos la tabla
                    tblModulo = new PdfPTable(3);
                }
                else
                {
                    AgregarClaveATabla(NuevoReporte.Clave);
                }
            }

            //Añadimos la tabla al documento
            //document.Add(tblInventario);
            lstModulos.Clear();
            Inventario.Close();
            document.Close();

            //Abrimos el pdf
            pe.AbrirPdf(NombreArchivo);
            conexion.CerrarConexion();
        }

        public void CrearTablaModulo(ReporteInventarioModulo NuevoReporte)
        {
            Pdf pe = new Pdf();
            //TITULOS
            PdfPCell clTituloModelo = new PdfPCell(new Phrase("MODELO", pe.FuenteParrafoBold)) { BorderWidth = .5f, Colspan = 1 };
            PdfPCell clTituloModulo = new PdfPCell(new Phrase("MODULO", pe.FuenteParrafoBold)) { BorderWidth = .5f, Colspan = 1 };
            PdfPCell clTituloCantidad = new PdfPCell(new Phrase("CANTIDAD", pe.FuenteParrafoBold)) { BorderWidth = .5f, Colspan = 1 };

            tblModulo.AddCell(clTituloModelo);
            tblModulo.AddCell(clTituloModulo);
            tblModulo.AddCell(clTituloCantidad);

            //DATOS
            PdfPCell clModelo = new PdfPCell(new Phrase(NuevoReporte.Modelo, pe.FuenteParrafo)) { BorderWidth = .5f, Colspan = 1 };
            PdfPCell clModulo = new PdfPCell(new Phrase(NuevoReporte.Modulo, pe.FuenteParrafo)) { BorderWidth = .5f, Colspan = 1 };
            PdfPCell clCantidad = new PdfPCell(new Phrase(NuevoReporte.Cantidad.ToString(), pe.FuenteParrafo)) { BorderWidth = .5f, Colspan = 1 };

            tblModulo.AddCell(clModelo);
            tblModulo.AddCell(clModulo);
            tblModulo.AddCell(clCantidad);
        }

        public void CrearTablaClaves(ReporteInventarioModulo NuevoReporte)
        {
            Pdf pe = new Pdf();
            PdfPCell clTituloClave = new PdfPCell(new Phrase("CLAVE", pe.FuenteParrafoBold)) { BorderWidth = .5f, Colspan = 1 , HorizontalAlignment = Element.ALIGN_CENTER };
            tblClaves.AddCell(clTituloClave);
            AgregarClaveATabla(NuevoReporte.Clave);
        }

        public void AgregarClaveATabla(string Clave)
        {
            Pdf pe = new Pdf();
            PdfPCell clClave = new PdfPCell(new Phrase(Clave, pe.FuenteParrafo)) { BorderWidth = .5f, Colspan = 1, HorizontalAlignment = Element.ALIGN_CENTER };
            tblClaves.AddCell(clClave);
        }
    }
}
