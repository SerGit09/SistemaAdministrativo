using CobranzaSP.Modelos;
using iTextSharp.text.pdf;
using iTextSharp.text;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CobranzaSP.Lógica
{
    internal class LogicaInventarioPartesRicoh
    {
        private CD_Conexion conexion = new CD_Conexion();
        SqlCommand comando = new SqlCommand();

        public string RegistroParteRicoh(ParteRicoh nuevaParte, string sp)
        {
            int respuesta = 0;
            string AccionRealizada;
            comando.Connection = conexion.AbrirConexion();
            comando.CommandText = sp;
            comando.CommandType = CommandType.StoredProcedure;
            comando.Parameters.Clear();

            //Preguntamos que accion realizaremos en la base de datos para posteriormente mostrale al usuario la accion que realizo
            AccionRealizada = (sp == "AgregarParteRicoh") ? "agrego" : "modifico";

            if (nuevaParte.IdNumeroParte != 0)
            {
                comando.Parameters.AddWithValue("@Id", nuevaParte.IdNumeroParte);
            }
            comando.Parameters.AddWithValue("@NumeroParte", nuevaParte.NumeroParte);
            comando.Parameters.AddWithValue("@Descripcion", nuevaParte.Descripcion);
            comando.Parameters.AddWithValue("@Cantidad", nuevaParte.Cantidad);


            respuesta = comando.ExecuteNonQuery();
            string Mensaje = (respuesta > 0) ? "Registro se " + AccionRealizada + " correctamente" : "Algo salio mal, no se " + AccionRealizada + " el registro";

            comando.Parameters.Clear();
            conexion.CerrarConexion();
            return Mensaje;
        }

        public void ImprimirInventario()
        {
            //Primero obtenemos de la base de datos nuestro inventario 
            SqlDataReader Inventario;
            comando.Connection = conexion.AbrirConexion();
            comando.CommandText = "ReporteInventarioRicoh";
            comando.CommandType = CommandType.StoredProcedure;
            Inventario = comando.ExecuteReader();

            //string NombreArchivo = @"C:\Users\Cobranza\Documents\Reportes\" + "ReporteServicio" + DateTime.Now.ToString("ddMMyyyyHHmmss") + ".pdf";
            //string NombreArchivo = @"C:\Users\DELL PC\Documents\Base de datos\" + "Inventario" + DateTime.Now.ToString("dd-MM-yyyy") + ".pdf";
            string NombreArchivo = @"\\administracion-pc\ARCHIVOS COMPARTIDOS\Reportes\Inventario\" + "Reporte-Partes-Ricoh" + DateTime.Now.ToString("ddMMyyyyHHmmss") + ".pdf";
            FileStream fs = new FileStream(NombreArchivo, FileMode.Create);
            Document document = new Document(PageSize.LETTER);
            document.SetMargins(25f, 25f, 25f, 25f);
            document.SetPageSize(iTextSharp.text.PageSize.LETTER);
            PdfWriter pw = PdfWriter.GetInstance(document, fs);

            var pe = new Pdf();
            pe.ColocarFormatoSuperior = false;
            pw.PageEvent = pe;

            document.Open();

            Paragraph titulo = new Paragraph("EXISTENCIAS PARTES RICOH SPEED TONER", pe.FuenteTitulo18);
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

            PdfPTable tblInventario = new PdfPTable(4);
            //PdfPTable tblInventario = new PdfPTable(5);

            document.Add(new Paragraph("\n"));

            //Agregamos los titulos de la tabla
            PdfPCell clClave = new PdfPCell(new Phrase("CLAVE", pe.FuenteParrafoBold)) { BorderWidth = .5f, Colspan = 1 };
            PdfPCell clDescripcion = new PdfPCell(new Phrase("DESCRIPCION", pe.FuenteParrafoBold)) { BorderWidth = .5f, Colspan = 2 };

            //PdfPCell clCantidadNueva = new PdfPCell(new Phrase("EXISTENCIAS NUEVA", pe.FuenteParrafoBold)) { BorderWidth = .5f, Colspan = 1 };
            //PdfPCell clCantidadUsada = new PdfPCell(new Phrase("EXISTENCIAS USADA", pe.FuenteParrafoBold)) { BorderWidth = .5f, Colspan = 1 };
            PdfPCell clExistencias = new PdfPCell(new Phrase("EXISTENCIAS", pe.FuenteParrafoBold)) { BorderWidth = .5f, Colspan = 1 };

            tblInventario.AddCell(clClave);
            tblInventario.AddCell(clDescripcion);
            tblInventario.AddCell(clExistencias);

            while (Inventario.Read())
            {

                //PdfPCell clCantidadOficinaDato = new PdfPCell(new Phrase(Inventario[1].ToString(), pe.FuenteParrafo)) { BorderWidth = .5f, Colspan = 1 };
                PdfPCell clClaveDato = new PdfPCell(new Phrase(Inventario[0].ToString(), pe.FuenteParrafo)) { BorderWidth = .5f, Colspan = 1 };
                PdfPCell clDescripcionDato = new PdfPCell(new Phrase(Inventario[1].ToString(), pe.FuenteParrafo)) { BorderWidth = .5f, Colspan = 2 };
                PdfPCell clCantidadDato = new PdfPCell(new Phrase(ComprobarValoresO(Inventario[2].ToString()), pe.FuenteParrafo)) { BorderWidth = .5f, Colspan = 1 };
                //PdfPCell clCantidadUsadaDato = new PdfPCell(new Phrase(ComprobarValoresO(Inventario[3].ToString()), pe.FuenteParrafo)) { BorderWidth = .5f, Colspan = 1 };

                tblInventario.AddCell(clClaveDato);
                tblInventario.AddCell(clDescripcionDato);
                tblInventario.AddCell(clCantidadDato);
                //tblInventario.AddCell(clCantidadNuevaDato);
                //tblInventario.AddCell(clCantidadUsadaDato);
            }

            //Añadimos la tabla al documento
            document.Add(tblInventario);
            Inventario.Close();
            document.Close();

            //Abrimos el pdf
            pe.AbrirPdf(NombreArchivo);
            conexion.CerrarConexion();
        }

        //Metodo para evitar que salgan valores en 0 en el inventario
        public string ComprobarValoresO(string Cantidad)
        {
            if (Cantidad == "0")
            {
                return " ";
            }
            else
            {
                return Cantidad;
            }
        }


    }
}
