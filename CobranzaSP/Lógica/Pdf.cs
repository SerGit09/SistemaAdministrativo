using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.Xml.Linq;
using System.Diagnostics;

namespace CobranzaSP.Lógica
{
    internal class Pdf:PdfPageEventHelper
    {
        //Permite escribir sobre un documento ya existente
        PdfContentByte cb;
        PdfTemplate template;

        //Por si queremos agregar alguna fuente
        BaseFont bf = null;
        BaseFont title = null;
        //Para añadir la fecha de impresion
        DateTime PrintTime = DateTime.Now;
        iTextSharp.text.Font font = new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 11, iTextSharp.text.Font.BOLD, BaseColor.BLACK);

        //Nos ayudara a controlar si mostrar los formatos superiores en los pdfs
        public bool ColocarFormatoSuperior { get; set; }

        //Fuentes globales para todos los pdfs
        public iTextSharp.text.Font FuenteTitulo = new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 14, iTextSharp.text.Font.BOLD, BaseColor.BLACK);
        public iTextSharp.text.Font FuenteTitulo18 = new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 18, iTextSharp.text.Font.BOLD, BaseColor.BLACK);
        public iTextSharp.text.Font FuenteParrafo = new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.TIMES_ROMAN, 10, iTextSharp.text.Font.NORMAL, BaseColor.BLACK);

        public iTextSharp.text.Font FuenteParrafoGrande = new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.TIMES_ROMAN, 12, iTextSharp.text.Font.BOLD, BaseColor.BLACK);
        //Fuente para los parrafos en negritas
        public iTextSharp.text.Font FuenteParrafoBold = new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.TIMES_ROMAN, 11, iTextSharp.text.Font.BOLD, BaseColor.BLACK);
        public iTextSharp.text.Font FuenteParrafoBold10 = new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.TIMES_ROMAN, 10, iTextSharp.text.Font.BOLD, BaseColor.BLACK);
        public iTextSharp.text.Font FuenteFecha = new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 16, iTextSharp.text.Font.BOLD, BaseColor.BLACK);
        public iTextSharp.text.Font FuenteFechaGrande = FontFactory.GetFont("arial", 13);



        public void AbrirPdf(string NombreArchivo)
        {
            //Abrimos el pdf 
            var p = new Process();
            p.StartInfo = new ProcessStartInfo(NombreArchivo)
            {
                UseShellExecute = true
            };
            p.Start();
        }

        public void ColocarFormatosSuperiores(Document document)
        {
            //Agregamos el logo de la izquierda
            iTextSharp.text.Image Logo = iTextSharp.text.Image.GetInstance(Properties.Resources.LogoSpeedToner, System.Drawing.Imaging.ImageFormat.Png);
            Logo.ScaleToFit(150, 80);
            //Logo.ScaleToFit(100, 80);
            Logo.Alignment = iTextSharp.text.Image.UNDERLYING;
            Logo.SetAbsolutePosition(document.LeftMargin, document.Top - 50);
            document.Add(Logo);

            //Agregamos el logo de la derecha
            iTextSharp.text.Image Logotipo = iTextSharp.text.Image.GetInstance(Properties.Resources.LogoSpeedToner, System.Drawing.Imaging.ImageFormat.Png);
            Logotipo.ScaleToFit(150, 80);
            //Logotipo.ScaleToFit(100, 80);
            Logotipo.Alignment = iTextSharp.text.Image.UNDERLYING;
            Logotipo.SetAbsolutePosition(document.Right - 150, document.Top - 50);
            //Logotipo.SetAbsolutePosition(document.Right - 100, document.Top - 50);
            document.Add(Logotipo);

            Paragraph NombreEmpresa = new Paragraph("SPEED TONER NUEVO LAREDO.", FuenteTitulo) { Alignment = Element.ALIGN_CENTER };
            document.Add(NombreEmpresa);

            Paragraph Telefono = new Paragraph("TEL.: (867) 712-0964 - (867)190-0381", FuenteTitulo) { Alignment = Element.ALIGN_CENTER };
            document.Add(Telefono);

            Paragraph Direccion = new Paragraph("LEONA VICARIO #1709 NUEVO LAREDO, TAMPS. C.P 88060", FuenteTitulo) { Alignment = Element.ALIGN_CENTER };
            document.Add(Direccion);
        }

        public override void OnOpenDocument(PdfWriter writer, Document document)
        {
            try
            {
                PrintTime = DateTime.Now;
                //Fuente que se utilizara en la clase

                bf = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                title = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1250, BaseFont.NOT_EMBEDDED);
                //Para poder agregar contenido
                cb = writer.DirectContent;
                //Sirve como canvas 
                template = cb.CreateTemplate(document.PageSize.Width, 50);
                if (ColocarFormatoSuperior)
                {
                    ColocarFormatosSuperiores(document);
                }


            }
            catch (DocumentException de)
            {

            }
            catch (System.IO.IOException ioe)
            {

            }
        }
        public override void OnEndPage(PdfWriter writer, Document document)
        {
            base.OnEndPage(writer, document);
            //Obtenemos la pagina actual
            int pageN = writer.PageNumber;
            String text = "Page " + pageN;
            //Tamaño del texto
            float len = bf.GetWidthPoint(text, 8);
            //Creamos rectangulo del tamaño de la pagina
            iTextSharp.text.Rectangle pagesize = document.PageSize;
            cb.SetRGBColorFill(100, 100, 100);
            cb.BeginText();
            cb.SetFontAndSize(bf, 8);
            cb.SetTextMatrix(pagesize.GetLeft(40), pagesize.GetBottom(10));
            cb.ShowText(text);
            cb.EndText();
            //Definimos otra platilla 
            cb.AddTemplate(template, pagesize.GetLeft(40) + len, pagesize.GetBottom(30));

            //Imprime la hora en la que fue impreso
            //Opcional
            cb.BeginText();
            cb.SetFontAndSize(bf, 8);
            cb.ShowTextAligned(PdfContentByte.ALIGN_RIGHT,
                "Printed On " + PrintTime.ToString(),
                pagesize.GetRight(40),
                pagesize.GetBottom(10), 0);
            cb.EndText();
        }

        public override void OnCloseDocument(PdfWriter writer, Document document)
        {
            base.OnCloseDocument(writer, document);
            template.BeginText();
            template.SetFontAndSize(bf, 8);
            template.SetTextMatrix(0, 0);
            //template.ShowText("" + (writer.PageNumber - 1));
            template.EndText();
        }
    }
}
