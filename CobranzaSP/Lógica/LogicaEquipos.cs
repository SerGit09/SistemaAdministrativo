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
using System.IO;
using SpreadsheetLight;
using System.Xml;
using CobranzaSP.Formularios;
using DocumentFormat.OpenXml.Office2013.Excel;
using System.Windows.Forms;

namespace CobranzaSP.Lógica
{
    internal class LogicaEquipos
    {
        private CD_Conexion conexion = new CD_Conexion();
        SqlCommand comando = new SqlCommand();
        SortedSet<string> lstMarcas = new SortedSet<string>();
        SortedSet<string> lstModelos = new SortedSet<string>();
        SortedSet<string> lstClientes = new SortedSet<string>();
        PdfPTable tblDatosEquipos;
        string TipoBusqueda;
        SqlDataReader reporte;
        string ParametroBusqueda;
        bool MostrarPrecios = false;
        string Modelo = "";


        public void GuardarRegistro(Equipo nuevoEquipo, string sp)
        {
            comando.Connection = conexion.AbrirConexion();
            comando.CommandText = sp;
            comando.CommandType = CommandType.StoredProcedure;

            if (sp == "ModificarEquipo")
            {
                comando.Parameters.AddWithValue("@Id", nuevoEquipo.IdEquipo);
            }

            comando.Parameters.AddWithValue("@IdCliente", nuevoEquipo.IdCliente);
            comando.Parameters.AddWithValue("@Ubicacion", nuevoEquipo.Ubicacion);
            comando.Parameters.AddWithValue("@IdMarca", nuevoEquipo.IdMarca);
            comando.Parameters.AddWithValue("@IdModelo", nuevoEquipo.IdModelo);
            comando.Parameters.AddWithValue("@Serie", nuevoEquipo.Serie);
            comando.Parameters.AddWithValue("@IdRenta", nuevoEquipo.IdRenta);
            comando.Parameters.AddWithValue("@Precio", nuevoEquipo.Precio);
            comando.Parameters.AddWithValue("@Fecha_Pago", nuevoEquipo.FechaPago);
            comando.Parameters.AddWithValue("@Valor", nuevoEquipo.Valor);

            comando.ExecuteNonQuery();

            comando.Parameters.Clear();
            conexion.CerrarConexion();
        }

        public void GuardarEquipoVendido(Equipo nuevoEquipo)
        {
            comando.Connection = conexion.AbrirConexion();
            comando.CommandText = "AgregarEquipoVendido";
            comando.CommandType = CommandType.StoredProcedure;
            comando.Parameters.Clear();
            comando.Parameters.AddWithValue("@IdMarca", nuevoEquipo.IdMarca);
            comando.Parameters.AddWithValue("@IdModelo", nuevoEquipo.IdModelo);
            comando.Parameters.AddWithValue("@Serie", nuevoEquipo.Serie);
            comando.Parameters.AddWithValue("@Fecha", nuevoEquipo.FechaVenta);
            comando.Parameters.AddWithValue("@Precio", nuevoEquipo.Precio);
            comando.Parameters.AddWithValue("@IdCliente", nuevoEquipo.IdCliente);

            comando.ExecuteNonQuery();

            
            conexion.CerrarConexion();
        }

        //Muestra los equipos dependiendo lo que necesite el usuario
        public bool OrdenarEquiposPrueba(ReporteEquipoParametro NuevoReporte, string StoredProcedure)
        {
            bool DatosObtenidos = true;
            TipoBusqueda = NuevoReporte.TipoBusqueda;
            ParametroBusqueda = NuevoReporte.Parametro;
            comando.Connection = conexion.AbrirConexion();
            comando.CommandText = StoredProcedure;
            comando.CommandType = CommandType.StoredProcedure;
            comando.Parameters.Clear();
            comando.Parameters.AddWithValue("@ParametroBusqueda", NuevoReporte.Parametro);
            comando.Parameters.AddWithValue("@Marca", NuevoReporte.Marca);
            comando.Parameters.AddWithValue("@Modelo", NuevoReporte.Modelo);
            reporte = comando.ExecuteReader();
            if (!reporte.HasRows)
            {
                reporte.Close();
                return DatosObtenidos = false;
            }

            Modelo = "";
            switch (TipoBusqueda)
            {
                case "Cliente":
                    GenerarReporteEquipos(reporte);
                    break;
                case "Precios de Equipos":
                    MostrarPrecios = true;
                    Modelo = NuevoReporte.Modelo;
                    GenerarPdf();
                    break;
            }
            
            reporte.Close();
            return DatosObtenidos;
        }

        public void GenerarPdf()
        {
            string NombreArchivo = @"\\DESKTOP-D3OQEJR\Archivos Compartidos\Reportes\Equipos\" + "ReporteEquipos" + DateTime.Now.ToString("ddMMyyyyHHmmss") + ".pdf";
            //string NombreArchivo = @"\\administracion-pc\ARCHIVOS COMPARTIDOS\Reportes\Equipos\" + "ReporteEquipos" + DateTime.Now.ToString("ddMMyyyyHHmmss") + ".pdf";


            FileStream fs = new FileStream(NombreArchivo, FileMode.Create);
            Document document = new Document(PageSize.LETTER);
            document.SetMargins(25f, 25f, 25f, 25f);
            //Colocamos el pdf en horizontal
            document.SetPageSize(iTextSharp.text.PageSize.LETTER);
            PdfWriter pw = PdfWriter.GetInstance(document, fs);

            //Instanciamos la clase para la paginacion
            var pe = new Pdf();
            pe.ColocarFormatoSuperior = true;
            pw.PageEvent = pe;

            document.Open();
            
            tblDatosEquipos = (TipoBusqueda != "Modelo") ? new PdfPTable(4) : new PdfPTable(3);
            if (MostrarPrecios)
            {
                tblDatosEquipos = (TipoBusqueda != "Modelo") ? new PdfPTable(6) : new PdfPTable(5);
            }

            //Colocar el titulo del reporte
            string NombreReporte = ColocarTituloReporte();
            Paragraph titulo = new Paragraph(NombreReporte, pe.FuenteTitulo) { Alignment = Element.ALIGN_CENTER };
            document.Add(titulo);
            document.Add(new Chunk("\n"));

            while (reporte.Read())
            {
                ReporteEquipo equipo = new ReporteEquipo()
                {
                    Marca = reporte[0].ToString(),
                    Modelo = reporte[1].ToString(),
                    Serie = reporte[3].ToString(),
                    Precio = double.Parse(reporte[4].ToString()),
                    Cliente = reporte[2].ToString()
                };

                ColocarSerieADocumento(equipo, document);
            }
            document.Add(tblDatosEquipos);
            lstMarcas.Clear();
            reporte.Close();

            iTextSharp.text.pdf.draw.LineSeparator lineSeparator = new iTextSharp.text.pdf.draw.LineSeparator();
            lineSeparator.Offset = 1f;
            document.Add(new iTextSharp.text.Chunk(lineSeparator));
            //Ahora deberemos de obtener la cantidad total en dinero de los equipos
            if(MostrarPrecios)
            {
                ColocarTotales(document);
            }
            

            document.Close();


            //Abrimos el pdf 
            pe.AbrirPdf(NombreArchivo);
        }

        public void ColocarSerieADocumento(ReporteEquipo equipo, Document document)
        {
            var pe = new Pdf();

            if (!lstMarcas.Contains(equipo.Marca))
            {
                document.Add(tblDatosEquipos);
                tblDatosEquipos = (TipoBusqueda == "Modelo") ? new PdfPTable(5) : new PdfPTable(6);

                //Agregamos la marca al documento y a la lista
                lstMarcas.Add(equipo.Marca);
                if (TipoBusqueda != "Marca")
                {
                    Paragraph Marca = new Paragraph(equipo.Marca, pe.FuenteParrafoGrande) { Alignment = Element.ALIGN_LEFT };
                    document.Add(Marca);
                }

                if(TipoBusqueda != "Modelo")
                {
                    AgregarTituloColumnaTabla("Modelo", 1);
                }
                
                AgregarTituloColumnaTabla("Cliente", 2);
                AgregarTituloColumnaTabla("Serie", 2);
                AgregarTituloColumnaTabla("Precio", 1);
                AgregarModelosATabla(equipo);
            }
            else
            {
                AgregarModelosATabla(equipo);
            }
        }

        public void AgregarModelosATabla(ReporteEquipo equipo)
        {
            var pe = new Pdf();

            if (TipoBusqueda != "Modelo")
            {
                PdfPCell clModelo = new PdfPCell(new Phrase(equipo.Modelo, pe.FuenteParrafo)) { BorderWidth = .5f, Colspan = 1 };
                tblDatosEquipos.AddCell(clModelo);
            }

            if(equipo.Cliente != "")
            {
                PdfPCell clCliente = new PdfPCell(new Phrase(equipo.Cliente, pe.FuenteParrafo)) { BorderWidth = .5f, Colspan = 2 };
                tblDatosEquipos.AddCell(clCliente);
            }
            
            PdfPCell clSerie = new PdfPCell(new Phrase(equipo.Serie, pe.FuenteParrafo)) { BorderWidth = .5f, Colspan = 2 };
            PdfPCell clPrecio = new PdfPCell(new Phrase("$" + String.Format("{0:n2}", equipo.Precio), pe.FuenteParrafo)) { BorderWidth = .5f, Colspan = 1 };
            
            tblDatosEquipos.AddCell(clSerie);
            tblDatosEquipos.AddCell(clPrecio);
        }

        public void GenerarReporteEquipos(SqlDataReader leer)
        {
            string NombreArchivo = @"\\DESKTOP-D3OQEJR\Archivos Compartidos\Reportes\Equipos\" + "ReporteEquipos" + DateTime.Now.ToString("ddMMyyyyHHmmss") + ".pdf";
            //string NombreArchivo = @"C:\Users\DELL PC\Documents\Base de datos\" + "ReporteEquipos" + DateTime.Now.ToString("ddMMyyyyHHmmss") + ".pdf";
            SaveFileDialog guardar = new SaveFileDialog();
            //Lista de las series para evitar que se repitan las series en algunas consultas
            FileStream fs = new FileStream(NombreArchivo, FileMode.Create);
            Document document = new Document(PageSize.LETTER);
            document.SetMargins(25f, 25f, 25f, 25f);
            //Colocamos el pdf en horizontal
            document.SetPageSize(iTextSharp.text.PageSize.LETTER.Rotate());
            PdfWriter pw = PdfWriter.GetInstance(document, fs);
            int TamañoTabla = 0;

            var pe = new Pdf();
            pw.PageEvent = pe;
            pe.ColocarFormatoSuperior = true;
            document.Open();
            this.TipoBusqueda = TipoBusqueda;
            //TIPO DE FUENTE

            Paragraph titulo = new Paragraph(ColocarTituloReporte(), pe.FuenteTitulo);
            titulo.Alignment = Element.ALIGN_CENTER;

            document.Add(titulo);


            iTextSharp.text.pdf.draw.LineSeparator lineSeparator = new iTextSharp.text.pdf.draw.LineSeparator();
            lineSeparator.Offset = 1f;
            document.Add(new iTextSharp.text.Chunk(lineSeparator));
            tblDatosEquipos = (TipoBusqueda != "Modelo") ? new PdfPTable(5) :new PdfPTable(4); 
            

            while (leer.Read())
            {
                ReporteEquipo DatosEquipo = new ReporteEquipo()
                {
                    Cliente = leer[1].ToString(),
                    Marca = leer[3].ToString(),
                    Modelo = leer[4].ToString(),
                    Serie = leer[5].ToString(),
                    TipoRenta = leer[6].ToString(),
                    Precio = double.Parse(leer[7].ToString()),
                    FechaPago = leer[8].ToString()
                };
                if (!lstClientes.Contains(leer[1].ToString()))
                {
                    lstMarcas.Clear();
                    document.Add(tblDatosEquipos);
                    tblDatosEquipos = new PdfPTable(5);
                    //Agregamos el cliente al documento y a la lista
                    Paragraph Cliente = new Paragraph(DatosEquipo.Cliente, pe.FuenteParrafoGrande) { Alignment = Element.ALIGN_LEFT };
                    document.Add(Cliente);
                    lstClientes.Add(DatosEquipo.Cliente);


                    AgregarMarcaAlDocumento(document, DatosEquipo);
                }
                else
                {
                    if(TipoBusqueda != "Marca")
                    {
                        AgregarMarcaAlDocumento(document, DatosEquipo);
                    }
                    else
                    {
                        AgregarDatosATabla(DatosEquipo);
                    }
                    
                }
            }
            //Añadimos la tabla al documento
            document.Add(tblDatosEquipos);

            lstClientes.Clear();
            lstMarcas.Clear();
            document.Close();

            //Abrimos el pdf

            pe.AbrirPdf(NombreArchivo);
        }

        public string ColocarTituloReporte()
        {
            string Titulo = "EQUIPOS EN RENTA POR ";
            switch (TipoBusqueda)
            {
                case "Cliente": 
                    Titulo += (ParametroBusqueda == "")? " CLIENTES": "CLIENTE:" + ParametroBusqueda;
                    break;
                case "Precios de Equipos":
                    Titulo = "VALOR DE EQUIPOS EN RENTA";
                    break;
            }
            return Titulo;
        }

        public void AgregarMarcaAlDocumento(Document document, ReporteEquipo DatosReporte)
        {
            var pe = new Pdf();
            if (!lstMarcas.Contains(DatosReporte.Marca))
            {
                document.Add(tblDatosEquipos);
                if (TipoBusqueda != "Modelo" && TipoBusqueda != "Marca")
                {
                    //Agregamos la marca al documento y a la lista
                    Paragraph Cliente = new Paragraph(DatosReporte.Marca, pe.FuenteParrafoGrande) { Alignment = Element.ALIGN_LEFT };
                    document.Add(Cliente);
                }
                lstMarcas.Add(DatosReporte.Marca);
                CrearTablaDatosEquipos(DatosReporte);
            }
            else
            {
                AgregarDatosATabla(DatosReporte);
            }
        }

        public void CrearTablaDatosEquipos(ReporteEquipo reporte)
        {
            tblDatosEquipos = new PdfPTable(4);
            if(TipoBusqueda != "Modelo")
            {
                tblDatosEquipos = new PdfPTable(5);
                AgregarTituloColumnaTabla("MODELO", 1);
            }
            AgregarTituloColumnaTabla("SERIE", 1);
            AgregarTituloColumnaTabla("TIPO RENTA", 1);
            AgregarTituloColumnaTabla("PRECIO", 1);
            AgregarTituloColumnaTabla("FECHA DE PAGO", 1);

            AgregarDatosATabla(reporte);
        }

        public void AgregarTituloColumnaTabla(string celda, int colspan)
        {
            Pdf pdf = new Pdf();
            PdfPCell clTitulo = new PdfPCell(new Phrase(celda, pdf.FuenteParrafoBold)) { BorderWidth = .5f, Colspan = colspan };
            tblDatosEquipos.AddCell(clTitulo);
        }

        public void AgregarDatosATabla(ReporteEquipo DatosEquipo)
        {
            Pdf pe = new Pdf();
            if(TipoBusqueda != "Modelo")
            {
                PdfPCell clModeloDato = new PdfPCell(new Phrase(DatosEquipo.Modelo, pe.FuenteParrafoBold)) { BorderWidth = .5f, Colspan = 1 };
                tblDatosEquipos.AddCell(clModeloDato);
            }
            PdfPCell clSerieDato = new PdfPCell(new Phrase(DatosEquipo.Serie, pe.FuenteParrafoBold)) { BorderWidth = .5f, Colspan = 1 };
            PdfPCell clTipoPagoDato = new PdfPCell(new Phrase(DatosEquipo.TipoRenta, pe.FuenteParrafoBold)) { BorderWidth = .5f, Colspan = 1 };
            PdfPCell clPrecioDato = new PdfPCell(new Phrase("$" + DatosEquipo.Precio, pe.FuenteParrafo)) { BorderWidth = .5f, Colspan = 1 };
            PdfPCell clFechaPagoDato = new PdfPCell(new Phrase(DatosEquipo.FechaPago, pe.FuenteParrafoBold)) { BorderWidth = .5f, Colspan = 1 };

            tblDatosEquipos.AddCell(clSerieDato);
            tblDatosEquipos.AddCell(clTipoPagoDato);
            tblDatosEquipos.AddCell(clPrecioDato);
            tblDatosEquipos.AddCell(clFechaPagoDato);
        }

        //Metodos para mostrar la cantidad y precio totales de los equipos
        public void ColocarTotales(Document document)
        {
            var pe = new Pdf();
            int CantidadTotal = 0;
            double PrecioTotal = 0;

            Paragraph titulo = new Paragraph("TOTAL", pe.FuenteTitulo) { Alignment = Element.ALIGN_CENTER };
            document.Add(titulo);
            //document.Add(new Chunk("\n"));

            tblDatosEquipos = new PdfPTable(3);
            PdfPCell clTipo = new PdfPCell();
            clTipo = new PdfPCell(new Phrase("MARCA", pe.FuenteParrafoBold)) { BorderWidth = .5f, Colspan = 1 };

            if (TipoBusqueda == "Modelo")
            {
                clTipo = new PdfPCell(new Phrase("MODELO", pe.FuenteParrafoBold)) { BorderWidth = .5f, Colspan = 1 };
            }

            PdfPCell clPrecio = new PdfPCell(new Phrase("PRECIO", pe.FuenteParrafoBold)) { BorderWidth = .5f, Colspan = 1 };
            PdfPCell clCantidad = new PdfPCell(new Phrase("CANTIDAD", pe.FuenteParrafoBold)) { BorderWidth = .5f, Colspan = 1 };
            tblDatosEquipos.AddCell(clTipo);
            tblDatosEquipos.AddCell(clPrecio);
            tblDatosEquipos.AddCell(clCantidad);

            if(Modelo != "")
            {
                ParametroBusqueda = Modelo;
            }
            ObtenerPreciosTotalesEquipos();

            while (reporte.Read())
            {
                PdfPCell clDatoTipo = (TipoBusqueda == "Modelo") ? new PdfPCell(new Phrase(ParametroBusqueda, pe.FuenteParrafo)) { BorderWidth = .5f, Colspan = 1 } : new PdfPCell(new Phrase(reporte[0].ToString(), pe.FuenteParrafo)) { BorderWidth = .5f, Colspan = 1 };
                PdfPCell clDatosTotal = new PdfPCell(new Phrase("$" + String.Format("{0:n2}", double.Parse(reporte[1].ToString())), pe.FuenteParrafo)) { BorderWidth = .5f, Colspan = 1 };
                PdfPCell clDatosCantidad = new PdfPCell(new Phrase(reporte[2].ToString(), pe.FuenteParrafo)) { BorderWidth = .5f, Colspan = 1 };
                PrecioTotal += double.Parse(reporte[1].ToString().Replace("$", "").Replace(",", ""));
                CantidadTotal += int.Parse(reporte[2].ToString());
                tblDatosEquipos.AddCell(clDatoTipo);
                tblDatosEquipos.AddCell(clDatosTotal);
                tblDatosEquipos.AddCell(clDatosCantidad);
            }

            PdfPCell clVacio = new PdfPCell(new Phrase("", pe.FuenteParrafo)) { BorderWidth = .5f, Colspan = 1 };
            PdfPCell clSumaTotal = new PdfPCell(new Phrase("$" + String.Format("{0:n2}", PrecioTotal), pe.FuenteParrafoBold)) { BorderWidth = .5f, Colspan = 1 };
            PdfPCell clCantidadTotal = new PdfPCell(new Phrase(CantidadTotal.ToString(), pe.FuenteParrafoBold)) { BorderWidth = .5f, Colspan = 1 };
            tblDatosEquipos.AddCell(clVacio);
            tblDatosEquipos.AddCell(clSumaTotal);
            tblDatosEquipos.AddCell(clCantidadTotal);

            reporte.Close();
            document.Add(tblDatosEquipos);
        }

        public void ObtenerPreciosTotalesEquipos()
        {
            comando.Connection = conexion.AbrirConexion();
            comando.CommandText = "ObtenerTotalPreciosEquiposEnRenta";
            comando.CommandType = CommandType.StoredProcedure;
            comando.Parameters.Clear();
            comando.Parameters.AddWithValue("@Parametro", ParametroBusqueda);
            reporte = comando.ExecuteReader();
        }


    }
}
