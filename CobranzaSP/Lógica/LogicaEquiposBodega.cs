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

namespace CobranzaSP.Lógica
{
    internal class LogicaEquiposBodega
    {
        private CD_Conexion conexion = new CD_Conexion();
        SqlCommand comando = new SqlCommand();
        SqlDataReader reporte;
        string TipoBusqueda = "";
        string ParametroBusqueda = "";
        SortedSet<string> lstMarcas = new SortedSet<string>();
        PdfPTable tblModelos;
        
        public string GuardarEquipo(EquipoBodega nuevoEquipo, string sp)
        {
            string AccionRealizada = "agrego";
            int Resultado = 0;
            comando.Connection = conexion.AbrirConexion();
            comando.CommandText = sp;
            comando.CommandType = CommandType.StoredProcedure;
            comando.Parameters.Clear();


            if (sp == "ModificarEquipoBodega")
            {
                AccionRealizada = "modifico";
                comando.Parameters.AddWithValue("@IdEquipo", nuevoEquipo.IdEquipo);
            }

            comando.Parameters.AddWithValue("@IdMarca", nuevoEquipo.IdMarca);
            comando.Parameters.AddWithValue("@IdModelo", nuevoEquipo.IdModelo);
            comando.Parameters.AddWithValue("@Serie", nuevoEquipo.Serie);
            comando.Parameters.AddWithValue("@Precio", nuevoEquipo.Precio);
            comando.Parameters.AddWithValue("@Estado", nuevoEquipo.Estado);
            comando.Parameters.AddWithValue("@Notas", nuevoEquipo.Notas);

            Resultado = comando.ExecuteNonQuery();
            string Mensaje = (Resultado > 0) ? "Equipo se " + AccionRealizada + " correctamente" : "Algo salio mal, no se " + AccionRealizada + " el registro";
            conexion.CerrarConexion();
            return Mensaje;
        }

        public bool ObtenerDatosTodasLasMarcas(string TipoBusquedaElegida, string Parametro)
        {
            bool DatosObtenidos = true;
            TipoBusqueda = TipoBusquedaElegida;
            ParametroBusqueda = Parametro;

            comando.Connection = conexion.AbrirConexion();
            comando.CommandText = "ReporteEquiposBodegaPorMarca";
            comando.CommandType = CommandType.StoredProcedure;
            comando.Parameters.Clear();
            reporte = comando.ExecuteReader();
            if (!reporte.HasRows)
            {
                reporte.Close();
                return DatosObtenidos = false;
            }
            GenerarPdf();
            return DatosObtenidos;
        }

        public bool ObtenerDatosMarcaElegida(string TipoBusquedaElegida, string Parametro, int Id)
        {
            bool DatosObtenidos = true;
            TipoBusqueda = TipoBusquedaElegida;
            ParametroBusqueda = Parametro;

            comando.Connection = conexion.AbrirConexion();
            comando.CommandText = "ReporteEquiposBodegaPorMarcaEspecifica";
            comando.CommandType = CommandType.StoredProcedure;
            comando.Parameters.Clear();
            comando.Parameters.AddWithValue("@IdMarca", Id);
            reporte = comando.ExecuteReader();
            if (!reporte.HasRows)
            {
                reporte.Close();
                return DatosObtenidos = false;
            }
            GenerarPdf();
            return DatosObtenidos;
        }

        public bool ObtenerDatosModeloElegida(string TipoBusquedaElegida, string Parametro, int Id)
        {
            bool DatosObtenidos = true;
            TipoBusqueda = TipoBusquedaElegida;
            ParametroBusqueda = Parametro;

            comando.Connection = conexion.AbrirConexion();
            comando.CommandText = "ReporteEquiposBodegaPorModelo";
            comando.CommandType = CommandType.StoredProcedure;
            comando.Parameters.Clear();
            comando.Parameters.AddWithValue("@IdModelo", Id);
            reporte = comando.ExecuteReader();
            if (!reporte.HasRows)
            {
                reporte.Close();
                return DatosObtenidos = false;
            }
            GenerarPdf();
            return DatosObtenidos;
        }

        public void GenerarPdf()
        {
            string NombreArchivo = @"\\DESKTOP-D3OQEJR\Archivos Compartidos\Reportes\EquiposBodega\" + "ReporteEquipos" + DateTime.Now.ToString("ddMMyyyyHHmmss") + ".pdf";
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
            tblModelos = (TipoBusqueda == "Modelo") ?new PdfPTable(5): new PdfPTable(7);

            //Colocar el titulo del reporte
            string NombreReporte = ColocarNombreReporte();
            Paragraph titulo = new Paragraph(NombreReporte, pe.FuenteTitulo) { Alignment = Element.ALIGN_CENTER };
            document.Add(titulo);
            document.Add(new Chunk("\n"));

            while (reporte.Read())
            {
                ReporteEquipoBodega equipo = new ReporteEquipoBodega()
                {
                    Marca = reporte[0].ToString(),
                    Modelo = reporte[1].ToString(),
                    Serie = reporte[2].ToString(),
                    Precio = reporte[3].ToString(),
                    Estado = reporte[4].ToString(),
                    Notas = reporte[5].ToString()
                };

                ColocarSerieADocumento(equipo, document);
            }
            document.Add(tblModelos);
            lstMarcas.Clear();
            reporte.Close();

            iTextSharp.text.pdf.draw.LineSeparator lineSeparator = new iTextSharp.text.pdf.draw.LineSeparator();
            lineSeparator.Offset = 1f;
            document.Add(new iTextSharp.text.Chunk(lineSeparator));
            //Ahora deberemos de obtener la cantidad total en dinero de los equipos
            ColocarTotales(document);

            document.Close();


            //Abrimos el pdf 
            pe.AbrirPdf(NombreArchivo);
        }

        public string ColocarNombreReporte()
        {
            string NombreReporte = "";
            switch (TipoBusqueda)
            {
                case "Marca": NombreReporte = "REPORTE EQUIPOS EN BODEGA POR MARCA: " + ParametroBusqueda.ToUpper(); break;
                case "Modelo": NombreReporte = "REPORTE EQUIPOS EN BODEGA POR MODELO: " + ParametroBusqueda.ToUpper(); break;
                default: NombreReporte = "REPORTE EQUIPOS EN BODEGA"; break;
            }
            return NombreReporte;
        }

        public void ColocarSerieADocumento(ReporteEquipoBodega equipo, Document document)
        {
            var pe = new Pdf();

            if (!lstMarcas.Contains(equipo.Marca))
            {
                document.Add(tblModelos);
                tblModelos = (TipoBusqueda == "Modelo") ? new PdfPTable(5) : new PdfPTable(7);

                //Agregamos la marca al documento y a la lista
                lstMarcas.Add(equipo.Marca);
                if(TipoBusqueda != "Marca")
                {
                    Paragraph Marca = new Paragraph(equipo.Marca, pe.FuenteParrafoGrande) { Alignment = Element.ALIGN_LEFT };
                    document.Add(Marca);
                }
                ColocarTitulosTabla();
                AgregarModelosATabla(equipo);
            }
            else
            {
                AgregarModelosATabla(equipo);
            }
        }

        public void ColocarTitulosTabla()
        {
            var pe = new Pdf();
            if(TipoBusqueda != "Modelo")
            {
                PdfPCell clModelo = new PdfPCell(new Phrase("MODELO", pe.FuenteParrafoBold)) { BorderWidth = .5f, Colspan = 2 };
                tblModelos.AddCell(clModelo);
            }

            PdfPCell clSerie = new PdfPCell(new Phrase("SERIE",pe.FuenteParrafoBold)) { BorderWidth = .5f, Colspan = 2 };
            PdfPCell clPrecio = new PdfPCell(new Phrase("PRECIO", pe.FuenteParrafoBold)) { BorderWidth = .5f, Colspan = 1 };
            PdfPCell clEstado = new PdfPCell(new Phrase("ESTADO", pe.FuenteParrafoBold)) { BorderWidth = .5f, Colspan = 1 };
            PdfPCell clNotas = new PdfPCell(new Phrase("NOTAS", pe.FuenteParrafoBold)) { BorderWidth = .5f, Colspan = 1 };

            tblModelos.AddCell(clSerie);
            tblModelos.AddCell(clPrecio);
            tblModelos.AddCell(clEstado);
            tblModelos.AddCell(clNotas);
        }

        public void AgregarModelosATabla(ReporteEquipoBodega equipo)
        {
            var pe = new Pdf();

            if(TipoBusqueda != "Modelo")
            {
                PdfPCell clModelo = new PdfPCell(new Phrase(equipo.Modelo, pe.FuenteParrafo)) { BorderWidth = .5f, Colspan = 2 };
                tblModelos.AddCell(clModelo);
            }

            PdfPCell clSerie = new PdfPCell(new Phrase(equipo.Serie, pe.FuenteParrafo)) { BorderWidth = .5f, Colspan = 2 };
            PdfPCell clPrecio = new PdfPCell(new Phrase(String.Format("{0:n0}", equipo.Precio), pe.FuenteParrafo)) { BorderWidth = .5f, Colspan = 1 };
            PdfPCell clEstado = new PdfPCell(new Phrase(equipo.Estado, pe.FuenteParrafo)) { BorderWidth = .5f, Colspan = 1 };
            PdfPCell clNotas = new PdfPCell(new Phrase(equipo.Notas, pe.FuenteParrafo)) { BorderWidth = .5f, Colspan = 1 };


            tblModelos.AddCell(clSerie);
            tblModelos.AddCell(clPrecio);
            tblModelos.AddCell(clEstado);
            tblModelos.AddCell(clNotas);
        }

        public void ColocarTotales(Document document)
        {
            var pe = new Pdf();
            int CantidadTotal = 0;
            double PrecioTotal = 0;

            Paragraph titulo = new Paragraph("TOTAL", pe.FuenteTitulo) { Alignment = Element.ALIGN_CENTER };
            document.Add(titulo);
            //document.Add(new Chunk("\n"));

            tblModelos = new PdfPTable(3);
            PdfPCell clTipo = new PdfPCell();
            clTipo = new PdfPCell(new Phrase("MARCA", pe.FuenteParrafoBold)) { BorderWidth = .5f, Colspan = 1 };

            if(TipoBusqueda == "Modelo")
            {
                clTipo = new PdfPCell(new Phrase("MODELO", pe.FuenteParrafoBold)) { BorderWidth = .5f, Colspan = 1 };
            }

            PdfPCell clPrecio = new PdfPCell(new Phrase("PRECIO", pe.FuenteParrafoBold)) { BorderWidth = .5f, Colspan = 1};
            PdfPCell clCantidad = new PdfPCell(new Phrase("CANTIDAD", pe.FuenteParrafoBold)) { BorderWidth = .5f, Colspan = 1 };
            tblModelos.AddCell(clTipo);
            tblModelos.AddCell(clPrecio);
            tblModelos.AddCell(clCantidad);

            ObtenerPreciosTotalesEquipos();

            while (reporte.Read())
            {
                PdfPCell clDatoTipo = (TipoBusqueda == "Modelo") ? new PdfPCell(new Phrase(ParametroBusqueda, pe.FuenteParrafo)) { BorderWidth = .5f, Colspan = 1 } : new PdfPCell(new Phrase(reporte[0].ToString(), pe.FuenteParrafo)) { BorderWidth = .5f, Colspan = 1 };
                PdfPCell clDatosTotal = new PdfPCell(new Phrase(reporte[1].ToString(), pe.FuenteParrafo)) { BorderWidth = .5f, Colspan = 1 };
                PdfPCell clDatosCantidad = new PdfPCell(new Phrase(reporte[2].ToString(), pe.FuenteParrafo)) { BorderWidth = .5f, Colspan = 1 };
                PrecioTotal += double.Parse(reporte[1].ToString().Replace("$", "").Replace(",",""));
                CantidadTotal += int.Parse(reporte[2].ToString());
                tblModelos.AddCell(clDatoTipo);
                tblModelos.AddCell(clDatosTotal);
                tblModelos.AddCell(clDatosCantidad);
            }

            PdfPCell clVacio = new PdfPCell(new Phrase("", pe.FuenteParrafo)) { BorderWidth = .5f, Colspan = 1 };
            PdfPCell clSumaTotal = new PdfPCell(new Phrase("$"+String.Format("{0:n2}", PrecioTotal), pe.FuenteParrafoBold)) { BorderWidth = .5f, Colspan = 1 };
            PdfPCell clCantidadTotal = new PdfPCell(new Phrase(CantidadTotal.ToString(), pe.FuenteParrafoBold)) { BorderWidth = .5f, Colspan = 1 };
            tblModelos.AddCell(clVacio);
            tblModelos.AddCell(clSumaTotal);
            tblModelos.AddCell(clCantidadTotal);


            reporte.Close();
            document.Add(tblModelos);
        }

        public void ObtenerPreciosTotalesEquipos()
        {
            comando.Connection = conexion.AbrirConexion();
            comando.CommandText = "ObtenerTotalPreciosEquipos";
            comando.CommandType = CommandType.StoredProcedure;
            comando.Parameters.Clear();
            comando.Parameters.AddWithValue("@Parametro", ParametroBusqueda);
            reporte = comando.ExecuteReader();
        }
    }
}
