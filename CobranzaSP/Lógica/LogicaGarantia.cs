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
using System.Runtime.Remoting.Messaging;

namespace CobranzaSP.Lógica
{
    internal class LogicaGarantia
    {
        private CD_Conexion conexion = new CD_Conexion();
        SqlCommand comando = new SqlCommand();
        SqlDataReader leer;
        PdfPTable tblGarantias;
        string TipoBusqueda = "";
        string ParametroBusqueda = "";
        string SegundoParametro = "";

        SortedSet<string> lstFechas = new SortedSet<string>();
        SortedSet<string> lstModelos = new SortedSet<string>();
        SortedSet<string> lstMarcas = new SortedSet<string>();
        SortedSet<string> lstClientes = new SortedSet<string>();

        Document document;

        public string RegistrarGarantia(Garantia NuevaGarantia, string sp)
        {
            int respuesta = 0;
            string AccionRealizada;
            comando.Connection = conexion.AbrirConexion();
            comando.CommandText = sp;
            comando.CommandType = CommandType.StoredProcedure;

            //Preguntamos que accion realizaremos en la base de datos para posteriormente mostrale al usuario la accion que realizo
            AccionRealizada = (sp == "AgregarGarantia") ? "agrego" : "modifico";

            if (NuevaGarantia.Id > 0)
            {
                comando.Parameters.AddWithValue("@IdGarantia", NuevaGarantia.Id);
            }
            comando.Parameters.AddWithValue("@IdMarca", NuevaGarantia.IdMarca);
            comando.Parameters.AddWithValue("@IdCartucho", NuevaGarantia.IdCartucho);
            comando.Parameters.AddWithValue("@IdCliente", NuevaGarantia.IdCliente);
            comando.Parameters.AddWithValue("@Fecha", NuevaGarantia.Fecha);
            comando.Parameters.AddWithValue("@Cantidad", NuevaGarantia.Cantidad);
            comando.Parameters.AddWithValue("@Descripcion", NuevaGarantia.Descripcion);

            respuesta = comando.ExecuteNonQuery();
            string Mensaje = (respuesta > 0) ? "Garantia se " + AccionRealizada + " correctamente" : "Algo salio mal, no se " + AccionRealizada + " la garantia";

            comando.Parameters.Clear();
            conexion.CerrarConexion();
            return Mensaje;
        }

        #region PDF

        #region ObtenerDatos
        public bool DeterminarTipoReporte(DatosReporteGarantia NuevoReporte)
        {
            bool DatosEncontrados = false;
            TipoBusqueda = NuevoReporte.TipoBusqueda;

            switch (TipoBusqueda)
            {
                case "Modelo":
                    DatosEncontrados = ObtenerDatosReporteModeloGarantia(NuevoReporte);
                    break;
                case "Cliente":
                    DatosEncontrados = ObtenerDatosReporteClienteModeloGarantia(NuevoReporte);
                    ; break;
                case "Rango Fecha":
                    DatosEncontrados = ObtenerDatosReporteModeloGarantia(NuevoReporte);
                    ; break;
            }
            return DatosEncontrados;
        }

        public bool ObtenerDatosReporteModeloGarantia(DatosReporteGarantia NuevoReporte)
        {
            bool DatosEncontrados = true;
            comando.Connection = conexion.AbrirConexion();
            comando.CommandText = "ObtenerDatosReporteModeloGarantia";
            comando.CommandType = CommandType.StoredProcedure;
            comando.Parameters.Clear();
            comando.Parameters.AddWithValue("@Cartucho", NuevoReporte.ParametroBusqueda);
            comando.Parameters.AddWithValue("@FechaInicial", NuevoReporte.FechaInicio);
            comando.Parameters.AddWithValue("@FechaFinal", NuevoReporte.FechaFinal);
            leer = comando.ExecuteReader();
            if (!leer.HasRows)
            {
                leer.Close();
                return DatosEncontrados = false;
            }
            Pdf(NuevoReporte);
            return DatosEncontrados;
        }

        public void ObtenerTotalesGarantia(DatosReporteGarantia NuevoReporte)
        {
            comando.Connection = conexion.AbrirConexion();
            comando.CommandText = "ObtenerTotalesGarantias";
            comando.CommandType = CommandType.StoredProcedure;
            comando.Parameters.Clear();
            switch (TipoBusqueda)
            {
                case "Modelo":
                    comando.Parameters.AddWithValue("@Cartucho", NuevoReporte.ParametroBusqueda);
                    comando.Parameters.AddWithValue("@Cliente", "");
                    break;
                case "Cliente":
                    comando.Parameters.AddWithValue("@Cartucho", NuevoReporte.SegundoParametro);
                    comando.Parameters.AddWithValue("@Cliente", NuevoReporte.ParametroBusqueda);
                    ; break;
                default:
                    comando.Parameters.AddWithValue("@Cartucho", "");
                    comando.Parameters.AddWithValue("@Cliente", "");
                    ; break;
            }
            comando.Parameters.AddWithValue("@FechaInicial", NuevoReporte.FechaInicio);
            comando.Parameters.AddWithValue("@FechaFinal", NuevoReporte.FechaFinal);
            leer = comando.ExecuteReader();

            if (!leer.HasRows)
            {
                bool Bandera = true;
            }
        }

        public bool ObtenerDatosReporteClienteModeloGarantia(DatosReporteGarantia NuevoReporte)
        {
            bool DatosEncontrados = true;
            comando.Connection = conexion.AbrirConexion();
            comando.CommandText = "ObtenerDatosReporteClienteModelosGarantia";
            comando.CommandType = CommandType.StoredProcedure;
            comando.Parameters.Clear();
            comando.Parameters.AddWithValue("@Cartucho", NuevoReporte.SegundoParametro);
            comando.Parameters.AddWithValue("@Cliente", NuevoReporte.ParametroBusqueda);
            comando.Parameters.AddWithValue("@FechaInicial", NuevoReporte.FechaInicio);
            comando.Parameters.AddWithValue("@FechaFinal", NuevoReporte.FechaFinal);
            leer = comando.ExecuteReader();
            if (!leer.HasRows)
            {
                leer.Close();
                return DatosEncontrados = false;
            }
            Pdf(NuevoReporte);
            return DatosEncontrados;
        }
        #endregion

        #region CreacionPdf
        public void Pdf(DatosReporteGarantia NuevoReporte)
        {
            string NombreArchivo = @"\\DESKTOP-D3OQEJR\Archivos Compartidos\Reportes\Garantias\" + "ReporteGarantia" + DateTime.Now.ToString("ddMMyyyyHHmmss") + ".pdf";
            //string NombreArchivo = @"\\administracion-pc\ARCHIVOS COMPARTIDOS\Reportes\Equipos\" + "ReporteEquipos" + DateTime.Now.ToString("ddMMyyyyHHmmss") + ".pdf";


            FileStream fs = new FileStream(NombreArchivo, FileMode.Create);
            document = new Document(PageSize.LETTER);
            document.SetMargins(25f, 25f, 25f, 25f);
            //Colocamos el pdf en horizontal
            document.SetPageSize(iTextSharp.text.PageSize.LETTER.Rotate());
            PdfWriter pw = PdfWriter.GetInstance(document, fs);

            //Instanciamos la clase para la paginacion
            var pe = new Pdf();
            pe.ColocarFormatoSuperior = true;
            pw.PageEvent = pe;

            document.Open();
            string NombreReporte = ColocarTituloReporte(NuevoReporte);
            //Colocar el titulo del reporte
            Paragraph titulo = new Paragraph(NombreReporte, pe.FuenteTitulo) { Alignment = Element.ALIGN_CENTER };
            document.Add(titulo);

            Paragraph Fechas = new Paragraph("DE: " + NuevoReporte.FechaInicio.ToString("dd/MM/yyyy") + "       AL: " + NuevoReporte.FechaFinal.ToString("dd/MM/yyyy"), pe.FuenteFecha) { Alignment = Element.ALIGN_CENTER };
            document.Add(Fechas);

            tblGarantias = new PdfPTable(4);

            switch (TipoBusqueda)
            {
                case "Modelo":
                    GenerarPdf();
                    break;
                case "Cliente":
                    GenerarPdf();

                    ; break;

                case "Rango Fecha":
                    GenerarPdf();
                    ; break;
            }


            document.Add(tblGarantias);

            lstClientes.Clear();
            lstMarcas.Clear();
            lstFechas.Clear();
            lstModelos.Clear();
            leer.Close();

            //Separamos con una linea los totales
            iTextSharp.text.pdf.draw.LineSeparator lineSeparator = new iTextSharp.text.pdf.draw.LineSeparator() { Offset = 2f };
            document.Add(new Chunk(lineSeparator));
            //TITULO DE LOS TOTALES
            Paragraph TituloTotales = new Paragraph("RESUMEN GARANTIAS DEL " + NuevoReporte.FechaInicio.ToString("dd/MM/yyyy") + " - " + NuevoReporte.FechaFinal.ToString("dd/MM/yyyy"), pe.FuenteTitulo) { Alignment = Element.ALIGN_CENTER };
            document.Add(TituloTotales);
            document.Add(new Chunk());

            ObtenerTotalesGarantia(NuevoReporte);
            CrearTablaTotales();
            document.Add(tblGarantias);
            leer.Close();
            document.Close();
            //Abrimos el pdf

            pe.AbrirPdf(NombreArchivo);
        }

        public void GenerarPdf()
        {
            Pdf pe = new Pdf();

            while (leer.Read())
            {
                ReporteGarantia ReporteGarantia = new ReporteGarantia()
                {
                    Marca = leer[0].ToString(),
                    Cartucho = leer[1].ToString(),
                    Fecha = DateTime.Parse(leer[2].ToString()),
                    Cliente = leer[3].ToString(),
                    Cantidad = int.Parse(leer[4].ToString()),
                    Descripcion = leer[5].ToString()
                };
                if (TipoBusqueda != "Cliente")
                {
                    AgregarMarca(ReporteGarantia.Marca, ReporteGarantia);
                }
                else
                {
                    AgregarCliente(ReporteGarantia.Cliente, ReporteGarantia);
                }
            }
        }

        public void AgregarModelo(string Modelo, ReporteGarantia ReporteGarantia)
        {
            Pdf pe = new Pdf();

            if (!lstModelos.Contains(Modelo))
            {
                document.Add(tblGarantias);//Requerimos mover esta linea a otra
                if (TipoBusqueda != "Modelo")
                {
                    Paragraph ParrafoModelo;
                    ParrafoModelo = new Paragraph(Modelo, pe.FuenteParrafoGrande) { Alignment = Element.ALIGN_LEFT };
                    //Agregamos la fecha tanto a la lista que tenemos de fechas, como al documento colocando la cantidad que tenemos en dicho día
                    document.Add(ParrafoModelo);
                }

                lstModelos.Add(Modelo);
                CrearTablaGarantia();
                AgregarGarantiaATabla(ReporteGarantia);
            }
            else
            {
                AgregarGarantiaATabla(ReporteGarantia);
            }
        }

        public void CrearTablaGarantia()
        {
            Pdf pe = new Pdf();
            tblGarantias = (TipoBusqueda != "Cliente") ? new PdfPTable(5) : new PdfPTable(3);

            if (TipoBusqueda != "Cliente")
            {
                PdfPCell clCliente = new PdfPCell(new Phrase("CLIENTE", pe.FuenteParrafoBold)) { BorderWidth = .5f, Colspan = 2 };
                tblGarantias.AddCell(clCliente);
            }
            else
            {
            }
            PdfPCell clFecha = new PdfPCell(new Phrase("FECHA", pe.FuenteParrafoBold)) { BorderWidth = .5f, Colspan = 1 };
            tblGarantias.AddCell(clFecha);
            PdfPCell clCantidad = new PdfPCell(new Phrase("CANTIDAD", pe.FuenteParrafoBold)) { BorderWidth = .5f, Colspan = 1 };
            tblGarantias.AddCell(clCantidad);
            PdfPCell clDescripcion = new PdfPCell(new Phrase("DESCRIPCION", pe.FuenteParrafoBold)) { BorderWidth = .5f, Colspan = 1 };
            tblGarantias.AddCell(clDescripcion);
        }

        public void AgregarGarantiaATabla(ReporteGarantia ReporteGarantia)
        {
            Pdf pe = new Pdf();

            if (TipoBusqueda != "Cliente")
            {
                PdfPCell clCliente = new PdfPCell(new Phrase(ReporteGarantia.Cliente, pe.FuenteParrafo)) { BorderWidth = .5f, Colspan = 2 };
                tblGarantias.AddCell(clCliente);
            }
            else
            {
            }
            PdfPCell clFecha = new PdfPCell(new Phrase(ReporteGarantia.Fecha.ToString("dd/MM/yyyy"), pe.FuenteParrafo)) { BorderWidth = .5f, Colspan = 1 };
            tblGarantias.AddCell(clFecha);
            PdfPCell clCantidad = new PdfPCell(new Phrase(ReporteGarantia.Cantidad.ToString(), pe.FuenteParrafo)) { BorderWidth = .5f, Colspan = 1 };
            tblGarantias.AddCell(clCantidad);
            PdfPCell clDescripcion = new PdfPCell(new Phrase(ReporteGarantia.Descripcion, pe.FuenteParrafo)) { BorderWidth = .5f, Colspan = 1 };
            tblGarantias.AddCell(clDescripcion);
        }

        public void AgregarMarca(string Marca, ReporteGarantia ReporteGarantia)
        {
            Pdf pe = new Pdf();

            if (!lstMarcas.Contains(Marca))
            {
                //Si tenemos dicha fecha en la lista, agregamos los datos que tenemos de la anterior fecha
                document.Add(tblGarantias);//Requerimos mover esta linea a otra
                //lstMarcas.Clear();
                lstModelos.Clear();
                tblGarantias = new PdfPTable(4);
                Paragraph ParrafoMarca;
                if (TipoBusqueda == "Rango Fecha")
                {
                    ParrafoMarca = new Paragraph(Marca, pe.FuenteParrafoGrande) { Alignment = Element.ALIGN_LEFT };
                    //Agregamos la fecha tanto a la lista que tenemos de fechas, como al documento colocando la cantidad que tenemos en dicho día
                    document.Add(ParrafoMarca);
                }
                lstMarcas.Add(Marca);

                AgregarModelo(ReporteGarantia.Cartucho, ReporteGarantia);
            }
            else
            {
                AgregarModelo(ReporteGarantia.Cartucho, ReporteGarantia);
            }
        }

        public void AgregarCliente(string Cliente, ReporteGarantia ReporteGarantia)
        {
            Pdf pe = new Pdf();

            if (!lstClientes.Contains(Cliente))
            {

                document.Add(tblGarantias);//Requerimos mover esta linea a otra
                iTextSharp.text.pdf.draw.LineSeparator lineSeparator = new iTextSharp.text.pdf.draw.LineSeparator() { Offset = 2f };
                document.Add(new Chunk(lineSeparator));
                //lstMarcas.Clear();
                tblGarantias = new PdfPTable(4);
                lstMarcas.Clear();
                if (TipoBusqueda != "Cliente")
                {
                    Paragraph ParrafoCliente;
                    ParrafoCliente = new Paragraph(Cliente, pe.FuenteParrafoGrande) { Alignment = Element.ALIGN_LEFT };
                    //Agregamos la fecha tanto a la lista que tenemos de fechas, como al documento colocando la cantidad que tenemos en dicho día
                    document.Add(ParrafoCliente);
                }
                lstClientes.Add(Cliente);
                AgregarMarca(ReporteGarantia.Marca, ReporteGarantia);
            }
            else
            {
                AgregarMarca(ReporteGarantia.Marca, ReporteGarantia);
            }
        }



        public string ColocarTituloReporte(DatosReporteGarantia NuevoReporte)
        {
            string Titulo = "GARANTIAS DE ";
            switch (TipoBusqueda)
            {
                case "Cliente":
                    Titulo += (NuevoReporte.ParametroBusqueda == "") ? " CLIENTES" : "CLIENTE:" + NuevoReporte.ParametroBusqueda;
                    break;
                case "Modelo":
                    Titulo += NuevoReporte.Marca + " " + NuevoReporte.ParametroBusqueda;
                    break;
                case "Rango Fecha": Titulo = "REPORTE DE GARANTIAS"; break;
            }
            return Titulo;
        }

        public void CrearTablaTotales()
        {
            Pdf pe = new Pdf();
            tblGarantias = new PdfPTable(3);

            PdfPCell clMarca = new PdfPCell(new Phrase("MARCA", pe.FuenteParrafoBold)) { BorderWidth = .5f, Colspan = 1 };
            tblGarantias.AddCell(clMarca);
            PdfPCell clModelo = new PdfPCell(new Phrase("MODELO", pe.FuenteParrafoBold)) { BorderWidth = .5f, Colspan = 1 };
            tblGarantias.AddCell(clModelo);
            PdfPCell clTotal = new PdfPCell(new Phrase("TOTAL", pe.FuenteParrafoBold)) { BorderWidth = .5f, Colspan = 1 };
            tblGarantias.AddCell(clTotal);
            while (leer.Read())
            {
                PdfPCell clDatoMarca = new PdfPCell(new Phrase(leer[0].ToString(), pe.FuenteParrafoBold)) { BorderWidth = .5f, Colspan = 1 };
                tblGarantias.AddCell(clDatoMarca);
                PdfPCell clDatoModelo = new PdfPCell(new Phrase(leer[1].ToString(), pe.FuenteParrafoBold)) { BorderWidth = .5f, Colspan = 1 };
                tblGarantias.AddCell(clDatoModelo);
                PdfPCell clDatoTotal = new PdfPCell(new Phrase(leer[2].ToString(), pe.FuenteParrafoBold)) { BorderWidth = .5f, Colspan = 1 };
                tblGarantias.AddCell(clDatoTotal);
            }

        }

        #endregion


        #endregion
    }
}
