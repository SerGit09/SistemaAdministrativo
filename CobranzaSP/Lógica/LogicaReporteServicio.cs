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

namespace CobranzaSP.Lógica
{
    internal class LogicaReporteServicio
    {
        private CD_Conexion conexion = new CD_Conexion();
        private AccionesLógica NuevaAccion = new AccionesLógica();
        SqlCommand comando = new SqlCommand();
        SqlDataReader reporte;
        SortedSet<string> lstSeries = new SortedSet<string>();
        SortedSet<string> lstClientes = new SortedSet<string>();
        SortedSet<string> lstFechas = new SortedSet<string>();
        PdfPTable tblSeries; 
        List<string> Fusores = new List<string>();
        bool RangoHabilitado = false;
        string TipoBusqueda;
        string Cliente;

        public bool DeterminarTipoReporte(Reporte nuevoReporte,string Cliente)
        {
            bool DatosObtenidos = true;
            this.Cliente = Cliente;
            switch (nuevoReporte.TipoBusqueda.ToUpper())
            {
                case "SERIE": DatosObtenidos = DeterminarGenerarReporteConRangoFecha(nuevoReporte); break;
                case "FUSOR": DatosObtenidos = ObtenerDatosReporteFusor(nuevoReporte); break;
                case "FUSORES": DatosObtenidos = ObtenerFusoresReporte(nuevoReporte); break;
                case "CLIENTE": DatosObtenidos = ReporteCliente(nuevoReporte, "GenerarReporteCliente"); break;
                case "MARCA": DatosObtenidos = ReporteCliente(nuevoReporte, "GenerarReporteCliente"); ; break;
                case "MODELO": DatosObtenidos = ReporteCliente(nuevoReporte, "GenerarReporteCliente"); ; break;
                case "FECHA": DatosObtenidos = ObtenerDatosReporteRangoDeFecha(nuevoReporte); break;
                case "TODOS LOS FUSORES": ObtenerFusoresReporte(nuevoReporte); break;
            }

            return DatosObtenidos;
        }

        public bool ReporteCliente(Reporte nuevoReporte, string sp)
        {
            //Posiblemente se regrese
            if(Cliente != nuevoReporte.ParametroBusqueda)
            {
                //nuevoReporte.ParametroBusqueda = "";
                return ObtenerDatosReporteCliente(nuevoReporte, sp);
            }
            else
            {
                return DeterminarGenerarReporteConRangoFecha(nuevoReporte);
            }
        }

        public bool DeterminarGenerarReporteConRangoFecha(Reporte nuevoReporte)
        {
            if (nuevoReporte.RangoHabilitado)
            {
                RangoHabilitado = nuevoReporte.RangoHabilitado;
                return ObtenerDatosReporteRangoDeFecha(nuevoReporte);
            }
            else
                return ObtenerDatosReporteParametroEspecifico(nuevoReporte);
        }

        public bool DeterminarGenerarReporteConRangoFechaRicoh(Reporte nuevoReporte)
        {
            if (nuevoReporte.RangoHabilitado)
            {
                RangoHabilitado = nuevoReporte.RangoHabilitado;
                return ObtenerDatosReporteRangoDeFechaRicoh(nuevoReporte);
            }
            else
                return ObtenerDatosReporteParametroEspecificoRicoh(nuevoReporte);
        }

        #region ObtenerDatos
        //Obtenemos datos de un parametro en especifico sin especificar un rango de fecha
        public bool ObtenerDatosReporteParametroEspecifico(Reporte NuevoReporte)
        {
            bool DatosObtenidos = true;
            comando.Connection = conexion.AbrirConexion();
            comando.CommandText = "DatosReporteServicio";
            comando.CommandType = CommandType.StoredProcedure;
            comando.Parameters.Clear();
            comando.Parameters.AddWithValue("@CampoBusqueda", NuevoReporte.ParametroBusqueda);
            reporte = comando.ExecuteReader();
            if (!reporte.HasRows)
            {
                reporte.Close();
                return DatosObtenidos = false;
            }

            //Comenzamos a generar el reporte
            Pdf(NuevoReporte);
            return DatosObtenidos;
        }

        public bool ObtenerDatosReporteParametroEspecificoRicoh(Reporte NuevoReporte)
        {
            bool DatosObtenidos = true;
            comando.Connection = conexion.AbrirConexion();
            comando.CommandText = "DatosReporteServicioRicoh";
            comando.CommandType = CommandType.StoredProcedure;
            comando.Parameters.Clear();
            comando.Parameters.AddWithValue("@CampoBusqueda", NuevoReporte.ParametroBusqueda);
            reporte = comando.ExecuteReader();
            if (!reporte.HasRows)
            {
                reporte.Close();
                return DatosObtenidos = false;
            }

            //Comenzamos a generar el reporte
            Pdf(NuevoReporte);
            return DatosObtenidos;
        }

        public bool ObtenerDatosReporteCliente(Reporte NuevoReporte, string sp)
        {
            bool DatosObtenidos = true;
            comando.Connection = conexion.AbrirConexion();
            //comando.CommandText = "GenerarReporteCliente";
            comando.CommandText = sp;
            comando.Parameters.Clear();
            comando.Parameters.AddWithValue("@FechaInicio", NuevoReporte.FechaInicio);
            comando.Parameters.AddWithValue("@FechaFinal", NuevoReporte.FechaFinal);
            comando.Parameters.AddWithValue("@CampoBusqueda", NuevoReporte.ParametroBusqueda);
            comando.Parameters.AddWithValue("@Cliente", Cliente);
            comando.Parameters.AddWithValue("@RangoHabilitado", NuevoReporte.RangoHabilitado);
            comando.CommandType = CommandType.StoredProcedure;
            reporte = comando.ExecuteReader();
            if (!reporte.HasRows)
            {
                return DatosObtenidos = false;
            }

            //Comenzamos a generar el pdf
            Pdf(NuevoReporte);

            return DatosObtenidos;
        }

        public bool ObtenerFusoresReporte(Reporte NuevoReporte)
        {
            bool DatosObtenidos = true;
            comando.Connection = conexion.AbrirConexion();
            comando.CommandText = "GenerarReporteFusoresServicios";
            comando.Parameters.Clear();
            comando.Parameters.AddWithValue("@FechaInicio", NuevoReporte.FechaInicio);
            comando.Parameters.AddWithValue("@FechaFinal", NuevoReporte.FechaFinal);
            comando.Parameters.AddWithValue("@RangoHabilitado", NuevoReporte.RangoHabilitado);
            comando.CommandType = CommandType.StoredProcedure;
            reporte = comando.ExecuteReader();
            if (!reporte.HasRows)
            {
                return DatosObtenidos = false;
            }

            //while (reporte.Read())
            //{
            //    Fusores.Add(reporte[0].ToString());
            //}
            //reporte.Close();

            //Comenzamos a generar el pdf
            Pdf(NuevoReporte);

            return DatosObtenidos;
        }

        //Reporte unicamente para servicios ricoh
        public bool ObtenerDatosReporteRangoDeFechaRicoh(Reporte NuevoReporte)
        {
            bool DatosObtenidos = true;
            comando.Connection = conexion.AbrirConexion();
            comando.CommandText = "GenerarReporteEspecificoServicioRicoh";
            comando.CommandType = CommandType.StoredProcedure;
            comando.Parameters.Clear();
            comando.Parameters.AddWithValue("@FechaInicio", NuevoReporte.FechaInicio);
            comando.Parameters.AddWithValue("@FechaFinal", NuevoReporte.FechaFinal);
            comando.Parameters.AddWithValue("@CampoBusqueda", NuevoReporte.ParametroBusqueda);
            reporte = comando.ExecuteReader();
            if (!reporte.HasRows)
            {
                reporte.Close();
                return DatosObtenidos = false;
            }

            //Comenzamos a generar el reporte
            Pdf(NuevoReporte);
            return DatosObtenidos;
        }

        public bool ObtenerDatosReporteRangoDeFecha(Reporte NuevoReporte)
        {
            bool DatosObtenidos = true;
            comando.Connection = conexion.AbrirConexion();
            comando.CommandText = "GenerarReporteRangoFecha";
            comando.CommandType = CommandType.StoredProcedure;
            comando.Parameters.Clear();
            comando.Parameters.AddWithValue("@FechaInicio", NuevoReporte.FechaInicio);
            comando.Parameters.AddWithValue("@FechaFinal", NuevoReporte.FechaFinal);
            comando.Parameters.AddWithValue("@CampoBusqueda", NuevoReporte.ParametroBusqueda);
            reporte = comando.ExecuteReader();
            if (!reporte.HasRows)
            {
                reporte.Close();
                return DatosObtenidos = false;
            }
            Pdf(NuevoReporte);
            return DatosObtenidos;
        }

        public bool ObtenerDatosReporteFusor(Reporte NuevoReporte)
        {
            bool DatosObtenidos = true;
            comando.Connection = conexion.AbrirConexion();
            comando.CommandText = "ReporteFusorServicios";
            comando.CommandType = CommandType.StoredProcedure;
            comando.Parameters.Clear();
            comando.Parameters.AddWithValue("@Fusor", NuevoReporte.ParametroBusqueda);
            reporte = comando.ExecuteReader();
            if (!reporte.HasRows)
            {
                reporte.Close();
                return DatosObtenidos = false;
            }
            Pdf(NuevoReporte);
            return DatosObtenidos;

        }
        #endregion

        //Nos ayuda a determinar si el reporte estara de acuerdo a un rango de fecha o no

        public void Pdf(Reporte NuevoReporte)
        {
            string NombreArchivo = @"\\DESKTOP-D3OQEJR\Archivos Compartidos\Reportes\Servicios\" + NuevoReporte.TipoBusqueda + DateTime.Now.ToString("ddMMyyyyHHmmss") + ".pdf";

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

            if (!string.IsNullOrEmpty(NuevoReporte.TipoBusquedaAdicional) && NuevoReporte.TipoBusqueda != "CLIENTE")
            {
                Paragraph ParrafoSubtitulo = new Paragraph("POR " + NuevoReporte.TipoBusquedaAdicional + " : " + Cliente, pe.FuenteTitulo) { Alignment = Element.ALIGN_CENTER };
                document.Add(ParrafoSubtitulo);
            }

            //Colocar el titulo del reporte
            string NombreReporte = ColocarNombreReporte(NuevoReporte);
            Paragraph titulo = new Paragraph(NombreReporte, pe.FuenteTitulo) { Alignment = Element.ALIGN_CENTER };
            document.Add(titulo);

            //Verificamos si tenemos rango de fecha
            if (NuevoReporte.RangoHabilitado || NuevoReporte.TipoBusqueda == "FECHA")
            {
                Paragraph Fechas = new Paragraph("DEL " + NuevoReporte.FechaInicio.ToString("dd/MM/yyyy") + " AL " + NuevoReporte.FechaFinal.ToString("dd/MM/yyyy"), pe.FuenteFecha) { Alignment = Element.ALIGN_CENTER };
                document.Add(Fechas);
            }

            //A partir de aqui va a variar el contenido de los pdfs, dependiendo cual se eligio
            switch (NuevoReporte.TipoBusqueda.ToUpper())
            {
                case "FUSORES":GenerarReporteFusor(document); break;
                case "FUSOR": GenerarReporteFusor(document); break;
                default: GenerarReporte(document, NuevoReporte); break;
            }

            lstClientes.Clear();
            lstSeries.Clear();
            reporte.Close();
            document.Close();

            //Abrimos el pdf 
            pe.AbrirPdf(NombreArchivo);
        }

        public string ColocarNombreReporte(Reporte NuevoReporte)
        {
            string NombreReporte = "REPORTE SERVICIO TECNICO";
            switch (NuevoReporte.TipoBusqueda.ToUpper())
            {
                case "CLIENTE": NombreReporte += " POR CLIENTE"; break;
                case "SERIE": NombreReporte += " POR SERIE: " + NuevoReporte.ParametroBusqueda.ToUpper(); break;
                case "FUSORES": NombreReporte += " FUSORES"; break;
                case "FUSOR": NombreReporte += " POR FUSOR: " + NuevoReporte.ParametroBusqueda; break;
                case "MARCA": NombreReporte += " POR MARCA: " + NuevoReporte.ParametroBusqueda; break;
                case "MODELO": NombreReporte += "POR MODELO: " + NuevoReporte.ParametroBusqueda; break;
            }
            return NombreReporte;
        }

        public void VerificarReportePorCliente(Reporte NuevoReporte)
        {
            
        }

        public void GenerarReporteFusor(Document document)
        {
            var pe = new Pdf();
            DateTime Fecha;

            PdfPTable tblFusor = new PdfPTable(6) { WidthPercentage = 100 };

            document.Add(new Paragraph("\n"));

            PdfPCell clSerieFusor = new PdfPCell(new Phrase("Serie Fusor", pe.FuenteParrafoBold)) { BorderWidth = .5f };
            PdfPCell clModelo = new PdfPCell(new Phrase("Modelo", pe.FuenteParrafoBold)) { BorderWidth = .5f };
            PdfPCell clInstalado = new PdfPCell(new Phrase("Estado", pe.FuenteParrafoBold)) { BorderWidth = .5f };
            PdfPCell clFecha = new PdfPCell(new Phrase("Fecha", pe.FuenteParrafoBold)) { BorderWidth = .5f };
            PdfPCell clFolio = new PdfPCell(new Phrase("Folio Reporte", pe.FuenteParrafoBold)) { BorderWidth = .5f };
            PdfPCell clCliente = new PdfPCell(new Phrase("Ubicación", pe.FuenteParrafoBold)) { BorderWidth = .5f };

            tblFusor.AddCell(clFolio);
            tblFusor.AddCell(clCliente);
            tblFusor.AddCell(clFecha);
            tblFusor.AddCell(clModelo);
            tblFusor.AddCell(clSerieFusor);
            tblFusor.AddCell(clInstalado);

            tblFusor.HorizontalAlignment = Element.ALIGN_LEFT;
            //document.Add(new Paragraph(ParametroBusqueda, pe.FuenteParrafoBold));

            while (reporte.Read())
            {
                PdfPCell cSerieFusor = new PdfPCell(new Phrase(reporte[0].ToString(), pe.FuenteParrafo)) { BorderWidth = .5f, Padding = 2 };
                PdfPCell cModelo = new PdfPCell(new Phrase(reporte[1].ToString(), pe.FuenteParrafo)) { BorderWidth = .5f, Padding = 2, Rowspan = 2 };
                PdfPCell cInstalado = new PdfPCell(new Phrase("Instalado", pe.FuenteParrafo)) { BorderWidth = .5f, Padding = 2 };
                Fecha = Convert.ToDateTime(reporte[2].ToString());
                PdfPCell cFecha = new PdfPCell(new Phrase(Fecha.ToString("dd/MM/yyyy"), pe.FuenteParrafo)) { BorderWidth = .5f, Padding = 2, Rowspan = 2 };

                PdfPCell cFolio = new PdfPCell(new Phrase(reporte[3].ToString(), pe.FuenteParrafo)) { BorderWidth = .5f, Padding = 2, Rowspan = 2 };
                PdfPCell cCliente = new PdfPCell(new Phrase(reporte[4].ToString(), pe.FuenteParrafo)) { BorderWidth = .5f, Padding = 2, Rowspan = 2 };

                PdfPCell cSerieSaliente = new PdfPCell(new Phrase(reporte[5].ToString(), pe.FuenteParrafo)) { BorderWidth = .5f, Padding = 2 };
                PdfPCell cInstaladoR = new PdfPCell(new Phrase("Retirado", pe.FuenteParrafo)) { BorderWidth = .5f, Padding = 2 };

                tblFusor.AddCell(cFolio);
                tblFusor.AddCell(cCliente);
                tblFusor.AddCell(cFecha);
                tblFusor.AddCell(cModelo);
                tblFusor.AddCell(cSerieFusor);
                tblFusor.AddCell(cInstalado);

                tblFusor.AddCell(cSerieSaliente);
                tblFusor.AddCell(cInstaladoR);
            }
            document.Add(tblFusor);
        }

        public void GenerarReporteFusores(Document document)
        {
            var pe = new Pdf();
            DateTime Fecha;

            PdfPTable tblCliente = new PdfPTable(6) { WidthPercentage = 100 };

            document.Add(new Paragraph("\n"));

            PdfPCell clSerieFusor = new PdfPCell(new Phrase("Serie Fusor", pe.FuenteParrafoBold)) { BorderWidth = .5f };
            PdfPCell clModelo = new PdfPCell(new Phrase("Modelo", pe.FuenteParrafoBold)) { BorderWidth = .5f };
            PdfPCell clInstalado = new PdfPCell(new Phrase("Estado", pe.FuenteParrafoBold)) { BorderWidth = .5f };
            PdfPCell clFecha = new PdfPCell(new Phrase("Fecha", pe.FuenteParrafoBold)) { BorderWidth = .5f };
            PdfPCell clFolio = new PdfPCell(new Phrase("Folio Reporte", pe.FuenteParrafoBold)) { BorderWidth = .5f };
            PdfPCell clCliente = new PdfPCell(new Phrase("Ubicación", pe.FuenteParrafoBold)) { BorderWidth = .5f };

            tblCliente.AddCell(clSerieFusor);
            tblCliente.AddCell(clCliente);
            tblCliente.AddCell(clModelo);
            tblCliente.AddCell(clInstalado);
            tblCliente.AddCell(clFecha);
            tblCliente.AddCell(clFolio);

            tblCliente.HorizontalAlignment = Element.ALIGN_LEFT;

            foreach (string Fusor in Fusores)
            {
                comando.Connection = conexion.AbrirConexion();
                comando.CommandText = "GenerarReporteFusoresNuevo";
                comando.Parameters.Clear();
                comando.Parameters.AddWithValue("@Fusor", Fusor);
                comando.CommandType = CommandType.StoredProcedure;
                reporte = comando.ExecuteReader();

                while (reporte.Read())
                {
                    PdfPCell cSerieFusor = new PdfPCell(new Phrase(reporte[0].ToString(), pe.FuenteParrafo)) { BorderWidth = .5f, Padding = 2, Rowspan = 2 };
                    PdfPCell cCliente = new PdfPCell(new Phrase(reporte[1].ToString(), pe.FuenteParrafo)) { BorderWidth = .5f, Padding = 2, Rowspan = 2 };
                    PdfPCell cModelo = new PdfPCell(new Phrase(reporte[2].ToString(), pe.FuenteParrafo)) { BorderWidth = .5f, Padding = 2, Rowspan = 2 };
                    Fecha = Convert.ToDateTime(reporte[3].ToString());
                    PdfPCell cFecha = new PdfPCell(new Phrase(Fecha.ToString("dd/MM/yyyy"), pe.FuenteParrafo)) { BorderWidth = .5f, Padding = 2 };

                    //DateTime FechaRetiro = Convert.ToDateTime(reporte[4].ToString());
                    string FechaRetiro = reporte[4].ToString();
                    PdfPCell cFechaRetiro;
                    if (FechaRetiro != "")
                    {
                        Fecha = Convert.ToDateTime(reporte[4].ToString());
                        cFechaRetiro = new PdfPCell(new Phrase(Fecha.ToString("dd/MM/yyyy"), pe.FuenteParrafo)) { BorderWidth = .5f, Padding = 2 };
                    }
                    else
                    {
                        cFechaRetiro = new PdfPCell(new Phrase(FechaRetiro, pe.FuenteParrafo)) { BorderWidth = .5f, Padding = 2 };
                    }
                    PdfPCell cInstalado = new PdfPCell(new Phrase("Instalado", pe.FuenteParrafo)) { BorderWidth = .5f, Padding = 2 };

                    PdfPCell cInstaladoR = new PdfPCell(new Phrase("Retirado", pe.FuenteParrafo)) { BorderWidth = .5f, Padding = 2 };
                    PdfPCell cFolio = new PdfPCell(new Phrase(reporte[5].ToString(), pe.FuenteParrafo)) { BorderWidth = .5f, Padding = 2 };
                    PdfPCell cFolioRetiro = new PdfPCell(new Phrase(reporte[6].ToString(), pe.FuenteParrafo)) { BorderWidth = .5f, Padding = 2 };

                    tblCliente.AddCell(cSerieFusor);
                    tblCliente.AddCell(cCliente);
                    tblCliente.AddCell(cModelo);

                    //INSTALADO
                    tblCliente.AddCell(cInstalado);
                    tblCliente.AddCell(cFecha);
                    //tblCliente.AddCell(cSerieSaliente);
                    tblCliente.AddCell(cFolio);

                    //RETIRADO
                    tblCliente.AddCell(cInstaladoR);
                    tblCliente.AddCell(cFechaRetiro);
                    tblCliente.AddCell(cFolioRetiro);

                }
                reporte.Close();
            }
            document.Add(tblCliente);
            Fusores.Clear();
        }

        public void GenerarReporte(Document document, Reporte nuevoReporte)
        {
            var pe = new Pdf();
            //Tabla para cuando se requiera hacer reporte por cliente
            //PdfPTable tblCliente = new PdfPTable(4) { WidthPercentage = 80 };
            PdfPTable tblCliente = (RangoHabilitado) ? new PdfPTable(3) { WidthPercentage = 80 }:new PdfPTable(4) { WidthPercentage = 80 }; ;

            SortedSet<string> lstClientes = new SortedSet<string>();
            this.TipoBusqueda = nuevoReporte.TipoBusqueda;
            document.Add(new Paragraph("\n"));

            //Si estamos filtrando por cliente, entonces se añade una tabla con los titulos de los datos que nos arrojara
            if (nuevoReporte.TipoBusqueda == "Cliente")
            {
                PdfPCell clSerie = new PdfPCell(new Phrase("Serie", pe.FuenteParrafoBold)) { BorderWidth = .5f };
                tblCliente.AddCell(clSerie);
                if (!nuevoReporte.RangoHabilitado)
                {
                    PdfPCell clFecha = new PdfPCell(new Phrase("Fecha Servicio", pe.FuenteParrafoBold)) { BorderWidth = .5f };
                    tblCliente.AddCell(clFecha);
                }
                PdfPCell clTecnico = new PdfPCell(new Phrase("Técnico", pe.FuenteParrafoBold)) { BorderWidth = .5f };
                PdfPCell clFolio = new PdfPCell(new Phrase("Numero Folio", pe.FuenteParrafoBold)) { BorderWidth = .5f };

                tblCliente.AddCell(clTecnico);
                tblCliente.AddCell(clFolio);

                tblCliente.HorizontalAlignment = Element.ALIGN_LEFT;
                document.Add(tblCliente);
            }
            bool PrimerRegistro = true;
            //Recorremos el arreglo que nos genero la consulta
            while (reporte.Read())
            {
                ReporteServicios reporteServicio = new ReporteServicios()
                {
                    Cliente = reporte[1].ToString(),
                    Marca = reporte[2].ToString(),
                    Modelo = reporte[3].ToString(),
                    Contador = int.Parse(reporte[5].ToString()),
                    Fecha = Convert.ToDateTime(reporte[6].ToString()),
                    Tecnico = reporte[7].ToString(),
                    Serie = reporte[4].ToString(),
                    NumeroFolio = reporte[0].ToString(),
                    Fusor = reporte[8].ToString(),
                    FusorSaliente = reporte[9].ToString(),
                    ServicioRealizado = reporte[10].ToString(),
                    ReporteFallo = reporte[11].ToString()
                };

                //if (!lstFechas.Contains(reporteServicio.Fecha.ToString()) && (nuevoReporte.TipoBusqueda == "Fecha"  || nuevoReporte.RangoHabilitado))
                //{
                //    lstFechas.Add(reporteServicio.Fecha.ToString());
                //    document.Add(new Paragraph(reporteServicio.Fecha.ToString("dd/MM/yyyy"), pe.FuenteParrafoBold));
                //    lstSeries.Clear();
                //    lstClientes.Clear();
                //    AgregarSerieAlDocumento(document, reporteServicio);
                //}

                if (!lstFechas.Contains(reporteServicio.Fecha.ToString()))
                {
                    lstFechas.Add(reporteServicio.Fecha.ToString());
                    document.Add(new Paragraph(reporteServicio.Fecha.ToString("dd/MM/yyyy"), pe.FuenteFecha));
                    lstSeries.Clear();
                    lstClientes.Clear();
                    AgregarSerieAlDocumento(document, reporteServicio);
                }
                else
                {
                    AgregarSerieAlDocumento(document, reporteServicio);
                }
            }
            lstClientes.Clear();
            lstFechas.Clear();
            lstSeries.Clear();
            reporte.Close();
        }

        public void AgregarSerieAlDocumento(Document document, ReporteServicios ReporteServicio)
        {
            var pe = new Pdf();
            //tblSeries = new PdfPTable(3) { WidthPercentage = 80 };
            //Verificamos que la serie no este en la lista
            if (!lstSeries.Contains(ReporteServicio.Serie))
            {
                document.Add(new Chunk());
                lstSeries.Add(ReporteServicio.Serie);
                tblSeries = new PdfPTable(3) { WidthPercentage = 80 };
                
                //Verificamos que el cliente no este en la lista
                //if (!lstClientes.Contains(ReporteServicio.Cliente) && TipoBusqueda != "Cliente" && TipoBusqueda != "Marca")
                //{
                //    //Si no esta, lo agregamos al documento
                //    document.Add(new Paragraph("CLIENTE: "+ReporteServicio.Cliente, pe.FuenteParrafoBold));
                //    document.Add(new Chunk());
                //}
                tblSeries = AgregarTablaSerie(ReporteServicio.Serie, ReporteServicio.Marca, ReporteServicio.Modelo);
                document.Add(tblSeries);
                AgregarDatosServicio(document, ReporteServicio);
            }
            else
            {
                iTextSharp.text.pdf.draw.LineSeparator LineaSeparacion = new iTextSharp.text.pdf.draw.LineSeparator() { Offset = 2f };
                document.Add(new Chunk(LineaSeparacion));

                AgregarDatosServicio(document,ReporteServicio);
            }
        }

        public PdfPTable AgregarTablaSerie(string Serie, string Marca, string Modelo)
        {
            Pdf pdf = new Pdf();

            //Encabezados

            PdfPTable tblSerie = new PdfPTable(3) { HorizontalAlignment = Element.ALIGN_LEFT, WidthPercentage = 80, PaddingTop = 10f };
            PdfPCell clMarca;
            PdfPCell clModelo;

            //Encabezados
            PdfPCell clEncabezadoSerie = new PdfPCell(new Phrase("SERIE", pdf.FuenteParrafoBold)) { BorderWidth = .5f };
            PdfPCell clEncabezadoMarca = new PdfPCell(new Phrase("MARCA", pdf.FuenteParrafoBold)) { BorderWidth = .5f };
            PdfPCell clEncabezadoModelo = new PdfPCell(new Phrase("MODELO", pdf.FuenteParrafoBold)) { BorderWidth = .5f };

            tblSerie.AddCell(clEncabezadoSerie);
            tblSerie.AddCell(clEncabezadoMarca);
            tblSerie.AddCell(clEncabezadoModelo);



            //PdfPCell clSerie = (TipoBusqueda == "Serie") ? new PdfPCell(new Phrase("", pdf.FuenteParrafoBold)) { BorderWidth = .5f }: new PdfPCell(new Phrase(Serie, pdf.FuenteParrafoBold)) { BorderWidth = .5f };
            PdfPCell clSerie = new PdfPCell(new Phrase(Serie, pdf.FuenteParrafo)) { BorderWidth = .5f };
            tblSerie.AddCell(clSerie);

            clMarca = (TipoBusqueda != "MARCA" && TipoBusqueda != "MODELO") ? new PdfPCell(new Phrase(Marca, pdf.FuenteParrafo)) { BorderWidth = .5f }: new PdfPCell(new Phrase("", pdf.FuenteParrafo)) { BorderWidth = .5f }; 
            clModelo = (TipoBusqueda != "MODELO") ? new PdfPCell(new Phrase(Modelo, pdf.FuenteParrafo)) { BorderWidth = .5f }: new PdfPCell(new Phrase("", pdf.FuenteParrafo)) { BorderWidth = .5f }; 

            tblSerie.AddCell(clMarca);
            tblSerie.AddCell(clModelo);

            return tblSerie;
        }

        //Esta parte no me gusta hay que cambiarla
        public void AgregarDatosServicio(Document document, ReporteServicios ReporteServicio)
        {
            var pe = new Pdf();
            string Cliente = "";
            //Colocamos los datos del servicio
            DateTime Fecha = Convert.ToDateTime(ReporteServicio.Fecha);
            string espacioVacio = new string(' ', 55);
            //if (TipoBusqueda != "Cliente" || Cliente != "")
            //{
            //    Cliente = ReporteServicio.Cliente;
            //}
            Cliente = ReporteServicio.Cliente;
            //if (RangoHabilitado)
            //{
            //    document.Add(new Paragraph(Cliente + new string(' ', 60 - Cliente.Length) + ReporteServicio.NumeroFolio.ToString(), pe.FuenteParrafo));
            //}
            //else
            //{
            //    document.Add(new Paragraph(Cliente + new string(' ', 60 - Cliente.Length) + Fecha.ToString("dd/MM/yyyy") + "                               " + ReporteServicio.NumeroFolio.ToString(), pe.FuenteParrafo));
            //}
            document.Add(new Paragraph(Cliente + new string(' ', 60 - Cliente.Length) + ReporteServicio.NumeroFolio.ToString(), pe.FuenteParrafoBold));

            document.Add(new Paragraph("DIAGNOSTICO: " + ReporteServicio.ReporteFallo.ToUpper(), pe.FuenteParrafo));
            document.Add(new Paragraph("SERVICIO: " + ReporteServicio.ServicioRealizado.ToUpper(), pe.FuenteParrafo));
            string Fusor = ReporteServicio.Fusor;
            if (!string.IsNullOrWhiteSpace(Fusor)) 
            { 
                document.Add(new Paragraph("FUSOR INSTALADO: " + ReporteServicio.Fusor.ToUpper() + "                " + "FUSOR RETIRADO: " + ReporteServicio.FusorSaliente.ToUpper(), pe.FuenteParrafo));
            }

            document.Add(new Paragraph("CONTADOR: " + string.Format("{0:n0}", int.Parse(ReporteServicio.Contador.ToString())), pe.FuenteParrafo));

        }


    }
}
