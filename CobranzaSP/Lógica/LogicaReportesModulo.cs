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
    internal class LogicaReportesModulo
    {
        private CD_Conexion conexion = new CD_Conexion();
        private AccionesLógica NuevaAccion = new AccionesLógica();
        SqlCommand comando = new SqlCommand();
        SqlDataReader reporte;
        PdfPTable tblRegistroPartesUsadas;
        PdfPTable tblModulos;

        LogicaReporteServicio lgReporteServicio = new LogicaReporteServicio();

        SortedSet<string> lstFechas = new SortedSet<string>();
        SortedSet<string> lstSeries = new SortedSet<string>();
        SortedSet<string> lstSeriesModulos = new SortedSet<string>();
        string TipoBusqueda = "";
        string ParametroBusqueda = "";
        bool DatosEncontrados = true;

        #region DatosReportes

        public bool ObtenerDatosReporte(Reporte NuevoReporte)
        {
            TipoBusqueda = NuevoReporte.TipoBusqueda;
            ParametroBusqueda = NuevoReporte.ParametroBusqueda;
            DatosEncontrados = true;

            switch (TipoBusqueda)
            {
                case "HISTORIAL MODULO": 
                    DatosEncontrados = ObtenerDatosHistorialModulos(NuevoReporte); break;
                case "UBICACION MODULO": 
                    DatosEncontrados = ObtenerDatosUbicacionModulos(NuevoReporte); break;
                case "MODULOS POR SERIE DE EQUIPO":
                    DatosEncontrados = ObtenerDatosModulosSeries(NuevoReporte); break;
                default: DatosEncontrados = lgReporteServicio.DeterminarGenerarReporteConRangoFechaRicoh(NuevoReporte); break;
            }
            return DatosEncontrados;
        }

        public int BuscarIdModulo(string Modulo, string sp, int IdModelo)
        {
            SqlDataReader ver;
            int id = 0;
            comando.Connection = conexion.AbrirConexion();
            comando.CommandText = sp;
            comando.CommandType = CommandType.StoredProcedure;
            comando.Parameters.Clear();

            comando.Parameters.AddWithValue("@CampoBusqueda", Modulo);
            comando.Parameters.AddWithValue("@IdModelo", IdModelo);
            ver = comando.ExecuteReader();

            while (ver.Read())
            {
                id = int.Parse(ver[0].ToString());
            }

            conexion.CerrarConexion();
            ver.Close();
            return id;
        }

        public bool ObtenerDatosModulosSeries(Reporte NuevoReporte)
        {
            comando.Connection = conexion.AbrirConexion();
            if(ParametroBusqueda != "TODOS")
            {
                comando.CommandText = "ReporteSerieModulos";
                comando.CommandType = CommandType.StoredProcedure;
                comando.Parameters.Clear();
                comando.Parameters.AddWithValue("@Serie", NuevoReporte.ParametroBusqueda);
            }
            else
            {
                comando.Parameters.Clear();
                comando.CommandText = "ReporteModulosTodosLosEquipos";
                comando.CommandType = CommandType.StoredProcedure;
            }
            
            reporte = comando.ExecuteReader();
            if (!reporte.HasRows)
            {
                reporte.Close();
                return DatosEncontrados = false;
            }
            return DatosEncontrados;
        }

        

        public bool ObtenerDatosHistorialModulos(Reporte NuevoReporte)
        {
            comando.Connection = conexion.AbrirConexion();
            comando.CommandText = "ReporteHistorialModulo";
            comando.CommandType = CommandType.StoredProcedure;
            comando.Parameters.Clear();
            comando.Parameters.AddWithValue("@ParametroBusqueda", NuevoReporte.ParametroBusqueda);
            comando.Parameters.AddWithValue("@FechaInicial", NuevoReporte.FechaInicio);
            comando.Parameters.AddWithValue("@FechaFinal", NuevoReporte.FechaFinal);
            reporte = comando.ExecuteReader();
            if (!reporte.HasRows)
            {
                reporte.Close();
                return DatosEncontrados = false;
            }
            return DatosEncontrados;
        }

        public bool ObtenerDatosUbicacionModulos(Reporte NuevoReporte)
        {
            comando.Connection = conexion.AbrirConexion();
            comando.CommandText = "ReporteUbicacionModulo";
            comando.CommandType = CommandType.StoredProcedure;
            comando.Parameters.Clear();
            comando.Parameters.AddWithValue("@ParametroBusqueda", NuevoReporte.ParametroBusqueda);
            reporte = comando.ExecuteReader();
            if (!reporte.HasRows)
            {
                reporte.Close();
                return DatosEncontrados = false;
            }

            return DatosEncontrados;
        }
        #endregion

        public bool Pdf(Reporte NuevoReporte)
        {
            TipoBusqueda = NuevoReporte.TipoBusqueda;
            ParametroBusqueda = NuevoReporte.ParametroBusqueda;
            bool CerrarDatos = true;
            //string NombreArchivo = @"C:\\ADMINISTRACION-\archivos compartidos\Reportes\Servicios\" + DateTime.Now.ToString("ddMMyyyyHHmmss") + ".pdf";
            string NombreArchivo;
            string NombreReporte;
            NombreArchivo = @"\\DESKTOP-D3OQEJR\Archivos Compartidos\Reportes\Modulos\" + DateTime.Now.ToString("ddMMyyyyHHmmss") + ".pdf";

            DateTime Fecha;
            FileStream fs = new FileStream(NombreArchivo, FileMode.Create);
            Document document = new Document(PageSize.LETTER);
            document.SetMargins(25f, 25f, 25f, 25f);
            //Colocamos el pdf en horizontal
            document.SetPageSize(iTextSharp.text.PageSize.LETTER.Rotate());
            PdfWriter pw = PdfWriter.GetInstance(document, fs);

            //Instanciamos la clase para la paginacion
            var pe = new Pdf();
            pe.ColocarFormatoSuperior = true;
            pw.PageEvent = pe;
            tblRegistroPartesUsadas = new PdfPTable(5);
            tblModulos = new PdfPTable(8);


            document.Open();

            NombreReporte = ColocarTituloReporte();

            Paragraph TituloReporte = new Paragraph(NombreReporte, pe.FuenteTitulo) { Alignment = Element.ALIGN_CENTER };
            document.Add(TituloReporte);


            if (NuevoReporte.RangoHabilitado)
            {
                Paragraph Fechas = new Paragraph("DEL " + NuevoReporte.FechaInicio.ToString("dd/MM/yyyy") + " AL " + NuevoReporte.FechaFinal.ToString("dd/MM/yyyy"), pe.FuenteTitulo) { Alignment = Element.ALIGN_CENTER };
                document.Add(Fechas);
            }

            //Tabla para cuando se requiera hacer reporte por cliente
            PdfPTable tblPartes = new PdfPTable(6) { WidthPercentage = 100 };

            document.Add(new Paragraph("\n"));

            if (!ObtenerDatosReporte(NuevoReporte))
            {
                return DatosEncontrados;
            }

            //Recorremos los datos para mostrarlos en el pdf
            switch (TipoBusqueda)
            {
                case "SERIE": CerrarDatos = false; break;
                case "CLIENTE":CerrarDatos = false ; break;
                case "RANGO DE FECHA":CerrarDatos = false ; break;
                case "HISTORIAL MODULO": GenerarReporteModulo(document); break;
                case "UBICACION MODULO": GenerarReporteModulo(document); break;
                case "MODULOS POR SERIE DE EQUIPO":
                    if(ParametroBusqueda != "TODOS")
                    {
                        GenerarReporteModulosSerie(document);
                    }
                    else
                    {

                        GenerarReporteModulosEquipos(document);
                    }
                    break;
            }

            if (CerrarDatos)
            {
                document.Close();
                reporte.Close();
                //Abrimos el pdf

                pe.AbrirPdf(NombreArchivo);
            }
            
            

            return DatosEncontrados;
        }


        #region ReporteModulo
        public void GenerarReporteModulo(Document document)
        {
            //Debemos de crear la tabla de modulos
            CrearTablaModulos();

            List<int> Contadores = new List<int>();
            int PrimerContador = 0;
            int UltimoContador = 0;
            string UltimoEstado = "";

            while (reporte.Read())
            {
                ReporteModuloEquipo ModuloEquipo = new ReporteModuloEquipo()
                {
                    Serie = reporte[0].ToString(),
                    Cliente = reporte[1].ToString(),
                    Estado = reporte[2].ToString(),
                    Folio = reporte[3].ToString(),
                    Fecha = DateTime.Parse(reporte[4].ToString()),
                    Contador = int.Parse(reporte[5].ToString()),
                    Observacion = reporte[6].ToString()
                };

                if(TipoBusqueda == "UBICACION MODULO")
                {
                    if(ModuloEquipo.Estado == "INSTALADO")
                    {
                        //Quitamos los registros que teniamos anteriormente y volvemos a crear la tabla
                        CrearTablaModulos();
                        Contadores.Clear();
                    }
                    AgregarDatosATablaModulos(ModuloEquipo, document);
                }
                else
                {
                    AgregarDatosATablaModulos(ModuloEquipo, document);
                }
                
                Contadores.Add(ModuloEquipo.Contador);
                if (ModuloEquipo.Estado == "RETIRADO")
                {
                    //PrimerContador
                    PrimerContador = Contadores[0];

                    // Obtener el último elemento
                    UltimoContador = Contadores[Contadores.Count - 1];
                    ColocarRendimiento(PrimerContador, UltimoContador);

                    //Reiniciamos los contadores
                    Contadores.Clear();
                }
                UltimoEstado = ModuloEquipo.Estado;

            }

            //Debemos de validar que el ultimo registro no sea un retirado, nos causa error
            if (UltimoEstado == "ACTUALIZADO")
            {
                //Ahora deberemos de calcular el rendimiento de dicho modulo
                // Obtener el primer elemento
                PrimerContador = Contadores[0];

                // Obtener el último elemento
                UltimoContador = Contadores[Contadores.Count - 1];

                //Colocamos la diferencia en la tabla
                ColocarRendimiento(PrimerContador, UltimoContador);
            }


            //Agregamos la tabla al documento
            document.Add(tblModulos);
        }

        public void GenerarReporteModulosEquipos(Document document)
        {
            var pe = new Pdf();
            while (reporte.Read())
            {
                ReporteModuloSerie ModuloEquipo = new ReporteModuloSerie()
                {
                    Clave = reporte[0].ToString(),
                    Modulo = reporte[1].ToString(),
                    Fecha = DateTime.Parse(reporte[3].ToString()),
                    Folio = reporte[4].ToString(),
                    Contador = int.Parse(reporte[5].ToString()),
                    Serie = reporte[6].ToString(),
                    Marca = reporte[7].ToString(),
                    Modelo = reporte[8].ToString(),
                    Cliente = reporte[9].ToString(),
                    Observacion = reporte[10].ToString()

                };
                if (!lstSeries.Contains(ModuloEquipo.Serie))//Si no esta en la lista
                {
                    document.Add(tblModulos);
                    iTextSharp.text.pdf.draw.LineSeparator LineaSeparacion = new iTextSharp.text.pdf.draw.LineSeparator() { Offset = 2f };
                    document.Add(new Chunk(LineaSeparacion));
                    lstSeries.Add(ModuloEquipo.Serie);//La agregamos a la lista
                    document.Add(new Paragraph("CLIENTE: " + ModuloEquipo.Cliente.ToUpper(), pe.FuenteParrafoBold));
                    document.Add(new Paragraph("EQUIPO: " + ModuloEquipo.Marca.ToUpper() + " " + ModuloEquipo.Modelo.ToUpper() + " CON SERIE: " +  ModuloEquipo.Serie.ToUpper(), pe.FuenteParrafoBold));
                    document.Add(new Paragraph("\n"));
                    CrearTablaModulosSerie();

                    AgregarDatosATablaModulosSerie(ModuloEquipo);
                }
                else
                {
                    AgregarDatosATablaModulosSerie(ModuloEquipo);
                }
            }
            //Para agregar el ultimo
            document.Add(tblModulos);
        }

        public void ColocarRendimiento(int PrimerContador, int UltimoContador)
        {
            var pe = new Pdf();

            PdfPCell clSerie = new PdfPCell(new Phrase("", pe.FuenteParrafoBold)) { BorderWidth = .5f };
            PdfPCell clCliente = new PdfPCell(new Phrase("", pe.FuenteParrafoBold)) { BorderWidth = .5f };
            PdfPCell clEstado = new PdfPCell(new Phrase("", pe.FuenteParrafoBold)) { BorderWidth = .5f };
            PdfPCell clFolio = new PdfPCell(new Phrase("", pe.FuenteParrafoBold)) { BorderWidth = .5f };
            PdfPCell clObservacion = new PdfPCell(new Phrase("", pe.FuenteParrafoBold)) { BorderWidth = .5f};
            PdfPCell clRendimiento = new PdfPCell(new Phrase("RENDIMIENTO", pe.FuenteParrafoBold)) { BorderWidth = .5f, Colspan = 2 };

            int Rendimiento = UltimoContador - PrimerContador;
            PdfPCell clContador = new PdfPCell(new Phrase(Rendimiento.ToString("N0"), pe.FuenteParrafoBold)) { BorderWidth = .5f };

            tblModulos.AddCell(clSerie);
            tblModulos.AddCell(clCliente);
            tblModulos.AddCell(clEstado);
            tblModulos.AddCell(clFolio);
            tblModulos.AddCell(clObservacion);
            tblModulos.AddCell(clRendimiento);
            tblModulos.AddCell(clContador);
        }

        public void CrearTablaModulos()
        {
            tblModulos = new PdfPTable(8) { WidthPercentage = 100 };
            var pe = new Pdf();

            PdfPCell clSerie = new PdfPCell(new Phrase("SERIE", pe.FuenteParrafoBold)) { BorderWidth = .5f };
            PdfPCell clCliente = new PdfPCell(new Phrase("CLIENTE", pe.FuenteParrafoBold)) { BorderWidth = .5f };
            PdfPCell clEstado = new PdfPCell(new Phrase("ESTADO", pe.FuenteParrafoBold)) { BorderWidth = .5f };
            PdfPCell clFolio = new PdfPCell(new Phrase("FOLIO", pe.FuenteParrafoBold)) { BorderWidth = .5f };
            PdfPCell clFecha = new PdfPCell(new Phrase("FECHA", pe.FuenteParrafoBold)) { BorderWidth = .5f };
            PdfPCell clObservacion = new PdfPCell(new Phrase("OBSERVACIÓN", pe.FuenteParrafoBold)) { BorderWidth = .5f ,Colspan = 2};
            PdfPCell clContador = new PdfPCell(new Phrase("CONTADOR", pe.FuenteParrafoBold)) { BorderWidth = .5f };

            tblModulos.AddCell(clSerie);
            tblModulos.AddCell(clCliente);
            tblModulos.AddCell(clEstado);
            tblModulos.AddCell(clFolio);
            tblModulos.AddCell(clFecha);
            tblModulos.AddCell(clObservacion);
            tblModulos.AddCell(clContador);
        }

        public void AgregarDatosATablaModulos(ReporteModuloEquipo ModuloEquipo, Document document)
        {
            var pe = new Pdf();
            PdfPCell clSerie = new PdfPCell(new Phrase(ModuloEquipo.Serie, pe.FuenteParrafo)) { BorderWidth = .5f };
            PdfPCell clCliente = new PdfPCell(new Phrase(ModuloEquipo.Cliente, pe.FuenteParrafo)) { BorderWidth = .5f };
            PdfPCell clEstado = new PdfPCell(new Phrase(ModuloEquipo.Estado, pe.FuenteParrafo)) { BorderWidth = .5f };
            PdfPCell clFolio = new PdfPCell(new Phrase(ModuloEquipo.Folio, pe.FuenteParrafo)) { BorderWidth = .5f };
            PdfPCell clFecha = new PdfPCell(new Phrase(ModuloEquipo.Fecha.ToString("dd/MM/yyyy"), pe.FuenteParrafo)) { BorderWidth = .5f };
            PdfPCell clObservacion = new PdfPCell(new Phrase(ModuloEquipo.Observacion, pe.FuenteParrafo)) { BorderWidth = .5f ,Colspan = 2 };
            PdfPCell clContador = new PdfPCell(new Phrase(ModuloEquipo.Contador.ToString("N0"), pe.FuenteParrafo)) { BorderWidth = .5f };

            tblModulos.AddCell(clSerie);
            tblModulos.AddCell(clCliente);
            tblModulos.AddCell(clEstado);
            tblModulos.AddCell(clFolio);
            tblModulos.AddCell(clFecha);
            tblModulos.AddCell(clObservacion);
            tblModulos.AddCell(clContador);
        }

        public string ColocarTituloReporte()
        {
            string TituloReporte = "REPORTE ";
            switch (TipoBusqueda)
            {
                case "SERIE":; break;
                case "CLIENTE":; break;
                //case "RANGO DE FECHA": DatosObtenidos = ObtenerFusoresReporte(nuevoReporte); break;
                case "HISTORIAL MODULO": 
                    TituloReporte += "HISTORIAL MODULO" + "\n" + ColocarTipoModulo() + ":" + ParametroBusqueda + "\n" + ColocarMarcaModeloDeModulo(ParametroBusqueda); break;
                case "UBICACION MODULO": 
                    TituloReporte += "UBICACIÓN MODULO" + "\n" + ColocarTipoModulo() + ":" + ParametroBusqueda + "\n" + ColocarMarcaModeloDeModulo(ParametroBusqueda); break;
                case "MODULOS POR SERIE DE EQUIPO": TituloReporte += DeterminarNombreReporteModulosSerie();break;
            }
            return TituloReporte;
        }

        public string DeterminarNombreReporteModulosSerie()
        {
            string Serie = ParametroBusqueda;
            string TituloReporte = "";
            if(Serie != "TODOS")
            {
                TituloReporte = "MODULOS DE EQUIPO " + ColocarMarcaModeloEquipo(ParametroBusqueda);
            }
            else
            {
                TituloReporte = "MODULOS DE EQUIPOS";
            }
            return TituloReporte;
        }

        public string ColocarMarcaModeloEquipo(string Serie)
        {
            SqlDataReader leer;
            string Equipo = "";
            string Marca = "";
            string Modelo = "";
            string Cliente = "";
            comando.Connection = conexion.AbrirConexion();
            comando.CommandText = "ObtenerMarcaModeloEquipo";
            comando.CommandType = CommandType.StoredProcedure;
            comando.Parameters.Clear();

            comando.Parameters.AddWithValue("@Serie", Serie);
            leer = comando.ExecuteReader();
            if (leer.Read())
            {
                Marca = leer[0].ToString();
                Modelo = leer[1].ToString();
                Cliente = leer[2].ToString();
            }
            leer.Close();
            conexion.CerrarConexion();
            Equipo = Marca.ToUpper() + " " + Modelo.ToUpper() + " SERIE: " + Serie;
            Equipo += "\n" + "Cliente: " + Cliente;
            return Equipo;
        }

        public string ColocarMarcaModeloDeModulo(string Clave)
        {
            SqlDataReader leer;
            string Equipo = "";
            string Marca = "";
            string Modelo = "";
            comando.Connection = conexion.AbrirConexion();
            comando.CommandText = "ObtenerMarcaModeloDeClave";
            comando.CommandType = CommandType.StoredProcedure;
            comando.Parameters.Clear();

            comando.Parameters.AddWithValue("@Clave", Clave);
            leer = comando.ExecuteReader();
            if (leer.Read())
            {
                Marca = leer[0].ToString();
                Modelo = leer[1].ToString();
            }
            leer.Close();
            conexion.CerrarConexion();
            Equipo = Marca.ToUpper() + " " + Modelo.ToUpper();
            return Equipo;
        }

        public string ColocarTipoModulo()
        {
            string TipoModulo = "";
            Dictionary<string, string> tipoModuloMapping = new Dictionary<string, string>
            {
                { "UI-", "UNIDAD DE IMAGEN" },
                { "UR-", "UNIDAD DE REVELADO" },
                { "FR-", "FUSOR" },
                { "T-", "TELILLA WEB" },
                { "TT-", "UNIDAD DE REVELADO" },
                { "MT-", "MODULO DE TRANSFERENCIA" }
            };

            foreach (var kvp in tipoModuloMapping)
            {
                if (ParametroBusqueda.StartsWith(kvp.Key))
                {
                    TipoModulo = kvp.Value;
                    break;
                }
            }
            return TipoModulo;
        }
        #endregion

        #region ReporteModulosSerie
        public void GenerarReporteModulosSerie(Document document)
        {
            //Debemos de crear la tabla de modulos
            CrearTablaModulosSerie();

            //Leemos los datos
            while (reporte.Read())
            {
                ReporteModuloSerie ModuloEquipo = new ReporteModuloSerie()
                {
                    Clave = reporte[0].ToString(),
                    Modulo = reporte[1].ToString(),
                    Fecha = DateTime.Parse(reporte[3].ToString()),
                    Folio = reporte[4].ToString(),
                    Contador = int.Parse(reporte[5].ToString()),
                    Observacion = reporte[6].ToString()
                };
                AgregarDatosATablaModulosSerie(ModuloEquipo);
                
            }


            //Agregamos la tabla al documento
            document.Add(tblModulos);
        }

        public void CrearTablaModulosSerie()
        {
            tblModulos = new PdfPTable(8) { WidthPercentage = 100 };
            var pe = new Pdf();

            PdfPCell clClave = new PdfPCell(new Phrase("CLAVE", pe.FuenteParrafoBold)) { BorderWidth = .5f , Colspan = 1};
            PdfPCell clModulo = new PdfPCell(new Phrase("MODULO", pe.FuenteParrafoBold)) { BorderWidth = .5f , Colspan = 2 };
            PdfPCell clFecha = new PdfPCell(new Phrase("FECHA", pe.FuenteParrafoBold)) { BorderWidth = .5f, Colspan = 1 };
            PdfPCell clFolio = new PdfPCell(new Phrase("FOLIO", pe.FuenteParrafoBold)) { BorderWidth = .5f, Colspan = 1 };
            PdfPCell clObservacion = new PdfPCell(new Phrase("OBSERVACIÓN", pe.FuenteParrafoBold)) { BorderWidth = .5f, Colspan = 2 };
            PdfPCell clContador = new PdfPCell(new Phrase("PAGINAS", pe.FuenteParrafoBold)) { BorderWidth = .5f, Colspan = 1 };

            tblModulos.AddCell(clClave);
            tblModulos.AddCell(clModulo);
            tblModulos.AddCell(clFecha);
            tblModulos.AddCell(clFolio);
            tblModulos.AddCell(clObservacion);
            tblModulos.AddCell(clContador);
        }

        public void AgregarDatosATablaModulosSerie(ReporteModuloSerie ModuloEquipo)
        {
            var pe = new Pdf();

            PdfPCell clClave = new PdfPCell(new Phrase(ModuloEquipo.Clave, pe.FuenteParrafo)) { BorderWidth = .5f, Colspan = 1 };
            PdfPCell clModulo = new PdfPCell(new Phrase(ModuloEquipo.Modulo, pe.FuenteParrafo)) { BorderWidth = .5f, Colspan = 2 };
            PdfPCell clFecha = new PdfPCell(new Phrase(ModuloEquipo.Fecha.ToString("dd/MM/yyyy"), pe.FuenteParrafo)) { BorderWidth = .5f, Colspan = 1 };
            PdfPCell clFolio = new PdfPCell(new Phrase(ModuloEquipo.Folio, pe.FuenteParrafo)) { BorderWidth = .5f, Colspan = 1 };
            PdfPCell clObservacion = new PdfPCell(new Phrase(ModuloEquipo.Observacion, pe.FuenteParrafo)) { BorderWidth = .5f, Colspan = 2 };
            PdfPCell clContador = new PdfPCell(new Phrase(ModuloEquipo.Contador.ToString("N0"), pe.FuenteParrafo)) { BorderWidth = .5f, Colspan = 1 };

            tblModulos.AddCell(clClave);
            tblModulos.AddCell(clModulo);
            tblModulos.AddCell(clFecha);
            tblModulos.AddCell(clFolio);
            tblModulos.AddCell(clObservacion);
            tblModulos.AddCell(clContador);
        }
        #endregion

    }
}
