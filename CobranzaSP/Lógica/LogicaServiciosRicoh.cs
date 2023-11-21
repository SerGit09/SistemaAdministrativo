using CobranzaSP.Modelos;
using DocumentFormat.OpenXml.Bibliography;
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
using CobranzaSP.Formularios;
using System.Windows.Forms;
using DocumentFormat.OpenXml.Office2010.Excel;

namespace CobranzaSP.Lógica
{
    internal class LogicaServiciosRicoh
    {
        private CD_Conexion conexion = new CD_Conexion();
        SqlCommand comando = new SqlCommand();
        SqlDataReader reporte;
        LogicaModulosCliente lgModuloEquipo = new LogicaModulosCliente();
        LogicaModulosCliente lgModuloCliente = new LogicaModulosCliente();
        AccionesLógica NuevaAccion = new AccionesLógica();


        //LISTAS
        SortedSet<string> lstSeries = new SortedSet<string>();
        SortedSet<string> lstFechas = new SortedSet<string>();
        SortedSet<string> lstReportes = new SortedSet<string>();
        SortedSet<string> lstModulos = new SortedSet<string>();


        PdfPTable tblSeries;
        PdfPTable tblModulos;

        string TipoBusqueda = "";

        string Serie = "";

        public string RegistroServicio(Servicio nuevoServicio, string sp)
        {
            int respuesta = 0;
            string AccionRealizada;
            comando.Connection = conexion.AbrirConexion();
            comando.CommandText = sp;
            comando.CommandType = CommandType.StoredProcedure;

            //Preguntamos que accion realizaremos en la base de datos para posteriormente mostrale al usuario la accion que realizo
            AccionRealizada = (sp == "AgregarServicioRicoh") ? "agrego" : "modifico";

            //if (nuevoServicio.IdServicio != 0)
            //{
            //    comando.Parameters.AddWithValue("@IdServicio", nuevoServicio.IdServicio);
            //}

            comando.Parameters.AddWithValue("@NumeroFolio", nuevoServicio.NumeroFolio);
            comando.Parameters.AddWithValue("@IdCliente", nuevoServicio.IdCliente);
            comando.Parameters.AddWithValue("IdModelo", nuevoServicio.IdModelo);
            comando.Parameters.AddWithValue("@NumeroSerie", nuevoServicio.Serie);
            comando.Parameters.AddWithValue("@Contador", nuevoServicio.Contador);
            comando.Parameters.AddWithValue("@Fecha", nuevoServicio.Fecha);
            comando.Parameters.AddWithValue("@Tecnico", nuevoServicio.Tecnico);
            comando.Parameters.AddWithValue("@ServicioRealizado", nuevoServicio.ServicioRealizado);
            comando.Parameters.AddWithValue("@FallaReportada", nuevoServicio.ReporteFallo);

            respuesta = comando.ExecuteNonQuery();
            string Mensaje = (respuesta > 0) ? "Registro se " + AccionRealizada + " correctamente" : "Algo salio mal, no se " + AccionRealizada + " el registro";

            comando.Parameters.Clear();
            conexion.CerrarConexion();
            return Mensaje;
        }
        public string ObtenerClaveParteRicoh(string Descripcion)
        {
            SqlDataReader leer;
            string Clave = "";
            comando.Connection = conexion.AbrirConexion();
            comando.CommandText = "SeleccionarClaveParteRicoh";
            comando.CommandType = CommandType.StoredProcedure;
            comando.Parameters.Clear();


            comando.Parameters.AddWithValue("@Descripcion", Descripcion);
            leer = comando.ExecuteReader();
            while (leer.Read())
            {
                Clave = leer[0].ToString();
            }
            comando.Parameters.Clear();
            conexion.CerrarConexion();
            leer.Close();
            return Clave;
        }

        public string ObtenerModeloEquipo(string SerieEquipo)
        {
            SqlDataReader leer;
            string Modelo = "";
            comando.Connection = conexion.AbrirConexion();
            comando.CommandText = "SeleccionarModeloEquipo";
            comando.CommandType = CommandType.StoredProcedure;
            comando.Parameters.Clear();


            comando.Parameters.AddWithValue("@Serie", SerieEquipo);
            leer = comando.ExecuteReader();
            while (leer.Read())
            {
                Modelo = leer[0].ToString();
            }
            comando.Parameters.Clear();
            conexion.CerrarConexion();
            leer.Close();
            return Modelo;
        }

        //Con este metodo agregaremos un nuevo registro por cada clave cada vez que se registre un reporte en base al contador
        public void ActualizarModulosEquipo(Servicio nuevoServicio)
        {
            string Modelo = DeterminarModeloModulo(nuevoServicio.Modelo);
            //Actualizaremos cada clave, tanto en los modulos_clientes como en historial_modulo
            foreach (ModuloEquipo Modulo in LogicaModulosCliente.ModulosEquipo)
            {
                ActualizarModulo(nuevoServicio, Modulo,Modelo);
            }
        }

        public void ActualizarModulo(Servicio nuevoServicio, ModuloEquipo Modulo, string Modelo)
        {
            Modulo_Cliente NuevoModulo = new Modulo_Cliente()
            {
                IdModelo = NuevaAccion.BuscarId(Modelo, "ObtenerIdModeloModulo"),
                
                Reporte = nuevoServicio.NumeroFolio,
                Serie = nuevoServicio.Serie,
                Paginas = nuevoServicio.Contador,
                Estado = "ACTUALIZADO",
                Clave = Modulo.Clave,
                Observacion = ""
            };
            NuevoModulo.IdModulo = lgModuloCliente.BuscarIdModulo(Modulo.Modulo, NuevoModulo.IdModelo);
            lgModuloEquipo.RegistrarModulo(NuevoModulo, "AgregarEstadoModuloEquipo");
        }
        public string DeterminarModeloModulo(string Modelo)
        {
            var prefijosValidos = new HashSet<string> { "MP-400", "MP-500", "90" };
            //Para saber si se trata de un fusor
            if (prefijosValidos.Any(prefijo => Modelo.StartsWith(prefijo)))
            {
                return "4002";
            }

            return "5054";
        }

        public bool DeterminarTipoReporte(Reporte NuevoReporte, string Serie)
        {
            bool DatosObtenidos = true;
            this.Serie = Serie;
            switch (NuevoReporte.TipoBusqueda)
            {
                //case "Serie": DatosObtenidos = DeterminarGenerarReporteConRangoFecha(nuevoReporte); break;
                default: DatosObtenidos = ObtenerDatosReporteParametroEspecifico(NuevoReporte); break;
            }

            return DatosObtenidos;
        }

        public bool ObtenerDatosReporteParametroEspecifico(Reporte NuevoReporte)
        {
            bool DatosObtenidos = true;
            comando.Connection = conexion.AbrirConexion();
            comando.CommandText = "DatosReporteServicioRicoh";
            comando.CommandType = CommandType.StoredProcedure;
            comando.Parameters.Clear();
            comando.Parameters.AddWithValue("@CampoBusqueda", NuevoReporte.ParametroBusqueda);
            comando.Parameters.AddWithValue("@RangoHabilitado", NuevoReporte.RangoHabilitado);
            comando.Parameters.AddWithValue("@FechaInicial", NuevoReporte.FechaInicio);
            comando.Parameters.AddWithValue("@FechaFinal", NuevoReporte.FechaFinal);
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

        public void Pdf(Reporte NuevoReporte)
        {
            string NombreArchivo = @"C:\Users\DELL PC\Documents\Reportes\" + NuevoReporte.TipoBusqueda + DateTime.Now.ToString("ddMMyyyyHHmmss") + ".pdf";

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

            //if (!string.IsNullOrEmpty(NuevoReporte.TipoBusquedaAdicional) && NuevoReporte.TipoBusqueda != "Cliente")
            //{
            //    Paragraph ParrafoSubtitulo = new Paragraph("POR " + NuevoReporte.TipoBusquedaAdicional + " : " + Cliente, pe.FuenteTitulo) { Alignment = Element.ALIGN_CENTER };
            //    document.Add(ParrafoSubtitulo);
            //}

            //Colocar el titulo del reporte
            string NombreReporte = ColocarNombreReporte(NuevoReporte);
            Paragraph titulo = new Paragraph(NombreReporte, pe.FuenteTitulo) { Alignment = Element.ALIGN_CENTER };
            document.Add(titulo);

            //Verificamos si tenemos rango de fecha
            if (NuevoReporte.RangoHabilitado || NuevoReporte.TipoBusqueda == "Fecha")
            {
                Paragraph Fechas = new Paragraph("DEL " + NuevoReporte.FechaInicio.ToString("dd/MM/yyyy") + " AL " + NuevoReporte.FechaFinal.ToString("dd/MM/yyyy"), pe.FuenteFecha) { Alignment = Element.ALIGN_CENTER };
                document.Add(Fechas);
            }

            //A partir de aqui va a variar el contenido de los pdfs, dependiendo cual se eligio
            switch (NuevoReporte.TipoBusqueda)
            {
                //case "Fusores": GenerarReporteFusor(document); break;
                //case "Fusor": GenerarReporteFusor(document); break;
                default: GenerarReporte(document, NuevoReporte); break;
            }

            lstSeries.Clear();
            reporte.Close();
            document.Close();

            //Abrimos el pdf 
            pe.AbrirPdf(NombreArchivo);
        }


        public string ColocarNombreReporte(Reporte NuevoReporte)
        {
            string NombreReporte = "REPORTE SERVICIO TECNICO";
            switch (NuevoReporte.TipoBusqueda)
            {
                case "Cliente": NombreReporte += " POR CLIENTE"; break;
                case "Serie": NombreReporte += " POR SERIE: " + NuevoReporte.ParametroBusqueda.ToUpper(); break;
                case "Modelo": NombreReporte += "POR MODELO: " + NuevoReporte.ParametroBusqueda; break;
                case "Modulo": NombreReporte += "POR MODULO: " + NuevoReporte.ParametroBusqueda; break;
            }
            return NombreReporte;
        }

        public void GenerarReporte(Document document, Reporte nuevoReporte)
        {
            var pe = new Pdf();
            //Tabla para cuando se requiera hacer reporte por cliente
            //PdfPTable tblCliente = new PdfPTable(4) { WidthPercentage = 80 };
            PdfPTable tblCliente = (nuevoReporte.RangoHabilitado) ? new PdfPTable(3) { WidthPercentage = 80 } : new PdfPTable(4) { WidthPercentage = 80 }; ;

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
                ReporteModulo ReporteModulo = new ReporteModulo()
                {
                    NumeroFolio = reporte[0].ToString(),
                    Cliente = reporte[1].ToString(),
                    Marca = reporte[2].ToString(),
                    Modelo = reporte[3].ToString(),
                    Serie = reporte[4].ToString(),
                    Contador = int.Parse(reporte[5].ToString()),
                    Fecha = Convert.ToDateTime(reporte[6].ToString()),
                    Tecnico = reporte[7].ToString(),
                    ServicioRealizado = reporte[8].ToString(),
                    ReporteFallo = reporte[9].ToString(),
                    Modulo = reporte[10].ToString(),
                    Clave = reporte[11].ToString(),
                    Estado = reporte[12].ToString(),
                    Paginas = int.Parse(reporte[13].ToString())
                };
                MessageBox.Show(ReporteModulo.Clave);

                if (!lstFechas.Contains(ReporteModulo.Fecha.ToString()))
                {
                    lstFechas.Add(ReporteModulo.Fecha.ToString());
                    document.Add(new Paragraph(ReporteModulo.Fecha.ToString("dd/MM/yyyy"), pe.FuenteFecha));
                    //Aqui iremos agregando los reportes para evitar que se repitan los datos
                    AgregarReporteAlDocumento(document, ReporteModulo);

                    lstSeries.Clear();

                }
                else
                {
                    //AgregarReporteAlDocumento(document, ReporteModulo);
                    //Deberemos de agregar los modulos de dicho reporte
                }
            }
            lstReportes.Clear();
            lstFechas.Clear();
            lstSeries.Clear();
        }

        public void AgregarReporteAlDocumento(Document document, ReporteModulo ReporteModulo)
        {
            var pe = new Pdf();
            if (!lstReportes.Contains(ReporteModulo.NumeroFolio.ToString()))
            {
                lstReportes.Add(ReporteModulo.NumeroFolio.ToString());

                //Posiblemente aqui agreguemos los datos finales y si en dado caso se cambio algun modulo

                AgregarSerieAlDocumento(document, ReporteModulo);
                if(ReporteModulo.Modulo != "")
                {
                    CrearTablaModulos(document, ReporteModulo);
                }
            }
            else
            {

            }
        }



        public void AgregarSerieAlDocumento(Document document, ReporteModulo ReporteModulo)
        {
            var pe = new Pdf();
            //tblSeries = new PdfPTable(3) { WidthPercentage = 80 };
            //Verificamos que la serie no este en la lista
            if (!lstSeries.Contains(ReporteModulo.Serie))
            {
                document.Add(new Chunk());
                lstSeries.Add(ReporteModulo.Serie);
                tblSeries = new PdfPTable(3) { WidthPercentage = 80 };

                tblSeries = AgregarTablaSerie(ReporteModulo.Serie, ReporteModulo.Marca, ReporteModulo.Modelo);
                document.Add(tblSeries);
                AgregarDatosServicio(document, ReporteModulo);
                //Agregamos al cliente y el numero de reporte
                document.Add(new Paragraph(ReporteModulo.Cliente + new string(' ', 60 - ReporteModulo.Cliente.Length) + ReporteModulo.NumeroFolio.ToString(), pe.FuenteParrafoBold));
            }
            else
            {
                iTextSharp.text.pdf.draw.LineSeparator LineaSeparacion = new iTextSharp.text.pdf.draw.LineSeparator() { Offset = 2f };
                document.Add(new Chunk(LineaSeparacion));

                AgregarDatosServicio(document, ReporteModulo);
            }
        }

        public void CrearTablaModulos(Document document, ReporteModulo reporteModulo)
        {
            Pdf pdf = new Pdf();
            PdfPTable tblModulos = new PdfPTable(3) { HorizontalAlignment = Element.ALIGN_LEFT, WidthPercentage = 80, PaddingTop = 10f };

            //Encabezados
            PdfPCell clEncabezadoModulo = new PdfPCell(new Phrase("MODULO", pdf.FuenteParrafoBold)) { BorderWidth = .5f };
            PdfPCell clEncabezadoClave = new PdfPCell(new Phrase("CLAVE", pdf.FuenteParrafoBold)) { BorderWidth = .5f };
            PdfPCell clEncabezadoContador = new PdfPCell(new Phrase("CONTADOR", pdf.FuenteParrafoBold)) { BorderWidth = .5f };

            tblModulos.AddCell(clEncabezadoModulo);
            tblModulos.AddCell(clEncabezadoClave);
            tblModulos.AddCell(clEncabezadoContador);

            //Agregamos el primer registro
        }

        public void AgregarModuloATabla(ReporteModulo ReporteModulo)
        {
            Pdf pdf = new Pdf();
            //DATOS
            if (ReporteModulo.Estado == "INSTALADO" || ReporteModulo.Estado == "ACTUALIZADO")
            {
                //Agregaremos a la lista de modulos
                //lstModulos.Add()

                PdfPCell clModulo = new PdfPCell(new Phrase(ReporteModulo.Modulo, pdf.FuenteParrafo)) { BorderWidth = .5f };
                PdfPCell clClave = new PdfPCell(new Phrase(ReporteModulo.Clave, pdf.FuenteParrafo)) { BorderWidth = .5f };
                PdfPCell clContador = new PdfPCell(new Phrase(ReporteModulo.Paginas.ToString(), pdf.FuenteParrafo)) { BorderWidth = .5f };
            }
            else//Si no quiere decir que es un modulo retirado
            {

            }
        }


        public PdfPTable AgregarTablaSerie(string Serie, string Marca, string Modelo)
        {
            Pdf pdf = new Pdf();

            //Encabezados

            PdfPTable tblSerie = new PdfPTable(3) { HorizontalAlignment = Element.ALIGN_LEFT, WidthPercentage = 80, PaddingTop = 10f };
            

            //Encabezados
            PdfPCell clEncabezadoSerie = new PdfPCell(new Phrase("SERIE", pdf.FuenteParrafoBold)) { BorderWidth = .5f };
            PdfPCell clEncabezadoMarca = new PdfPCell(new Phrase("MARCA", pdf.FuenteParrafoBold)) { BorderWidth = .5f };
            PdfPCell clEncabezadoModelo = new PdfPCell(new Phrase("MODELO", pdf.FuenteParrafoBold)) { BorderWidth = .5f };

            tblSerie.AddCell(clEncabezadoSerie);
            tblSerie.AddCell(clEncabezadoMarca);
            tblSerie.AddCell(clEncabezadoModelo);

            //DATOS
            PdfPCell clMarca;
            PdfPCell clModelo;
            //PdfPCell clSerie = (TipoBusqueda == "Serie") ? new PdfPCell(new Phrase("", pdf.FuenteParrafoBold)) { BorderWidth = .5f }: new PdfPCell(new Phrase(Serie, pdf.FuenteParrafoBold)) { BorderWidth = .5f };
            PdfPCell clSerie = new PdfPCell(new Phrase(Serie, pdf.FuenteParrafo)) { BorderWidth = .5f };
            tblSerie.AddCell(clSerie);

            clMarca = (TipoBusqueda != "Marca" && TipoBusqueda != "Modelo") ? new PdfPCell(new Phrase(Marca, pdf.FuenteParrafo)) { BorderWidth = .5f } : new PdfPCell(new Phrase("", pdf.FuenteParrafo)) { BorderWidth = .5f };
            clModelo = (TipoBusqueda != "Modelo") ? new PdfPCell(new Phrase(Modelo, pdf.FuenteParrafo)) { BorderWidth = .5f } : new PdfPCell(new Phrase("", pdf.FuenteParrafo)) { BorderWidth = .5f };

            tblSerie.AddCell(clMarca);
            tblSerie.AddCell(clModelo);

            return tblSerie;
        }

        //Esta parte la tendremos que colocar hasta el final, junto con el modulo que se haya cambiado
        public void AgregarDatosServicio(Document document, ReporteModulo ReporteModulo)
        {
            var pe = new Pdf();
            document.Add(new Paragraph("DIAGNOSTICO: " + ReporteModulo.ReporteFallo.ToUpper(), pe.FuenteParrafo));
            document.Add(new Paragraph("SERVICIO: " + ReporteModulo.ServicioRealizado.ToUpper(), pe.FuenteParrafo));

            document.Add(new Paragraph("CONTADOR: " + string.Format("{0:n0}", int.Parse(ReporteModulo.Contador.ToString())), pe.FuenteParrafo));
        }
    }
}
