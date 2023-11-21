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
using CobranzaSP.Formularios;

namespace CobranzaSP.Lógica
{
    internal class LogicaInventario
    {
        private CD_Conexion conexion = new CD_Conexion();
        SqlCommand comando = new SqlCommand();
        PdfPTable tblInventarioRegistrosSalidas;
        LogicaRegistro lgRegistro = new LogicaRegistro();
        SqlDataReader Inventario;
        SqlDataReader Totales;


        public void RegistrarInventario(InventarioDatos nuevoCartucho, string sp)
        {
            comando.Connection = conexion.AbrirConexion();
            comando.CommandText = sp;
            comando.CommandType = CommandType.StoredProcedure;

            if (nuevoCartucho.Id > 0)
            {
                comando.Parameters.AddWithValue("@Id", nuevoCartucho.Id);
            }
            comando.Parameters.AddWithValue("@Modelo", nuevoCartucho.Modelo);
            comando.Parameters.AddWithValue("@IdMarca", nuevoCartucho.IdMarca);
            comando.Parameters.AddWithValue("@CantidadOficina", nuevoCartucho.CantidadOficina);
            comando.Parameters.AddWithValue("@Fecha", nuevoCartucho.Fecha);
            comando.Parameters.AddWithValue("@Precio", nuevoCartucho.Precio);

            comando.ExecuteNonQuery();

            comando.Parameters.Clear();
            conexion.CerrarConexion();
        }

        public DataTable MostrarFiltroCartucho(string sp, string Marca)
        {
            DataTable tabla = new DataTable();
            SqlDataReader leer;
            comando.Connection = conexion.AbrirConexion();
            comando.CommandText = sp;
            comando.CommandType = CommandType.StoredProcedure;
            comando.Parameters.AddWithValue("@Marca", Marca);
            leer = comando.ExecuteReader();
            tabla.Load(leer);
            comando.Parameters.Clear();
            conexion.CerrarConexion();
            leer.Close();
            return tabla;
        }

        public void ImprimirInventario()
        {
            //Primero obtenemos de la base de datos nuestro inventario 
            comando.Connection = conexion.AbrirConexion();
            comando.CommandText = "ReporteInventario";
            comando.CommandType = CommandType.StoredProcedure;
            comando.Parameters.Clear();
            Inventario = comando.ExecuteReader();

            string NombreArchivo = @"\\DESKTOP-D3OQEJR\Archivos Compartidos\Reportes\Inventario\"+"Inventario" + DateTime.Now.ToString("ddMMyyyyHHmmss") + ".pdf";
            //string NombreArchivo = @"C:\Users\DELL PC\Documents\Base de datos\" + "Inventario" + DateTime.Now.ToString("dd-MM-yyyy") + ".pdf";
            //string NombreArchivo = @"\\Desktop-de0cg86\archivos compartidos\Reportes\" + "Reporte" + DateTime.Now.ToString("ddMMyyyyHHmmss") + ".pdf";
            FileStream fs = new FileStream(NombreArchivo, FileMode.Create);
            Document document = new Document(PageSize.LETTER);
            document.SetMargins(25f, 25f, 25f, 25f);
            document.SetPageSize(iTextSharp.text.PageSize.LETTER);
            PdfWriter pw = PdfWriter.GetInstance(document, fs);

            var pe = new Pdf();
            pe.ColocarFormatoSuperior = false;
            pw.PageEvent = pe;

            document.Open();

            Paragraph titulo = new Paragraph("EXISTENCIAS CARTUCHOS SPEED TONER", pe.FuenteTitulo18);
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

            PdfPTable tblInventario = new PdfPTable(5);
            //PdfPTable tblInventario = new PdfPTable(5);

            document.Add(new Paragraph("\n"));

            //Agregamos los titulos de la tabla
            PdfPCell clModelo = new PdfPCell(new Phrase("MODELO", pe.FuenteParrafoBold)) { BorderWidth = .5f, Colspan = 2 };
            PdfPCell clCantidadOficina = new PdfPCell(new Phrase("CANTIDAD OFICINA", pe.FuenteParrafoBold)) { BorderWidth = .5f, Colspan = 1 };

            PdfPCell clCantidadSalida = new PdfPCell(new Phrase("CANTIDAD SALIDA", pe.FuenteParrafoBold)) { BorderWidth = .5f, Colspan = 1 };
            PdfPCell clCantidadEntrada = new PdfPCell(new Phrase("CANTIDAD ENTRADA", pe.FuenteParrafoBold)) { BorderWidth = .5f, Colspan = 1 };

            tblInventario.AddCell(clModelo);
            tblInventario.AddCell(clCantidadOficina);
            tblInventario.AddCell(clCantidadSalida);
            tblInventario.AddCell(clCantidadEntrada);

            while (Inventario.Read())
            {
                PdfPCell clModeloDato;
                if (Inventario[1].ToString().StartsWith("RM1-") || Inventario[1].ToString().StartsWith("D01SE"))
                {
                    clModeloDato = new PdfPCell(new Phrase(Inventario[1].ToString(), pe.FuenteParrafo)) { BorderWidth = .5f, Colspan = 2 };
                }
                else
                {
                    clModeloDato = new PdfPCell(new Phrase(Inventario[0].ToString() + " " + Inventario[1].ToString(), pe.FuenteParrafo)) { BorderWidth = .5f, Colspan = 2 };
                }

                //PdfPCell clCantidadOficinaDato = new PdfPCell(new Phrase(Inventario[1].ToString(), pe.FuenteParrafo)) { BorderWidth = .5f, Colspan = 1 };
                PdfPCell clCantidadOficinaDato = new PdfPCell(new Phrase(ComprobarValoresO(Inventario[2].ToString()), pe.FuenteParrafo)) { BorderWidth = .5f, Colspan = 1 };
                PdfPCell clCantidadSalidaDato = new PdfPCell(new Phrase(Inventario[3].ToString(), pe.FuenteParrafo)) { BorderWidth = .5f, Colspan = 1 };
                PdfPCell clCantidadEntradaDato = new PdfPCell(new Phrase(Inventario[4].ToString(), pe.FuenteParrafo)) { BorderWidth = .5f, Colspan = 1 };

                tblInventario.AddCell(clModeloDato);
                tblInventario.AddCell(clCantidadOficinaDato);
                tblInventario.AddCell(clCantidadSalidaDato);
                tblInventario.AddCell(clCantidadEntradaDato);
            }

            //Añadimos la tabla al documento
            document.Add(tblInventario);
            Inventario.Close();

            //Agregamos un salto de página
            document.NewPage();
            AgregarMovimientosRegistro(document);

            //Abrimos el pdf
            pe.AbrirPdf(NombreArchivo);
            conexion.CerrarConexion();
        }

        public void AgregarMovimientosRegistro(Document document)
        {
            var pe = new Pdf();

            pe.ColocarFormatosSuperiores(document);

            //tblInventarioRegistrosSalidas = new PdfPTable(4);
            //bool BusquedaCliente = (Parametro == "CLIENTE") ? true : false;


            Paragraph Fechas = new Paragraph("DE: " + DateTime.Now.ToString("dd/MM/yyyy") + " AL: " + DateTime.Now.ToString("dd/MM/yyyy"), pe.FuenteFecha) { Alignment = Element.ALIGN_CENTER };
            document.Add(Fechas);
            Paragraph TituloReporte = new Paragraph(lgRegistro.ColocarTituloReporte("", document, ""), pe.FuenteTitulo) { Alignment = Element.ALIGN_CENTER };
            document.Add(TituloReporte);
            lgRegistro.tblInventarioRegistrosSalidas = new PdfPTable(5);//QUITAR EN CASO DE QUE NO QUEDE


            comando.Connection = conexion.AbrirConexion();
            comando.CommandText = "ReporteRegistros";
            comando.CommandType = CommandType.StoredProcedure;
            comando.Parameters.Clear();
            comando.Parameters.AddWithValue("@FechaInicio", DateTime.Now.ToString("yyyy-MM-dd"));
            comando.Parameters.AddWithValue("@FechaFinal", DateTime.Now.ToString("yyyy-MM-dd"));
            comando.Parameters.AddWithValue("@Parametro", "");
            Inventario = comando.ExecuteReader();
            if (!Inventario.HasRows)
            {
                Inventario.Close();
                document.Close();
                return;
            }

            //Recorremos los datos que hayamos obtenido
            while (Inventario.Read())
            {
                DateTime fechaActual = Convert.ToDateTime(Inventario[5].ToString());
                ReporteRegistro nuevoReporteRegistro = new ReporteRegistro()
                {
                    FechaActual = fechaActual.ToString("dd/MM/yyyy"),
                    Marca = Inventario[0].ToString(),
                    Modelo = Inventario[1].ToString(),
                    CantidadSalida = int.Parse(Inventario[2].ToString()),
                    CantidadEntrada = int.Parse(Inventario[3].ToString()),
                    CantidadGarantia = int.Parse(Inventario[7].ToString()),
                    Cliente_Proveedor = Inventario[4].ToString(),
                    BusquedaCliente = false,
                    ClaveFusor = Inventario[6].ToString()
                };
                lgRegistro.AgregarRegistrosATabla(document, nuevoReporteRegistro);
            }

            document.Add(lgRegistro.tblInventarioRegistrosSalidas);//Finalmente agregamos el ultimo registro a la lista

            //Reiniciamos nuestras tablas, listas y cerramos en la base de datos
            //tblInventarioRegistrosSalidas = new PdfPTable(4);
            tblInventarioRegistrosSalidas = new PdfPTable(5);

            Inventario.Close();

            //TOTALES
            iTextSharp.text.pdf.draw.LineSeparator lineSeparator = new iTextSharp.text.pdf.draw.LineSeparator() { Offset = 2f };
            document.Add(new Chunk(lineSeparator));
            Paragraph Resumen = new Paragraph("RESUMEN ENTRADAS Y SALIDAS DEL " + DateTime.Now.ToString("dd/MM/yyyy") + " - " + DateTime.Now.ToString("dd/MM/yyyy"), pe.FuenteTitulo) { Alignment = Element.ALIGN_CENTER };
            document.Add(Resumen);
            document.Add(new Chunk());

            //ObtenerResumenTotales(NuevoReporte);

            comando.Connection = conexion.AbrirConexion();
            comando.CommandText = "CalcularTotales";
            comando.CommandType = CommandType.StoredProcedure;
            comando.Parameters.Clear();
            comando.Parameters.AddWithValue("@FechaInicio", DateTime.Now.ToString("yyyy-MM-dd"));
            comando.Parameters.AddWithValue("@FechaFinal", DateTime.Now.ToString("yyyy-MM-dd"));
            comando.Parameters.AddWithValue("@Parametro", "");
            Totales = comando.ExecuteReader();

            //Definicion de columnas y asignacion de encabezados de columnas
            //PdfPTable tblTotalesRegistros = (BusquedaCliente) ? new PdfPTable(4) : new PdfPTable(5);
            PdfPTable tblTotalesRegistros = new PdfPTable(6);
            PdfPCell clTituloMarca = new PdfPCell(new Phrase("Marca", pe.FuenteParrafoBold)) { BorderWidth = .5f, Colspan = 1 };
            PdfPCell clTituloModelo = new PdfPCell(new Phrase("Modelo", pe.FuenteParrafoBold)) { BorderWidth = .5f, Colspan = 2 };
            PdfPCell clTituloSalidas = new PdfPCell(new Phrase("Total Salida", pe.FuenteParrafoBold)) { BorderWidth = .5f, Colspan = 1 };
            PdfPCell clTituloEntradas = new PdfPCell(new Phrase("Total Entrada", pe.FuenteParrafoBold)) { BorderWidth = .5f, Colspan = 1 };
            PdfPCell clTituloGarantias = new PdfPCell(new Phrase("Total Garantias", pe.FuenteParrafoBold)) { BorderWidth = .5f, Colspan = 1 };

            tblTotalesRegistros.AddCell(clTituloMarca);
            tblTotalesRegistros.AddCell(clTituloModelo);

            //if (NuevoReporte.TipoBusqueda)
            //{


            tblTotalesRegistros.AddCell(clTituloSalidas);
            tblTotalesRegistros.AddCell(clTituloGarantias);
            //En dado caso de que se cumpla la condicion, quiere decir que estamos generando por cliente, por lo tanto solo veremos las salidas que se hayan tenido
            tblTotalesRegistros.AddCell(clTituloEntradas);


            while (Totales.Read())
            {
                PdfPCell clMarca = new PdfPCell(new Phrase(Totales[0].ToString(), pe.FuenteParrafo)) { BorderWidth = .5f, Colspan = 1 };
                PdfPCell clModelo = new PdfPCell(new Phrase(Totales[1].ToString(), pe.FuenteParrafo)) { BorderWidth = .5f, Colspan = 2 };
                tblTotalesRegistros.AddCell(clMarca);
                tblTotalesRegistros.AddCell(clModelo);
                PdfPCell clTotalSalida = new PdfPCell(new Phrase(Totales[2].ToString(), pe.FuenteParrafo)) { BorderWidth = .5f, Colspan = 1 };
                PdfPCell clTotalEntrada = new PdfPCell(new Phrase(Totales[3].ToString(), pe.FuenteParrafo)) { BorderWidth = .5f, Colspan = 1 };
                PdfPCell clTotalGarantias = new PdfPCell(new Phrase(Totales[4].ToString(), pe.FuenteParrafo)) { BorderWidth = .5f, Colspan = 1 };


                tblTotalesRegistros.AddCell(clTotalSalida);
                tblTotalesRegistros.AddCell(clTotalGarantias);
                tblTotalesRegistros.AddCell(clTotalEntrada);
            }
            document.Add(tblTotalesRegistros);
            Totales.Close();
            Inventario.Close();
            document.Close();
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
