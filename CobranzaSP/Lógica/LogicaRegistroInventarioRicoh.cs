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
using Microsoft.Win32;

namespace CobranzaSP.Lógica
{
    internal class LogicaRegistroInventarioRicoh
    {
        private CD_Conexion conexion = new CD_Conexion();
        private AccionesLógica NuevaAccion = new AccionesLógica();
        SqlCommand comando = new SqlCommand();
        SqlDataReader reporte;
        PdfPTable tblClavesUsadas;

        SortedSet<string> lstFechas = new SortedSet<string>();
        SortedSet<string> lstClaves = new SortedSet<string>();
        public bool ObtenerDatosReporte(Reporte NuevoReporte)
        {
            bool DatosObtenidos = true;
            comando.Connection = conexion.AbrirConexion();
            comando.CommandText = "ReporteRegistroInventarioPartesRico";
            comando.CommandType = CommandType.StoredProcedure;
            comando.Parameters.Clear();
            comando.Parameters.AddWithValue("@ParametroBusqueda", NuevoReporte.ParametroBusqueda);
            comando.Parameters.AddWithValue("@FechaInicial", NuevoReporte.FechaInicio);
            comando.Parameters.AddWithValue("@FechaFinal", NuevoReporte.FechaFinal);
            reporte = comando.ExecuteReader();
            if (!reporte.HasRows)
            {
                reporte.Close();
                return DatosObtenidos = false;
            }

            Pdf(NuevoReporte);
            return DatosObtenidos;
        }

        public void ObtenerDatosResumen(Reporte NuevoReporte)
        {
            comando.Connection = conexion.AbrirConexion();
            comando.CommandText = "CalcularTotalesInventarioRicoh";
            comando.CommandType = CommandType.StoredProcedure;
            comando.Parameters.Clear();
            comando.Parameters.AddWithValue("@FechaInicio", NuevoReporte.FechaInicio);
            comando.Parameters.AddWithValue("@FechaFinal", NuevoReporte.FechaFinal);
            comando.Parameters.AddWithValue("@Parametro", NuevoReporte.ParametroBusqueda);
            reporte = comando.ExecuteReader();
            if (!reporte.HasRows)
            {
                reporte.Close();
            }

        }

        public void Pdf(Reporte NuevoReporte)
        {
            string NombreArchivo = @"\\DESKTOP-D3OQEJR\Archivos Compartidos\Reportes\Partes\" + "RegistroPartesRicoh" + DateTime.Now.ToString("ddMMyyyyHHmmss") + ".pdf";

            FileStream fs = new FileStream(NombreArchivo, FileMode.Create);
            iTextSharp.text.Document document = new Document(PageSize.LETTER);
            document.SetMargins(25f, 25f, 25f, 25f);
            //Colocamos el pdf en horizontal
            document.SetPageSize(iTextSharp.text.PageSize.LETTER);
            PdfWriter pw = PdfWriter.GetInstance(document, fs);

            //Instanciamos la clase para la paginacion
            var pe = new Pdf();
            pe.ColocarFormatoSuperior = true;
            pw.PageEvent = pe;

            document.Open();

            tblClavesUsadas = new PdfPTable(3);

            //Colocar el titulo del reporte
            string NombreReporte = "MOVIMIENTOS INVENTARIO PARTES RICOH";
            Paragraph titulo = new Paragraph(NombreReporte, pe.FuenteTitulo) { Alignment = Element.ALIGN_CENTER };
            document.Add(titulo);

            //Verificamos si tenemos rango de fecha
            
            Paragraph Fechas = new Paragraph("DEL " + NuevoReporte.FechaInicio.ToString("dd/MM/yyyy") + " AL " + NuevoReporte.FechaFinal.ToString("dd/MM/yyyy"), pe.FuenteFecha) { Alignment = Element.ALIGN_CENTER };
            document.Add(Fechas);

            while (reporte.Read())
            {
                MovimientoParteRicoh NuevoMovimiento = new MovimientoParteRicoh()
                {
                    IdRegistro = int.Parse(reporte[0].ToString()),
                    NumeroParte = reporte[1].ToString(),
                    TipoMovimiento = reporte[2].ToString(),
                    Cantidad = int.Parse(reporte[3].ToString()),
                    ClienteProveedor = reporte[4].ToString(),
                    Fecha = DateTime.Parse(reporte[5].ToString())
                };

                if (!lstFechas.Contains(NuevoMovimiento.Fecha.ToString()))//Estamos en la misma fecha
                {
                    //Si tenemos dicha fecha en la lista, agregamos los datos que tenemos de la anterior fecha
                    document.Add(tblClavesUsadas);//Requerimos mover esta linea a otra
                                                  //Reiniciamos nuestra tabla para agregarle datos de otra fecha, si hay
                    tblClavesUsadas = new PdfPTable(3);

                    iTextSharp.text.pdf.draw.LineSeparator lineSeparator = new iTextSharp.text.pdf.draw.LineSeparator() { Offset = 2f };
                    document.Add(new Chunk(lineSeparator));

                    Paragraph Fecha = new Paragraph(NuevoMovimiento.Fecha.ToString("dd/MM/yyyy"), pe.FuenteParrafoGrande) { Alignment = Element.ALIGN_LEFT }; ;

                    //Añadimos una cadena con un espacio muy grande para que se posicione a una distancia alejada

                    //Agregamos la fecha tanto a la lista que tenemos de fechas, como al documento colocando la cantidad que tenemos en dicho día
                    document.Add(Fecha);
                    lstFechas.Add(NuevoMovimiento.Fecha.ToString());

                    //Creamos una nueva tabla con su primer registro
                    AgregarClaveAlDocumento(document, NuevoMovimiento);
                }
                else
                {
                    AgregarClaveAlDocumento(document, NuevoMovimiento);
                }
            }
            //Agregamos el ultimo registro
            document.Add(tblClavesUsadas);
            reporte.Close();
            document.Add(new Chunk());

            //Agregamos el resumen de las salidas y entradas
            Paragraph Resumen = new Paragraph("RESUMEN ENTRADAS Y SALIDAS DEL " + NuevoReporte.FechaInicio.ToString("dd/MM/yyyy") + " - " + NuevoReporte.FechaFinal.ToString("dd/MM/yyyy"), pe.FuenteTitulo) { Alignment = Element.ALIGN_CENTER };
            document.Add(Resumen);
            document.Add(new Chunk());
            ObtenerDatosResumen(NuevoReporte);
            CrearTablaResumen();

            document.Add(tblClavesUsadas);
            reporte.Close();

            //Reinicamos
            lstFechas.Clear();
            lstClaves.Clear();
            
            document.Close();

            //Abrimos el pdf 
            pe.AbrirPdf(NombreArchivo);
        }


        public void AgregarClaveAlDocumento(Document document, MovimientoParteRicoh NuevoMovimiento)
        {
            Pdf pe = new Pdf();
            if (!lstClaves.Contains(NuevoMovimiento.NumeroParte.ToString()))
            {
                document.Add(tblClavesUsadas);//Requerimos mover esta linea a otra
                //Reiniciamos nuestra tabla para agregarle datos de otra fecha, si hay
                tblClavesUsadas = new PdfPTable(3);

                Paragraph Clave = new Paragraph("-> " + NuevoMovimiento.NumeroParte.ToString(), pe.FuenteParrafoGrande) { Alignment = Element.ALIGN_LEFT }; ;

                //Añadimos una cadena con un espacio muy grande para que se posicione a una distancia alejada

                //Agregamos la fecha tanto a la lista que tenemos de fechas, como al documento colocando la cantidad que tenemos en dicho día
                document.Add(Clave);
                lstClaves.Add(NuevoMovimiento.NumeroParte.ToString());
                CrearTablaClavesUsadas(NuevoMovimiento);
            }
            else
            {
                AgregarDatosATablaClaves(NuevoMovimiento);
            }
        }

        public void CrearTablaClavesUsadas(MovimientoParteRicoh NuevoMovimiento)
        {
            AgregarColumnaATabla("Cantidad");
            AgregarColumnaATabla("Tipo Movimiento");
            AgregarColumnaATabla("Cliente/Proveedor");

            AgregarDatosATablaClaves(NuevoMovimiento);
        }

        public void AgregarColumnaATabla(string texto)
        {
            Pdf pe = new Pdf();
            PdfPCell clCelda = new PdfPCell(new Phrase(texto, pe.FuenteParrafoBold)) { BorderWidth = .5f, Colspan = 1 };
            tblClavesUsadas.AddCell(clCelda);
        }

        public void AgregarDatosATablaClaves(MovimientoParteRicoh NuevoMovimiento)
        {
            AgregarColumnaATabla(NuevoMovimiento.Cantidad.ToString());
            AgregarColumnaATabla(NuevoMovimiento.TipoMovimiento.ToString());
            AgregarColumnaATabla(NuevoMovimiento.ClienteProveedor.ToString());
        }

        public void CrearTablaResumen()
        {
            tblClavesUsadas = new PdfPTable(3);

            //Titulos
            AgregarColumnaATabla("CLAVE");
            AgregarColumnaATabla("CANTIDAD DE SALIDA");
            AgregarColumnaATabla("CANTIDAD DE ENTRADA");

            while (reporte.Read())
            {
                AgregarColumnaATabla(reporte[0].ToString());
                AgregarColumnaATabla(reporte[1].ToString());
                AgregarColumnaATabla(reporte[2].ToString());
            }
        }
    }
}
