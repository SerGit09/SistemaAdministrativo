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
//using DocumentFormat.OpenXml.Drawing;
//using DocumentFormat.OpenXml.Math;

namespace CobranzaSP.Lógica
{
    internal class LogicaRegistro
    {
        private CD_Conexion conexion = new CD_Conexion();
        private AccionesLógica NuevaAccion = new AccionesLógica();
        SqlCommand comando = new SqlCommand();

        //VARIABLES PARA PDF
        public PdfPTable tblInventarioRegistrosSalidas;
        SortedSet<string> lstFechas = new SortedSet<string>();
        SortedSet<string> lstMarcas = new SortedSet<string>();
        SortedSet<string> lstModelos = new SortedSet<string>();
        SqlDataReader Inventario;
        SqlDataReader Totales;


        string Parametro = "";
        //Variables para obtener el historial de salida de un toner en especifico
        int CantidadSalidaTotal = 0;
        int CantidadEntradaTotal = 0;
        int CantidadInicial = 0;

        //Variable para sumar las cantidades en dado caso de que se den de alta en diferentes registros del inventario
        int Cantidad = 0;

        #region CRUD
        public string AgregarRegistroInventario(RegistroInventario nuevoRegistro)
        {
            SqlDataReader leer;
            int valor = 0;
            comando.Connection = conexion.AbrirConexion();

            comando.CommandText = "AgregarSERegistro";
            comando.CommandType = CommandType.StoredProcedure;

            comando.Parameters.AddWithValue("@IdCartucho", nuevoRegistro.IdCartucho);
            comando.Parameters.AddWithValue("@CantidadSalida", nuevoRegistro.CantidadSalida);
            comando.Parameters.AddWithValue("@CantidadEntrada", nuevoRegistro.CantidadEntrada);
            comando.Parameters.AddWithValue("@CantidadGarantia", nuevoRegistro.CantidadGarantia);


            valor = comando.ExecuteNonQuery();
            comando.Parameters.Clear();
            //Nos ayuda a comprobar si el inventario fue modificado(Dependiendo si se haya modificado algo o no)
            if (valor > 0)
            {
                //En dado caso de que haya modificado el inventario, se agregara el registro a la tabla de registros

                //Este es el que se utiliza en mi base de datos
                comando.CommandText = "AgregarRegistroInventario";

                //BD PRINCIPAL
                //comando.CommandText = "AgregarRegistroInventarioRespaldo";

                comando.CommandType = CommandType.StoredProcedure;

                comando.Parameters.AddWithValue("@IdMarca", nuevoRegistro.IdMarca);
                comando.Parameters.AddWithValue("@IdCartucho", nuevoRegistro.IdCartucho);
                comando.Parameters.AddWithValue("@CantidadSalida", nuevoRegistro.CantidadSalida);
                comando.Parameters.AddWithValue("@CantidadEntrada", nuevoRegistro.CantidadEntrada);
                comando.Parameters.AddWithValue("@CantidadGarantia", nuevoRegistro.CantidadGarantia);


                comando.Parameters.AddWithValue("@Cliente", nuevoRegistro.Cliente);
                comando.Parameters.AddWithValue("@Fecha", nuevoRegistro.Fecha);

                if (nuevoRegistro.NumeroSerie != null)
                {
                    comando.Parameters.AddWithValue("@NumeroSerie", nuevoRegistro.NumeroSerie);
                }
                else
                {
                    comando.Parameters.AddWithValue("@NumeroSerie", " ");
                }

                valor = comando.ExecuteNonQuery();
                comando.Parameters.Clear();
                conexion.CerrarConexion();
                return "Se ha agregado el resgitro correctamente. Se ha actualizado el inventario";
            }
            else
            {
                conexion.CerrarConexion();
                return "No se ha agregado el registro. La cantidad excede la cantidad en existencia";
            }
        }

        //Metodo que agregara al inventario fisico lo que se le ingrese ya sea entrada o salida
        public int ModificarInventario(RegistroInventario nuevoRegistro)
        {
            int valor;
            comando.Connection = conexion.AbrirConexion();
            comando.CommandText = "UpdateInventario";
            comando.CommandType = CommandType.StoredProcedure;

            comando.Parameters.AddWithValue("@IdCartucho", nuevoRegistro.IdCartucho);
            comando.Parameters.AddWithValue("@CantidadSalida", nuevoRegistro.CantidadSalida);
            comando.Parameters.AddWithValue("@CantidadEntrada", nuevoRegistro.CantidadEntrada);
            valor = comando.ExecuteNonQuery();
            comando.Parameters.Clear();
            return valor;
        }

        public void ModificarRegistro(RegistroInventario nuevoRegistro)
        {
            comando.Connection = conexion.AbrirConexion();
            comando.CommandText = "ModificarRegistroInventario";
            comando.CommandType = CommandType.StoredProcedure;

            comando.Parameters.AddWithValue("@IdRegistro", nuevoRegistro.Id);
            comando.Parameters.AddWithValue("@IdMarca", nuevoRegistro.IdMarca);
            comando.Parameters.AddWithValue("@IdCartucho", nuevoRegistro.IdCartucho);
            comando.Parameters.AddWithValue("@CantidadSalida", nuevoRegistro.CantidadSalida);
            comando.Parameters.AddWithValue("@CantidadEntrada", nuevoRegistro.CantidadEntrada);
            comando.Parameters.AddWithValue("@Cliente", nuevoRegistro.Cliente);
            comando.Parameters.AddWithValue("@Fecha", nuevoRegistro.Fecha);

            if (nuevoRegistro.NumeroSerie != null)
                comando.Parameters.AddWithValue("@NumeroSerie", nuevoRegistro.NumeroSerie);
            else
                comando.Parameters.AddWithValue("@NumeroSerie", " ");

            comando.ExecuteNonQuery();

            comando.Parameters.Clear();
        }

        //Metodo que se utilizara solo para las modificaciones de salidas
        public bool ModificarRegistroInventario(RegistroInventario nuevoRegistro, bool modificacionNormal)
        {
            int IdInventario;
            int Valor = 0;
            if (modificacionNormal || Valor > 0)
            {
                if (!modificacionNormal)
                {
                    IdInventario = ObtenerIdInventario(nuevoRegistro.IdCartucho);
                    if (nuevoRegistro.CantidadEntrada > 0)
                    {
                        Valor = ModificarInventario(nuevoRegistro);
                    }
                    else if (!ComprobarCantidadInventario(nuevoRegistro))
                    {
                        return false;
                    }
                    else
                    {
                        Valor = ModificarInventario(nuevoRegistro);
                    }
                }
                ModificarRegistro(nuevoRegistro);
                conexion.CerrarConexion();
                return true;
            }
            return false;
        }

        //Metodo que sera util para hacer cambios o eliminaciones del inventario, dependiendo el stored procedure que se utilice
        public void ModificarInventario(RegistroInventario registroAnterior, string sp)
        {
            int IdInventario = ObtenerIdInventario(registroAnterior.IdCartucho);
            comando.Connection = conexion.AbrirConexion();

            comando.CommandText = sp;
            comando.CommandType = CommandType.StoredProcedure;
            comando.Parameters.AddWithValue("@IdInventario", IdInventario);
            comando.Parameters.AddWithValue("@Id", registroAnterior.Id);
            comando.Parameters.AddWithValue("@CantidadSalida", registroAnterior.CantidadSalida);
            comando.Parameters.AddWithValue("@CantidadEntrada", registroAnterior.CantidadEntrada);
            comando.Parameters.AddWithValue("@CantidadGarantia", registroAnterior.CantidadGarantia);//QUITAR
            //Primero requerimos comprobar de que no se haya ingresado una cantidad de salida mayor a la que se tiene en el inventario


            comando.ExecuteNonQuery();

            comando.Parameters.Clear();
            conexion.CerrarConexion();
        }

        public bool ComprobarCantidadInventario(RegistroInventario registro)
        {
            bool CantidadSuficiente = false;
            int Id = ObtenerIdInventario(registro.IdCartucho);
            comando.Connection = conexion.AbrirConexion();
            comando.CommandText = "ComprobarCantidadInventario";
            comando.CommandType = CommandType.StoredProcedure;
            comando.Parameters.AddWithValue("@CantidadSalida", registro.CantidadSalida);
            comando.Parameters.AddWithValue("@IdInventario", Id);
            comando.Parameters.AddWithValue("@CantidadEntrada", registro.CantidadEntrada);


            SqlDataReader leer = comando.ExecuteReader();
            comando.Parameters.Clear();
            if (leer.Read())
            {
                CantidadSuficiente = true;
            }
            leer.Close();
            return CantidadSuficiente;
        }

        //Metodo que ayudara a buscar el id del inventario cuando se desee borrar algun registro de la base de datos
        public int ObtenerIdInventario(int Modelo)
        {
            SqlDataReader leer;
            int Id = 0;
            comando.Connection = conexion.AbrirConexion();
            comando.CommandText = "BuscarIdInventario";
            comando.CommandType = CommandType.StoredProcedure;
            comando.Parameters.AddWithValue("@Modelo", Modelo);
            leer = comando.ExecuteReader();
            if (leer.Read())
            {
                Id = int.Parse(leer[0].ToString());
            }
            comando.Parameters.Clear();
            leer.Close();
            conexion.CerrarConexion();
            return Id;
        }

        #endregion


        //Metodo que ayudara a saber si la cantidad ya sea de entrada o salida no excede las existencias en inventario

        #region Pdf

        #region ObtenerDatos

        public bool ObtenerDatosReporteRegistroInventario(Reporte NuevoReporte)
        {
            bool DatosEncontrados = true;
            comando.Connection = conexion.AbrirConexion();
            comando.CommandText = "ReporteRegistros";
            comando.CommandType = CommandType.StoredProcedure;
            comando.Parameters.Clear();
            comando.Parameters.AddWithValue("@FechaInicio", NuevoReporte.FechaInicio);
            comando.Parameters.AddWithValue("@FechaFinal", NuevoReporte.FechaFinal);
            comando.Parameters.AddWithValue("@Parametro", NuevoReporte.ParametroBusqueda);
            Inventario = comando.ExecuteReader();
            if (!Inventario.HasRows)
            {
                Inventario.Close();
                return DatosEncontrados = false;
            }
            //GenerarPdf(NuevoReporte);
            return DatosEncontrados;
        }

        #endregion 

        public void GenerarPdf(Reporte NuevoReporte, string Cliente, string Marca)
        {
            //string NombreArchivo = @"\\administracion-pc\ARCHIVOS COMPARTIDOS\Reportes\" + "Reporte" + DateTime.Now.ToString("ddMMyyyyHHmmss") + ".pdf";
            //string NombreArchivo = @"C:\Users\DELL PC\Documents\Base de datos\" + "ReporteRegistros" + DateTime.Now.ToString("ddMMyyyyHHmmss") + ".pdf";
            string NombreArchivo = @"\\administracion-pc\ARCHIVOS COMPARTIDOS\Reportes\Registro Inventario\" + "Reporte" + DateTime.Now.ToString("ddMMyyyyHHmmss") + ".pdf";

            FileStream fs = new FileStream(NombreArchivo, FileMode.Create);
            Document document = new Document(PageSize.LETTER);
            //Ayudará a validar que la generacion del reporte no quede en blanco
            bool registroVacio = true;
            document.SetMargins(25f, 25f, 25f, 25f);
            //Colocamos el pdf en horizontal
            document.SetPageSize(iTextSharp.text.PageSize.LETTER);
            PdfWriter pw = PdfWriter.GetInstance(document, fs);

            //Instanciamos la clase para la paginacion
            var pe = new Pdf();
            pw.PageEvent = pe;
            pe.ColocarFormatoSuperior = true;
            document.Open();
            Parametro = Cliente;


            bool BusquedaCliente = (Cliente != " ") ? true : false;
            //bool BusquedaCliente = (Parametro == "CLIENTE") ? true : false;


            Paragraph Fechas = new Paragraph("FECHA DE INICIO: " + NuevoReporte.FechaInicio.ToString("dd/MM/yyyy") + "       FECHA FINAL: " + NuevoReporte.FechaFinal.ToString("dd/MM/yyyy"), pe.FuenteFecha) { Alignment = Element.ALIGN_CENTER };
            document.Add(Fechas);

            tblInventarioRegistrosSalidas = new PdfPTable(4) { WidthPercentage = 100 };
            Paragraph Detallado;
            Detallado = new Paragraph("DETALLADO ENTRADAS Y SALIDAS", pe.FuenteTitulo) { Alignment = Element.ALIGN_CENTER };

            //Proceso para determinar el titulo que tendra nuestro reporte
            switch (NuevoReporte.TipoBusqueda)
            {
                case "Cliente":
                    Paragraph pCliente = (Parametro == "CADTONER") //Si es cadtoner quiere decir que será proveedor
                    ? new Paragraph("PROVEEDOR: " + Parametro, pe.FuenteTitulo) { Alignment = Element.ALIGN_CENTER }
                    : new Paragraph("CLIENTE: " + Parametro, pe.FuenteTitulo) { Alignment = Element.ALIGN_CENTER };
                    document.Add(pCliente);//Lo agregamos al documento
                    Detallado = (Parametro == "CADTONER")
                        ? new Paragraph("DETALLADO ENTRADAS", pe.FuenteTitulo) { Alignment = Element.ALIGN_CENTER } //Si es cadtoner solo veremos entradas
                        : new Paragraph("DETALLADO SALIDAS", pe.FuenteTitulo) { Alignment = Element.ALIGN_CENTER }; //En dado caso de que sea cualquier cliente veremos solo las salidas
                    ; BusquedaCliente = true; break;
                case "Modelo":
                    Paragraph pModelo;
                    if (Parametro.ToString().StartsWith("RM1-") || Parametro.ToString().StartsWith("D01SE"))
                        pModelo = new Paragraph("Modelo" + Parametro, pe.FuenteTitulo) { Alignment = Element.ALIGN_CENTER };
                    else
                        pModelo = new Paragraph("Modelo " + Marca + " " + Parametro, pe.FuenteTitulo) { Alignment = Element.ALIGN_CENTER };

                    document.Add(pModelo);
                    ; break;
            }
            Paragraph TituloReporte = new Paragraph(ColocarTituloReporte(NuevoReporte.TipoBusqueda, document,Marca), pe.FuenteTitulo);
            Detallado = new Paragraph("DETALLADO ENTRADAS Y SALIDAS", pe.FuenteTitulo) { Alignment = Element.ALIGN_CENTER };

            document.Add(Detallado);

            //Primero obtenemos de la base de datos nuestro inventario 
            SqlDataReader Inventario;
            comando.Connection = conexion.AbrirConexion();
            comando.CommandText = "ReporteRegistros";
            comando.CommandType = CommandType.StoredProcedure;
            comando.Parameters.AddWithValue("@FechaInicio", NuevoReporte.FechaInicio);
            comando.Parameters.AddWithValue("@FechaFinal", NuevoReporte.FechaFinal);
            comando.Parameters.AddWithValue("@Parametro", Parametro);
            Inventario = comando.ExecuteReader();
            comando.Parameters.Clear();

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
                    Cliente_Proveedor = Inventario[4].ToString(),
                    BusquedaCliente = BusquedaCliente,
                    ClaveFusor = Inventario[6].ToString()
                };

                //En dado caso de que sea modelo, nos generara un reporte distinto
                if (NuevoReporte.TipoBusqueda == "Modelo")
                    AgregarRegistroCartucho(document, nuevoReporteRegistro);
                else
                    //AgregarRegistrosATabla(document, nuevoReporteRegistro);

                registroVacio = false;
            }

            if (registroVacio)
            {
                //return "No existen registros en ese rango de fechas";
            }

            document.Add(tblInventarioRegistrosSalidas);//Finalmente agregamos el ultimo registro a la lista 
            //Si estamos generando reporte por modelo, colocara la cantidad final que quede en el ultimo día al final
            if (NuevoReporte.TipoBusqueda == "Modelo")
            {
                Paragraph Cantidad = new Paragraph("                                                                                                                                                              " + CantidadInicial, pe.FuenteParrafoGrande) { Alignment = Element.ALIGN_CENTER };
                document.Add(Cantidad);
            }

            //Reiniciamos nuestros valores
            tblInventarioRegistrosSalidas = new PdfPTable(4) { WidthPercentage = 100 };
            lstFechas.Clear();
            lstMarcas.Clear();
            lstModelos.Clear();
            Inventario.Close();

            //TOTALES
            iTextSharp.text.pdf.draw.LineSeparator lineSeparator = new iTextSharp.text.pdf.draw.LineSeparator() { Offset = 2f };
            document.Add(new Chunk(lineSeparator));
            Paragraph Resumen = new Paragraph("RESUMEN ENTRADAS Y SALIDAS DEL " + NuevoReporte.FechaInicio.ToString("dd/MM/yyyy") + " - " + NuevoReporte.FechaFinal.ToString("dd/MM/yyyy"), pe.FuenteTitulo) { Alignment = Element.ALIGN_CENTER };
            document.Add(Resumen);
            document.Add(new Chunk());

            SqlDataReader Totales;
            comando.Connection = conexion.AbrirConexion();
            comando.CommandText = "CalcularTotales";
            comando.CommandType = CommandType.StoredProcedure;
            comando.Parameters.AddWithValue("@FechaInicio", NuevoReporte.FechaInicio);
            comando.Parameters.AddWithValue("@FechaFinal", NuevoReporte.FechaFinal);
            comando.Parameters.AddWithValue("@Parametro", Cliente);
            Totales = comando.ExecuteReader();
            comando.Parameters.Clear();

            PdfPTable tblTotalesRegistros = new PdfPTable(5);
            //En dado caso que estemos generando por 
            if (NuevoReporte.TipoBusqueda == "Cliente")
                tblTotalesRegistros = new PdfPTable(4);

            PdfPCell clTitulo1 = new PdfPCell(new Phrase("Marca", pe.FuenteParrafoBold)) { BorderWidth = .5f, Colspan = 1 };
            PdfPCell clTitulo2 = new PdfPCell(new Phrase("Modelo", pe.FuenteParrafoBold)) { BorderWidth = .5f, Colspan = 2 };
            tblTotalesRegistros.AddCell(clTitulo1);
            tblTotalesRegistros.AddCell(clTitulo2);


            PdfPCell clTitulo3;
            PdfPCell clTitulo4;
            switch (NuevoReporte.TipoBusqueda)
            {
                case "Cliente":
                    if (Parametro == "CADTONER")//Si es verdadero quiere decir que queremos ver el total de entradas que obtuvimos con nuestro proveedor
                    {
                        clTitulo4 = new PdfPCell(new Phrase("Total Entrada", pe.FuenteParrafoBold)) { BorderWidth = .5f, Colspan = 1 };
                        tblTotalesRegistros.AddCell(clTitulo4);
                    }
                    else//Si no es un cliente, por lo tanto querremos ver el total de sus salidas
                    {
                        clTitulo3 = new PdfPCell(new Phrase("Total Salida", pe.FuenteParrafoBold)) { BorderWidth = .5f, Colspan = 1 };
                        tblTotalesRegistros.AddCell(clTitulo3);
                    }
                    ; break;
                default:
                    clTitulo3 = new PdfPCell(new Phrase("Total Salida", pe.FuenteParrafoBold)) { BorderWidth = .5f, Colspan = 1 };
                    clTitulo4 = new PdfPCell(new Phrase("Total Entrada", pe.FuenteParrafoBold)) { BorderWidth = .5f, Colspan = 1 };
                    tblTotalesRegistros.AddCell(clTitulo3);
                    tblTotalesRegistros.AddCell(clTitulo4);
                    break;
            }

            //Agregamos los totales del modelo o de los cartuchos
            while (Totales.Read())
            {
                PdfPCell clMarca = new PdfPCell(new Phrase(Totales[0].ToString(), pe.FuenteParrafo)) { BorderWidth = .5f, Colspan = 1 };
                PdfPCell clModelo = new PdfPCell(new Phrase(Totales[1].ToString(), pe.FuenteParrafo)) { BorderWidth = .5f, Colspan = 2 };
                tblTotalesRegistros.AddCell(clMarca);
                tblTotalesRegistros.AddCell(clModelo);

                switch (Parametro)
                {

                }

                if (BusquedaCliente)
                {
                    if (Cliente == "CADTONER")//Si es verdadero quiere decir que queremos ver el total de entradas que obtuvimos con nuestro proveedor
                    {
                        PdfPCell clTotalEntrada = new PdfPCell(new Phrase(Totales[3].ToString(), pe.FuenteParrafo)) { BorderWidth = .5f, Colspan = 1 };
                        tblTotalesRegistros.AddCell(clTotalEntrada);
                    }
                    else//Si no es un cliente, por lo tanto querremos ver el total de sus salidas
                    {
                        PdfPCell clTotalSalida = new PdfPCell(new Phrase(Totales[2].ToString(), pe.FuenteParrafo)) { BorderWidth = .5f, Colspan = 1 };
                        tblTotalesRegistros.AddCell(clTotalSalida);
                    }
                }
                else
                {
                    PdfPCell clTotalSalida = new PdfPCell(new Phrase(Totales[2].ToString(), pe.FuenteParrafo)) { BorderWidth = .5f, Colspan = 1 };
                    PdfPCell clTotalEntrada = new PdfPCell(new Phrase(Totales[3].ToString(), pe.FuenteParrafo)) { BorderWidth = .5f, Colspan = 1 };
                    tblTotalesRegistros.AddCell(clTotalSalida);
                    tblTotalesRegistros.AddCell(clTotalEntrada);
                }
            }
        }

        public string ColocarTituloReporte(string TipoBusqueda, Document document, string Marca)
        {
            string Titulo = "DETALLADO ENTRADAS Y SALIDAS";
            Pdf pe = new Pdf();
            switch (TipoBusqueda)
            {
                case "Cliente":
                    Titulo = (Parametro == "CADTONER") ?"DETALLADO ENTRADAS":"DETALLADO SALIDAS";
                    Titulo += "\n";
                    Titulo += (Parametro == "CADTONER") ?"PROVEEDOR: " + Parametro: "CLIENTE: " + Parametro;
                    //BusquedaCliente = true;
                    break;
                case "Modelo":
                    Titulo += "\n";
                    if (Parametro.ToString().StartsWith("RM1-") || Parametro.ToString().StartsWith("D01SE"))
                        Titulo += "Modelo: " + Parametro;
                    else
                        Titulo += "Modelo: " + Marca + Parametro;
                    ; break;
            }
            return Titulo;
        }

        public string ReporteRegistrosInventario(DateTime FechaInicio, DateTime FechaFinal, string Parametro, string Cliente)
        {
            //string NombreArchivo = @"\\Desktop-de0cg86\archivos compartidos\Reportes\" + "Reporte" + DateTime.Now.ToString("ddMMyyyyHHmmss") + ".pdf";
            string NombreArchivo = @"\\DESKTOP-D3OQEJR\Archivos Compartidos\Reportes\Registro Inventario\" + "ReporteRegistro" + DateTime.Now.ToString("ddMMyyyyHHmmss") + ".pdf";
            //string NombreArchivo = @"\\administracion-pc\ARCHIVOS COMPARTIDOS\Reportes\Registro Inventario\" + "Reporte" + DateTime.Now.ToString("ddMMyyyyHHmmss") + ".pdf";

            FileStream fs = new FileStream(NombreArchivo, FileMode.Create);
            Document document = new Document(PageSize.LETTER);
            //Ayudará a validar que la generacion del reporte no quede en blanco
            bool registroVacio = true;
            document.SetMargins(25f, 25f, 25f, 25f);
            //Colocamos el pdf en horizontal
            document.SetPageSize(iTextSharp.text.PageSize.LETTER);
            PdfWriter pw = PdfWriter.GetInstance(document, fs);

            //Instanciamos la clase para la paginacion
            var pe = new Pdf();
            pw.PageEvent = pe;
            pe.ColocarFormatoSuperior = true;
            document.Open();
            Parametro = Cliente;
            //tblInventarioRegistrosSalidas = new PdfPTable(4);
            tblInventarioRegistrosSalidas = new PdfPTable(5);//QUITAR EN CASO DE QUE NO QUEDE


            bool BusquedaCliente = (Cliente != " ") ? true : false;
            //bool BusquedaCliente = (Parametro == "CLIENTE") ? true : false;


            Paragraph Fechas = new Paragraph("DE: " + FechaInicio.ToString("dd/MM/yyyy") + " AL: " + FechaFinal.ToString("dd/MM/yyyy"), pe.FuenteFecha) { Alignment = Element.ALIGN_CENTER };
            document.Add(Fechas);
            Paragraph TituloReporte = new Paragraph(ColocarTituloReporte(Parametro, document, ""), pe.FuenteTitulo) { Alignment = Element.ALIGN_CENTER };
            document.Add(TituloReporte);


            comando.Connection = conexion.AbrirConexion();
            comando.CommandText = "ReporteRegistros";
            comando.CommandType = CommandType.StoredProcedure;
            comando.Parameters.Clear();
            comando.Parameters.AddWithValue("@FechaInicio",FechaInicio);
            comando.Parameters.AddWithValue("@FechaFinal", FechaFinal);
            comando.Parameters.AddWithValue("@Parametro", Parametro);
            Inventario = comando.ExecuteReader();
            //if (!Inventario.HasRows)
            //{
            //    Inventario.Close();
            //    return DatosEncontrados = false;
            //}

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
                    BusquedaCliente = BusquedaCliente,
                    ClaveFusor = Inventario[6].ToString()
                };
                AgregarRegistrosATabla(document, nuevoReporteRegistro);
            }
            
            document.Add(tblInventarioRegistrosSalidas);//Finalmente agregamos el ultimo registro a la lista
                   
            //Reiniciamos nuestras tablas, listas y cerramos en la base de datos
            //tblInventarioRegistrosSalidas = new PdfPTable(4);
            tblInventarioRegistrosSalidas = new PdfPTable(5);
            lstFechas.Clear();
            lstMarcas.Clear();
            lstModelos.Clear();
            Inventario.Close();

            //TOTALES
            iTextSharp.text.pdf.draw.LineSeparator lineSeparator = new iTextSharp.text.pdf.draw.LineSeparator() { Offset = 2f };
            document.Add(new Chunk(lineSeparator));
            Paragraph Resumen = new Paragraph("RESUMEN ENTRADAS Y SALIDAS DEL " + FechaInicio.ToString("dd/MM/yyyy") + " - " + FechaFinal.ToString("dd/MM/yyyy"), pe.FuenteTitulo) { Alignment = Element.ALIGN_CENTER };
            document.Add(Resumen);
            document.Add(new Chunk());

            //ObtenerResumenTotales(NuevoReporte);

            comando.Connection = conexion.AbrirConexion();
            comando.CommandText = "CalcularTotales";
            comando.CommandType = CommandType.StoredProcedure;
            comando.Parameters.Clear();
            comando.Parameters.AddWithValue("@FechaInicio", FechaInicio);
            comando.Parameters.AddWithValue("@FechaFinal", FechaFinal);
            comando.Parameters.AddWithValue("@Parametro", Parametro);
            Totales = comando.ExecuteReader();

            //Definicion de columnas y asignacion de encabezados de columnas
            //PdfPTable tblTotalesRegistros = (BusquedaCliente) ? new PdfPTable(4) : new PdfPTable(5);
            PdfPTable tblTotalesRegistros = (BusquedaCliente) ? new PdfPTable(5) : new PdfPTable(6);
            PdfPCell clTituloMarca = new PdfPCell(new Phrase("Marca", pe.FuenteParrafoBold)) { BorderWidth = .5f, Colspan = 1 };
            PdfPCell clTituloModelo = new PdfPCell(new Phrase("Modelo", pe.FuenteParrafoBold)) { BorderWidth = .5f, Colspan = 2 };
            PdfPCell clTituloSalidas = new PdfPCell(new Phrase("Total Salida", pe.FuenteParrafoBold)) { BorderWidth = .5f, Colspan = 1 };
            PdfPCell clTituloEntradas = new PdfPCell(new Phrase("Total Entrada", pe.FuenteParrafoBold)) { BorderWidth = .5f, Colspan = 1 };
            PdfPCell clTituloGarantias = new PdfPCell(new Phrase("Total Garantias", pe.FuenteParrafoBold)) { BorderWidth = .5f, Colspan = 1 };

            tblTotalesRegistros.AddCell(clTituloMarca);
            tblTotalesRegistros.AddCell(clTituloModelo);

            //if (NuevoReporte.TipoBusqueda)
            //{

            //}
            if (BusquedaCliente && Cliente == "CADTONER")
            {
                //Se trata de un proveedor por lo tanto solo veremos sus entradas
                tblTotalesRegistros.AddCell(clTituloEntradas);
            }
            else//Si no se trata de un cliente o un proveedor veremos tanto las entradas como las salidas
            {
                tblTotalesRegistros.AddCell(clTituloSalidas);
                tblTotalesRegistros.AddCell(clTituloGarantias);
                if (!BusquedaCliente)
                {
                    //En dado caso de que se cumpla la condicion, quiere decir que estamos generando por cliente, por lo tanto solo veremos las salidas que se hayan tenido
                    tblTotalesRegistros.AddCell(clTituloEntradas);
                }
            }


            while (Totales.Read())
            {
                PdfPCell clMarca = new PdfPCell(new Phrase(Totales[0].ToString(), pe.FuenteParrafo)) { BorderWidth = .5f, Colspan = 1 };
                PdfPCell clModelo = new PdfPCell(new Phrase(Totales[1].ToString(), pe.FuenteParrafo)) { BorderWidth = .5f, Colspan = 2 };
                tblTotalesRegistros.AddCell(clMarca);
                tblTotalesRegistros.AddCell(clModelo);
                PdfPCell clTotalSalida = new PdfPCell(new Phrase(Totales[2].ToString(), pe.FuenteParrafo)) { BorderWidth = .5f, Colspan = 1 };
                PdfPCell clTotalEntrada = new PdfPCell(new Phrase(Totales[3].ToString(), pe.FuenteParrafo)) { BorderWidth = .5f, Colspan = 1 };
                PdfPCell clTotalGarantias = new PdfPCell(new Phrase(Totales[4].ToString(), pe.FuenteParrafo)) { BorderWidth = .5f, Colspan = 1 };

                if (BusquedaCliente && Cliente == "CADTONER")
                {
                    tblTotalesRegistros.AddCell(clTotalEntrada);
                }
                else
                {
                    tblTotalesRegistros.AddCell(clTotalSalida);
                    tblTotalesRegistros.AddCell(clTotalGarantias);

                    if (!BusquedaCliente)
                    {
                        tblTotalesRegistros.AddCell(clTotalEntrada);
                    }
                }
            }
            document.Add(tblTotalesRegistros);
            Totales.Close();
            document.Close();

            //Abrimos el pdf
            pe.AbrirPdf(NombreArchivo);
            return "Reporte generado exitosamente";
        }
        public void ObtenerResumenTotales(Reporte NuevoReporte)
        {
            comando.Connection = conexion.AbrirConexion();
            comando.CommandText = "CalcularTotales";
            comando.CommandType = CommandType.StoredProcedure;
            comando.Parameters.Clear();
            comando.Parameters.AddWithValue("@FechaInicio", NuevoReporte.FechaInicio);
            comando.Parameters.AddWithValue("@FechaFinal", NuevoReporte.FechaFinal);
            comando.Parameters.AddWithValue("@Parametro", NuevoReporte.ParametroBusqueda);
            Totales = comando.ExecuteReader();
        }

        //Metodo para obtener cantidad inicial del toner o modelo
        public int CalcularTotalCartucho(DateTime FechaInicio, DateTime FechaFinal, string Modelo)
        {
            int valor = 0;
            SqlDataReader dr;
            comando.Connection = conexion.AbrirConexion();
            comando.CommandText = "CalcularTotalCartucho";
            comando.CommandType = CommandType.StoredProcedure;
            comando.Parameters.Clear();

            comando.Parameters.AddWithValue("@FechaInicio", FechaInicio);
            comando.Parameters.AddWithValue("@FechaFinal", FechaFinal);
            comando.Parameters.AddWithValue("@Modelo", Modelo);
            dr = comando.ExecuteReader();
            if (dr.Read())
            {
                valor = int.Parse(dr[2].ToString());
            }
            comando.Parameters.Clear();
            dr.Close();
            return valor;
        }

        public string ReporteRegistroCartucho(DateTime FechaInicio, DateTime FechaFinal, string TipoBusqueda, string Parametro, string Marca, string ClienteEspecifico)
        {
            //string NombreArchivo = @"\\administracion-pc\ARCHIVOS COMPARTIDOS\Reportes\" + "Reporte" + DateTime.Now.ToString("ddMMyyyyHHmmss") + ".pdf";
            //string NombreArchivo = @"C:\Users\DELL PC\Documents\Base de datos\" + "ReporteRegistros" + DateTime.Now.ToString("ddMMyyyyHHmmss") + ".pdf";
            //string NombreArchivo = @"\\Desktop-de0cg86\archivos compartidos\Reportes\Registro Inventario\" + "Reporte" + DateTime.Now.ToString("ddMMyyyyHHmmss") + ".pdf";
            string NombreArchivo = @"\\DESKTOP-D3OQEJR\Archivos Compartidos\Reportes\Registro Inventario\" + "ReporteRegistro" + DateTime.Now.ToString("ddMMyyyyHHmmss") + ".pdf";

            bool BusquedaCliente = false;

            FileStream fs = new FileStream(NombreArchivo, FileMode.Create);
            Document document = new Document(PageSize.LETTER);
            //Ayudará a validar que la generacion del reporte no quede en blanco
            bool registroVacio = true;
            document.SetMargins(25f, 25f, 25f, 25f);
            //Colocamos el pdf en horizontal
            document.SetPageSize(iTextSharp.text.PageSize.LETTER);
            PdfWriter pw = PdfWriter.GetInstance(document, fs);

            //Instanciamos la clase para la paginacion
            var pe = new Pdf();
            pw.PageEvent = pe;
            pe.ColocarFormatoSuperior = true;
            document.Open();


            Paragraph Fechas = new Paragraph("FECHA DE INICIO: " + FechaInicio.ToString("dd/MM/yyyy") + "       FECHA FINAL: " + FechaFinal.ToString("dd/MM/yyyy"), pe.FuenteFecha) { Alignment = Element.ALIGN_CENTER };

            document.Add(Fechas);
            tblInventarioRegistrosSalidas = new PdfPTable(4);
            Paragraph Detallado;
            //REGISTRO DE SALIDAS O ENTRADAS POR UN CLIENTE EN ESPECIFICO, DEPENDIENDO SI ES PROVEEDOR O CLIENTE
            

            Paragraph pModelo;
            if (Parametro.ToString().StartsWith("RM1-") || Parametro.ToString().StartsWith("D01SE") || Parametro.ToString().StartsWith("R6"))
            {
                pModelo = new Paragraph("Modelo" + Parametro, pe.FuenteTitulo) { Alignment = Element.ALIGN_CENTER };
            }else
                pModelo = new Paragraph("Modelo " + Marca + " " + Parametro, pe.FuenteTitulo) { Alignment = Element.ALIGN_CENTER };
            document.Add(pModelo);
            if (ClienteEspecifico != "")
            {
                Paragraph pClienteEspecifo = new Paragraph("CLIENTE: " + ClienteEspecifico,pe.FuenteTitulo) { Alignment = Element.ALIGN_CENTER};
                document.Add(pClienteEspecifo);
            }
            else
            {
                CantidadInicial = CalcularTotalCartucho(FechaInicio, FechaFinal, Parametro);
            }
            
            //Paragraph pCantidadCartucho = new Paragraph("Cantidad Inicial: " + CantidadInicial, pe.FuenteTitulo) { Alignment = Element.ALIGN_CENTER };
            //document.Add(pCantidadCartucho);
            Detallado = new Paragraph("DETALLADO ENTRADAS Y SALIDAS", pe.FuenteTitulo) { Alignment = Element.ALIGN_CENTER };
                    
            document.Add(Detallado);


            //Primero obtenemos de la base de datos nuestro inventario 
            SqlDataReader Inventario;
            comando.Connection = conexion.AbrirConexion();
            comando.CommandText = (ClienteEspecifico != "") ? "ReporteClienteCartuchoPorFecha" : "ReporteRegistros";  
            comando.CommandType = CommandType.StoredProcedure;
            comando.Parameters.AddWithValue("@FechaInicio", FechaInicio);
            comando.Parameters.AddWithValue("@FechaFinal", FechaFinal);
            comando.Parameters.AddWithValue("@Parametro", Parametro);
            if(ClienteEspecifico != "")
            {
                comando.Parameters.AddWithValue("@Cliente", ClienteEspecifico);
            }
            Inventario = comando.ExecuteReader();
            comando.Parameters.Clear();

            while (Inventario.Read())
            {
                DateTime fechaActual = Convert.ToDateTime(Inventario[5].ToString());
                ReporteRegistro nuevoReporteRegistro = new ReporteRegistro()
                {
                    FechaActual = fechaActual.ToString("dd/MM/yyyy"),
                    Marca = Inventario[0].ToString(),
                    Modelo = Inventario[1].ToString(),
                    CantidadSalida = int.Parse(Inventario[2].ToString()),
                    CantidadEntrada = int.Parse(Inventario[3].ToString()),//es este
                    CantidadGarantia = int.Parse(Inventario[7].ToString()),
                    Cliente_Proveedor = Inventario[4].ToString(),
                    BusquedaCliente = BusquedaCliente,
                    ClaveFusor = Inventario[6].ToString()
                };
                nuevoReporteRegistro.ClienteEspecifico = (ClienteEspecifico != "") ?true:false;
                
                AgregarRegistroCartucho(document, nuevoReporteRegistro);
                registroVacio = false;
            }

            if (registroVacio)
            {
                return "No existen registros en ese rango de fechas";
            }
            
            document.Add(tblInventarioRegistrosSalidas);//Finalmente agregamos el ultimo registro a la lista 
            Paragraph Cantidad = new Paragraph("                                                                                                                                                              " + CantidadInicial, pe.FuenteParrafoGrande) { Alignment = Element.ALIGN_CENTER };
            document.Add(Cantidad);

            tblInventarioRegistrosSalidas = new PdfPTable(4);
            lstFechas.Clear();
            lstMarcas.Clear();
            lstModelos.Clear();
            Inventario.Close();

            //Mostraremos los totales solamente si estamos buscando por cartucho nadamas
            if(ClienteEspecifico == "")
            {
                iTextSharp.text.pdf.draw.LineSeparator lineSeparator = new iTextSharp.text.pdf.draw.LineSeparator() { Offset = 2f };
                document.Add(new Chunk(lineSeparator));
                Paragraph Resumen = new Paragraph("RESUMEN ENTRADAS Y SALIDAS DEL " + FechaInicio.ToString("dd/MM/yyyy") + " - " + FechaFinal.ToString("dd/MM/yyyy"), pe.FuenteTitulo) { Alignment = Element.ALIGN_CENTER };
                document.Add(Resumen);
                document.Add(new Chunk());

                SqlDataReader Totales;
                comando.Connection = conexion.AbrirConexion();
                comando.CommandText = "CalcularTotales";
                comando.CommandType = CommandType.StoredProcedure;
                comando.Parameters.AddWithValue("@FechaInicio", FechaInicio);
                comando.Parameters.AddWithValue("@FechaFinal", FechaFinal);
                comando.Parameters.AddWithValue("@Parametro", Parametro);
                Totales = comando.ExecuteReader();
                comando.Parameters.Clear();

                PdfPTable tblTotalesRegistros = new PdfPTable(6);
                PdfPCell clTitulo1 = new PdfPCell(new Phrase("Marca", pe.FuenteParrafoBold)) { BorderWidth = .5f, Colspan = 1 };
                PdfPCell clTitulo2 = new PdfPCell(new Phrase("Modelo", pe.FuenteParrafoBold)) { BorderWidth = .5f, Colspan = 2 };
                tblTotalesRegistros.AddCell(clTitulo1);
                tblTotalesRegistros.AddCell(clTitulo2);

                //En este caso queremos ver tanto las salidas como las entradas
                PdfPCell clTitulo3 = new PdfPCell(new Phrase("Total Salida", pe.FuenteParrafoBold)) { BorderWidth = .5f, Colspan = 1 };
                PdfPCell clTitulo4 = new PdfPCell(new Phrase("Total Entrada", pe.FuenteParrafoBold)) { BorderWidth = .5f, Colspan = 1 };
                PdfPCell clTitulo5 = new PdfPCell(new Phrase("Total Garantias", pe.FuenteParrafoBold)) { BorderWidth = .5f, Colspan = 1 };
                tblTotalesRegistros.AddCell(clTitulo3);
                tblTotalesRegistros.AddCell(clTitulo4);
                tblTotalesRegistros.AddCell(clTitulo5);


                while (Totales.Read())
                {
                    PdfPCell clMarca = new PdfPCell(new Phrase(Totales[0].ToString(), pe.FuenteParrafo)) { BorderWidth = .5f, Colspan = 1 };
                    PdfPCell clModelo = new PdfPCell(new Phrase(Totales[1].ToString(), pe.FuenteParrafo)) { BorderWidth = .5f, Colspan = 2 };
                    tblTotalesRegistros.AddCell(clMarca);
                    tblTotalesRegistros.AddCell(clModelo);

                    PdfPCell clTotalSalida = new PdfPCell(new Phrase(Totales[2].ToString(), pe.FuenteParrafo)) { BorderWidth = .5f, Colspan = 1 };
                    PdfPCell clTotalEntrada = new PdfPCell(new Phrase(Totales[3].ToString(), pe.FuenteParrafo)) { BorderWidth = .5f, Colspan = 1 };
                    PdfPCell clTotalGarantia = new PdfPCell(new Phrase(Totales[4].ToString(), pe.FuenteParrafo)) { BorderWidth = .5f, Colspan = 1 };
                    tblTotalesRegistros.AddCell(clTotalSalida);
                    tblTotalesRegistros.AddCell(clTotalEntrada);
                    tblTotalesRegistros.AddCell(clTotalGarantia);
                }
                CantidadInicial = CantidadInicial - (CantidadSalidaTotal + CantidadEntradaTotal);


                PdfPCell clTotal = new PdfPCell(new Phrase("TOTAL ", pe.FuenteParrafo)) { BorderWidth = .5f, Colspan = 3 };
                PdfPCell clCantidadTotal = new PdfPCell(new Phrase(CantidadInicial.ToString(), pe.FuenteParrafo)) { BorderWidth = .5f, Colspan = 3 };

                tblTotalesRegistros.AddCell(clTotal);
                tblTotalesRegistros.AddCell(clCantidadTotal);

                document.Add(tblTotalesRegistros);
                Totales.Close();
            }
            
            document.Close();

            //Abrimos el pdf
            pe.AbrirPdf(NombreArchivo);
            return "Reporte generado exitosamente";
        }

        public void AgregarTituloColumnaTabla(string celda, int colspan)
        {
            Pdf pdf = new Pdf();
            PdfPCell clTitulo = new PdfPCell(new Phrase(celda, pdf.FuenteParrafoBold)) { BorderWidth = .5f, Colspan = colspan };
            tblInventarioRegistrosSalidas.AddCell(clTitulo);
        }

        //Metodo para agregar los registros cuando estamos filtrando por modelo
        public void AgregarRegistroCartucho(Document document, ReporteRegistro nuevoReporte)
        {
            Pdf pe = new Pdf();
            //Verificamos que dicha fecha no este en la lista
            if (!lstFechas.Contains(nuevoReporte.FechaActual.ToString()))
            {
                //Si tenemos dicha fecha en la lista, agregamos los datos que tenemos de la anterior fecha
                document.Add(tblInventarioRegistrosSalidas);//Requerimos mover esta linea a otra
                //Reiniciamos nuestra tabla para agregarle datos de otra fecha, si hay
                tblInventarioRegistrosSalidas = new PdfPTable(4);

                iTextSharp.text.pdf.draw.LineSeparator lineSeparator = new iTextSharp.text.pdf.draw.LineSeparator() { Offset = 2f };
                document.Add(new Chunk(lineSeparator));

                Paragraph Fecha;

                if (nuevoReporte.ClienteEspecifico)
                {
                    Fecha = new Paragraph(nuevoReporte.FechaActual.ToString(), pe.FuenteParrafoGrande) { Alignment = Element.ALIGN_LEFT };
                }
                else
                {
                    string Cantidad = "                                                                                                                                                       " + CantidadInicial;
                    Fecha = new Paragraph(nuevoReporte.FechaActual + Cantidad, pe.FuenteParrafoGrande) { Alignment = Element.ALIGN_LEFT };

                }
                //Añadimos una cadena con un espacio muy grande para que se posicione a una distancia alejada

                //Agregamos la fecha tanto a la lista que tenemos de fechas, como al documento colocando la cantidad que tenemos en dicho día
                document.Add(Fecha);
                lstFechas.Add(nuevoReporte.FechaActual.ToString());

                //Creamos una nueva tabla con su primer registro
                CrearTablaModelo(document, nuevoReporte);
            }
            else
            {
                //Como estamos en la misma fecha solamente le agregamos más datos a la tabla en la que estamos actualmente
                AgregarModeloATabla(nuevoReporte);
            }

        }

        public void AgregarRegistrosATabla(Document document, ReporteRegistro nuevoReporte)
        {
            Pdf pdf = new Pdf();

            //Preguntamos si no tenemos agregada dicha fecha
            if (!lstFechas.Contains(nuevoReporte.FechaActual.ToString()))//Si se cumple quiere decir no tenemos dicha fecha
            {
                //Reiniciamos todas nuestras listas
                lstMarcas.Clear();
                lstModelos.Clear();
                //Agregamos la tabla que estabamos llenando al documento
                document.Add(tblInventarioRegistrosSalidas);

                //Determinamos de que tipo de busqueda se trata para saber las columnas que tendra nuesta columna
                //tblInventarioRegistrosSalidas = (nuevoReporte.BusquedaCliente) ? new PdfPTable(2) : new PdfPTable(4);
                tblInventarioRegistrosSalidas = (nuevoReporte.BusquedaCliente) ? new PdfPTable(2) : new PdfPTable(5);

                //Linea separadora
                iTextSharp.text.pdf.draw.LineSeparator lineSeparator = new iTextSharp.text.pdf.draw.LineSeparator() { Offset = 2f };
                document.Add(new Chunk(lineSeparator));
                //Agregamos la fecha tanto al pdf, como al documento
                Paragraph Fecha = new Paragraph(nuevoReporte.FechaActual.ToString(), pdf.FuenteParrafoGrande) { Alignment = Element.ALIGN_LEFT };
                document.Add(Fecha);
                lstFechas.Add(nuevoReporte.FechaActual.ToString());

                //Agregamos la primer marca y modelo de dicha fecha
                lstMarcas.Add(nuevoReporte.Marca);
                document.Add(new Paragraph("-> " + nuevoReporte.Marca));

                RegistrarModelo(document, nuevoReporte);
            }
            else//Si no, estamos en la misma fecha
            {
                //Solo agregamos lo que ira en las tablas
                if (lstMarcas.Contains(nuevoReporte.Marca))//Si es la misma marca entonces entramos a agregar modelos
                {
                    //document.Add(tblInventarioRegistrosSalidas);    
                    //Si ya esta la marca, preguntamos ahora por lo modelos
                    if (!lstModelos.Contains(nuevoReporte.Modelo))//Si no esta el modelo
                    {
                        //Como ya cambiamos de modelo, agregamos al documento el modelo anterior, es decir, su tabla
                        document.Add(tblInventarioRegistrosSalidas);
                        //Agregamos el modelo
                        RegistrarModelo(document, nuevoReporte);
                    }
                    else //Si no esta ese modelo le seguimos agregando
                    {
                        if (nuevoReporte.BusquedaCliente)
                        {
                            AgregarCantidad(document, nuevoReporte);
                        }
                        else
                        {
                            AgregarModeloATabla(nuevoReporte);
                        }
                    }
                }
                else//Si no quiere decir que es una nueva marca
                {
                    //Agregamos la última tabla de modelos de esa marca
                    document.Add(tblInventarioRegistrosSalidas);
                    //Colocamos en el documento la nueva marca
                    lstMarcas.Add(nuevoReporte.Marca);
                    document.Add(new Paragraph("-> " + nuevoReporte.Marca));

                    //Agregamos el primer modelo
                    RegistrarModelo(document, nuevoReporte);
                }
            }
        }

        //Metodo para añadir el modelo a la lista de modelos y creamos su tabla
        public void RegistrarModelo(Document document, ReporteRegistro nuevoReporte)
        {
            lstModelos.Add(nuevoReporte.Modelo);
            document.Add(new Paragraph("      " + nuevoReporte.Modelo));
            //Creamos la tabla donde iran todos los registros de dicho modelo
            CrearTablaModelo(document, nuevoReporte);
        }

        //Se utiliza exclusivamente cuando es por cliente, para ir colocando las cantidades correspondientes
        public void AgregarCantidad(Document document, ReporteRegistro reporte)
        {
            //En dado caso de que sea cadtoner quiere decir que es una entrada
            if (reporte.Cliente_Proveedor == "CADTONER")
            {
                document.Add(new Paragraph("Cantidad: " + reporte.CantidadEntrada));
            }
            else
            {
                document.Add(new Paragraph("Cantidad: " + reporte.CantidadSalida));
            }
        }

        //Registramos un modelo en una tabla ya creada
        public void AgregarModeloATabla(ReporteRegistro registro)
        {
            Pdf pe = new Pdf();
            PdfPCell clCantidadSalidaDato = new PdfPCell(new Phrase(ComprobarValoresO(registro.CantidadSalida.ToString()), pe.FuenteParrafoBold)) { BorderWidth = .5f, Colspan = 1 };
            PdfPCell clCantidadEntradaDato = new PdfPCell(new Phrase(ComprobarValoresO(registro.CantidadEntrada.ToString()), pe.FuenteParrafoBold)) { BorderWidth = .5f, Colspan = 1 };
            PdfPCell clCantidadGarantiaDato = new PdfPCell(new Phrase(ComprobarValoresO(registro.CantidadGarantia.ToString()), pe.FuenteParrafoBold)) { BorderWidth = .5f, Colspan = 1 };//QUITAR


            tblInventarioRegistrosSalidas.AddCell(clCantidadSalidaDato);
            tblInventarioRegistrosSalidas.AddCell(clCantidadEntradaDato);
            tblInventarioRegistrosSalidas.AddCell(clCantidadGarantiaDato);//QUITAR
            //Sumar
            if (registro.CantidadSalida > 0)
                CantidadInicial -= registro.CantidadSalida;
            //Quitar posiblemente
            else if(registro.CantidadGarantia > 0)
                CantidadInicial -= registro.CantidadGarantia;
            if (registro.CantidadEntrada > 0 )
                CantidadInicial += registro.CantidadEntrada;
            PdfPCell clClienteProveedorDato = new PdfPCell(new Phrase(registro.Cliente_Proveedor, pe.FuenteParrafoBold)) { BorderWidth = .5f, Colspan = 2 };
            tblInventarioRegistrosSalidas.AddCell(clClienteProveedorDato);

            if (registro.ClaveFusor != "" && registro.ClaveFusor != " " || registro.Modelo.StartsWith("DR"))
            {
                PdfPCell clFusorDato = new PdfPCell(new Phrase(registro.ClaveFusor, pe.FuenteParrafoBold10)) { BorderWidth = .5f, Colspan = 1 };
                tblInventarioRegistrosSalidas.AddCell(clFusorDato);
            }
            //else
            //{
            //    PdfPCell clFusorDato = new PdfPCell(new Phrase(" ", pe.FuenteParrafoBold10)) { BorderWidth = .5f, Colspan = 1 };
            //    tblInventarioRegistrosSalidas.AddCell(clFusorDato);
            //}
        }

        //Metodo que crea la tabla de un modelo
        public void CrearTablaModelo(Document document, ReporteRegistro nuevoReporte)
        {
            Pdf pe = new Pdf();
            
            if (nuevoReporte.BusquedaCliente)
            {
                AgregarCantidad(document, nuevoReporte);
            }
            else
            {
                int NumeroColumnas = 4;
                //int NumeroColumnas = 5;
                var prefijosValidos = new HashSet<string> { "RM1-", "D01SE", "R6","DR" };
                //Para saber si se trata de un fusor
                if (prefijosValidos.Any(prefijo => nuevoReporte.Modelo.StartsWith(prefijo)))
                {
                    //NumeroColumnas = 5;
                    NumeroColumnas = 6;
                    if (nuevoReporte.CantidadSalida < 0 || nuevoReporte.CantidadGarantia <0)
                    {
                        //NumeroColumnas = 4;
                        NumeroColumnas = 5;
                    }
                }
                tblInventarioRegistrosSalidas = new PdfPTable(NumeroColumnas);
                AgregarTituloColumnaTabla("Salida",1);
                AgregarTituloColumnaTabla("Entrada",1);
                AgregarTituloColumnaTabla("Garantia", 1);//QUITAR
                AgregarTituloColumnaTabla("Cliente/Proveedor",2);

                if(NumeroColumnas > 4)
                {
                    AgregarTituloColumnaTabla("Clave", 1);
                }                                                                                                                                               
                
                AgregarModeloATabla(nuevoReporte);
            }
        }


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
        #endregion
    }
}
