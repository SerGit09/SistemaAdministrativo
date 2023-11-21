using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.SqlClient;
using System.Net.Mail;
using System.Configuration;
using System.Net;
using System.Security.Cryptography.X509Certificates;
using System.Net.Security;
using System.Net.Mime;
using CobranzaSP.Modelos;
using iTextSharp.text.pdf;
using iTextSharp.text;
using System.IO;

namespace CobranzaSP.Lógica
{
    internal class CobranzaLógica
    {
        private CD_Conexion conexion = new CD_Conexion();
        SqlCommand comando = new SqlCommand();
        AccionesLógica NuevaAccion = new AccionesLógica();
        string NombreArchivo = "";
        string ContenidoCorreo = "";
        int IdCliente = 0;
        double TotalImporte = 0;
        SqlDataReader leer;
        PdfPTable tblCuentas;
        public string Registrar(Cobranza objCobranza)
        {
            int respuesta = 0;
            string mensaje = "";
            bool FacturaDuplicada = false;
            comando.Connection = conexion.AbrirConexion();

            //FacturaDuplicada = NuevaAccion.VerificarDuplicados(objCobranza.Factura, "DuplicadosCobranza");
            if (FacturaDuplicada)
            {
                mensaje = "Factura duplicada!!, No se guardo el registro";
            }
            else
            {
                comando.CommandText = "AgregarCobranza";
                comando.CommandType = CommandType.StoredProcedure;

                comando.Parameters.AddWithValue("@IdCliente", objCobranza.IdCliente);
                comando.Parameters.AddWithValue("@Factura", objCobranza.NumeroFactura);
                comando.Parameters.AddWithValue("@DiasCredito", objCobranza.DiasCredito);
                comando.Parameters.AddWithValue("@Cantidad", objCobranza.Cantidad);
                comando.Parameters.AddWithValue("@FormaPago", objCobranza.FormaPago);
                comando.Parameters.AddWithValue("@FechaFactura", objCobranza.FechaFactura);
                comando.Parameters.AddWithValue("@Observaciones", objCobranza.Observaciones);
                comando.Parameters.AddWithValue("@PromesaPago", objCobranza.PromesaPago);

                respuesta = comando.ExecuteNonQuery();
                mensaje = (respuesta > 0) ? "Registro agregado correctamente" : "Algo ha salido mal, no agrego el registro";

                comando.Parameters.Clear();
            }

            conexion.CerrarConexion();
            return mensaje;
        }

        public string ModificarCobro(Cobranza objCobranza, int Id)
        {
            int respuesta = 0;
            string mensaje = "";
            comando.Connection = conexion.AbrirConexion();
            comando.CommandText = "ModificarCobro";
            comando.CommandType = CommandType.StoredProcedure;

            comando.Parameters.AddWithValue("@Id", Id);
            comando.Parameters.AddWithValue("@IdCliente", objCobranza.IdCliente);
            comando.Parameters.AddWithValue("@Factura", objCobranza.NumeroFactura);
            comando.Parameters.AddWithValue("@DiasCredito", objCobranza.DiasCredito);
            comando.Parameters.AddWithValue("@Cantidad", objCobranza.Cantidad);
            comando.Parameters.AddWithValue("@FechaFactura", objCobranza.FechaFactura);
            comando.Parameters.AddWithValue("@Observaciones", objCobranza.Observaciones);
            comando.Parameters.AddWithValue("@FormaPago", objCobranza.FormaPago);
            comando.Parameters.AddWithValue("@PromesaPago", objCobranza.PromesaPago);

            respuesta = comando.ExecuteNonQuery();
            mensaje = (respuesta > 0) ? "Registro modificado correctamente" : "Algo ha salido mal, no se modifico el registro";

            comando.Parameters.Clear();

            conexion.CerrarConexion();
            return mensaje;
        }

        public string EliminarCobro(int Id)
        {
            int respuesta = 0;
            string mensaje = "";
            comando.Connection = conexion.AbrirConexion();
            comando.CommandText = "EliminarCobro";
            comando.CommandType = CommandType.StoredProcedure;

            comando.Parameters.AddWithValue("@Id", Id);

            respuesta = comando.ExecuteNonQuery();
            mensaje = (respuesta > 0) ? "Registro eliminado correctamente" : "Algo ha salido mal, no se elimino el registro";

            comando.Parameters.Clear();

            conexion.CerrarConexion();
            return mensaje;
        }

        public string CuentaPagada(CuentaPagada nuevaCuentaPagada, int Id)
        {
            int respuesta = 0;
            string mensaje = "";
            comando.Connection = conexion.AbrirConexion();
            comando.CommandText = "CuentaPagada";
            comando.CommandType = CommandType.StoredProcedure;

            comando.Parameters.AddWithValue("@Id", Id);
            comando.Parameters.AddWithValue("@Factura", nuevaCuentaPagada.Factura);
            comando.Parameters.AddWithValue("@IdCliente", nuevaCuentaPagada.IdCliente);
            comando.Parameters.AddWithValue("@Cantidad", nuevaCuentaPagada.Cantidad);
            comando.Parameters.AddWithValue("@FechaFactura", nuevaCuentaPagada.FechaFactura);
            comando.Parameters.AddWithValue("@TipoPago", "FACTURA");

            respuesta = comando.ExecuteNonQuery();
            mensaje = (respuesta > 0) ? "Cuenta cobrada correctamente" : "Algo ha salido mal, no se pudo cobrar la cuenta";

            comando.Parameters.Clear();

            conexion.CerrarConexion();
            return mensaje;
        }

        public int MostrarDiasCredito(string NombreCliente)
        {
            int DiasCredito = 0;
            SqlDataReader dr;
            string mensaje = "";
            comando.Connection = conexion.AbrirConexion();
            NuevaAccion.BuscarId(NombreCliente, "BuscarIdCliente");
            comando.CommandText = "MostrarDiasCredito";
            comando.CommandType = CommandType.StoredProcedure;

            comando.Parameters.AddWithValue("@IdCliente", NuevaAccion.BuscarId(NombreCliente, "BuscarIdCliente"));
            dr = comando.ExecuteReader();
            comando.Parameters.Clear();
            if (dr.Read())
            {
                DiasCredito = int.Parse(dr[0].ToString());
            }

            comando.Parameters.Clear();

            conexion.CerrarConexion();
            return DiasCredito;
        }

        #region EnvioDeCorreo

        public void ObtenerDatosCorreo()
        {

        }

        public void ObtenerDatosCorreoPDF()
        {

        }

        public string GenerarCorreo(string NombreCliente)
        {
            TotalImporte = 0;
            ContenidoCorreo = "";
            comando.Connection = conexion.AbrirConexion();
            comando.CommandText = "DatosCorreoFacturas";
            comando.CommandType = CommandType.StoredProcedure;
            comando.Parameters.AddWithValue("@NombreCliente", NombreCliente);
            leer = comando.ExecuteReader();
            comando.Parameters.Clear();
            GenerarEstadoDeCuentaPdf();
            return EnviarCorreo(leer, NombreCliente);
        }

        public void GenerarEstadoDeCuentaPdf()
        {
            NombreArchivo = @"\\Desktop-de0cg86\archivos compartidos\Reportes\Estados de Cuenta\" + "Estado de Cuenta "+DateTime.Now.ToString("ddMMyyyyHHmmss") + ".pdf";
            double SaldoTotal = 0;
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

            Paragraph titulo = new Paragraph("ESTADO DE CUENTA CLIENTES", pe.FuenteTitulo) { Alignment = Element.ALIGN_CENTER };
            document.Add(titulo);

            document.Add(new Chunk("\n"));

            //En esta paarte se incluira las cantidades
            CrearTablaCuentas();

            while (leer.Read())
            {
                CuentaEstado NuevaCuenta = new CuentaEstado()
                {
                    Folio = leer[1].ToString(),
                    FechaFactura = Convert.ToDateTime(leer[6].ToString()),
                    FechaVencimiento = Convert.ToDateTime(leer[7].ToString()),
                    Saldo = double.Parse(leer[5].ToString()),
                    Estado = leer[8].ToString()
                };
                IdCliente = int.Parse(leer[0].ToString());
                ColocarDatosEnTablaCuentas(NuevaCuenta);
                //Vamos sumando las cuentas
                SaldoTotal += NuevaCuenta.Saldo;
            }
            ColocarTotalEnTabla(SaldoTotal);
            document.Add(tblCuentas);
            leer.Close();
            document.Close();
        }

        public void CrearTablaCuentas()
        {
            Pdf pdf = new Pdf();
            tblCuentas = new PdfPTable(4);
            PdfPCell clEncabezadoFolio = new PdfPCell(new Phrase("FOLIO", pdf.FuenteParrafoBold)) { BorderWidth = .5f };
            PdfPCell clEncabezadoFechaFactura = new PdfPCell(new Phrase("FECHA FACTURA", pdf.FuenteParrafoBold)) { BorderWidth = .5f };
            PdfPCell clEncabezadoFechaVencimiento = new PdfPCell(new Phrase("FECHA VENCIMIENTO", pdf.FuenteParrafoBold)) { BorderWidth = .5f };
            PdfPCell clEncabezadoSaldo = new PdfPCell(new Phrase("SALDO", pdf.FuenteParrafoBold)) { BorderWidth = .5f };

            tblCuentas.AddCell(clEncabezadoFolio);
            tblCuentas.AddCell(clEncabezadoFechaFactura);
            tblCuentas.AddCell(clEncabezadoFechaVencimiento);
            tblCuentas.AddCell(clEncabezadoSaldo);

            CrearContenidoCorreo();
        }

        public void ColocarDatosEnTablaCuentas(CuentaEstado NuevaCuenta)
        {
            Pdf pdf = new Pdf();
            PdfPCell clFolio = new PdfPCell(new Phrase(NuevaCuenta.Folio, pdf.FuenteParrafo)) { BorderWidth = .5f };
            PdfPCell clFechaFactura = new PdfPCell(new Phrase(NuevaCuenta.FechaFactura.ToString("dd-MM-yyyy"), pdf.FuenteParrafo)) { BorderWidth = .5f };
            PdfPCell clFechaVencimiento = new PdfPCell(new Phrase(NuevaCuenta.FechaVencimiento.ToString("dd-MM-yyyy"), pdf.FuenteParrafo)) { BorderWidth = .5f };
            PdfPCell clSaldo = new PdfPCell(new Phrase(NuevaCuenta.Saldo.ToString("C"), pdf.FuenteParrafo)) { BorderWidth = .5f };

            tblCuentas.AddCell(clFolio);
            tblCuentas.AddCell(clFechaFactura);
            tblCuentas.AddCell(clFechaVencimiento);
            tblCuentas.AddCell(clSaldo);

            if (NuevaCuenta.Estado == "Vencida")
            {
                ColocarCuentaEnTablaDeCorreo(NuevaCuenta);
            }
        }

        public void ColocarCuentaEnTablaDeCorreo(CuentaEstado NuevaCuenta)
        {
            ContenidoCorreo += "<tr>" +
                        "<td>" + NuevaCuenta.Folio.ToString() + "</td>" +
                        "<td>" + NuevaCuenta.FechaFactura.ToString("dd-MM-yyyy") + "</td>" +
                        "<td>" + NuevaCuenta.FechaVencimiento.ToString("dd-MM-yyyy") + "</td>" +
                        "<td ALIGN=Right>" + NuevaCuenta.Saldo.ToString("C") + "</td>" +
                    "</tr>";
            TotalImporte += NuevaCuenta.Saldo;
        }

        public void ColocarTotalEnTabla(double SaldoTotal)
        {
            Pdf pdf = new Pdf();
            PdfPCell clEncabezadoFolio = new PdfPCell(new Phrase("", pdf.FuenteParrafoBold)) { BorderWidth = .5f };
            PdfPCell clTituloTotal = new PdfPCell(new Phrase("TOTAL", pdf.FuenteParrafoBold)) { BorderWidth = .5f, Colspan = 2 };
            PdfPCell clSaldoTotal = new PdfPCell(new Phrase(SaldoTotal.ToString("C"), pdf.FuenteParrafoBold)) { BorderWidth = .5f };

            tblCuentas.AddCell(clEncabezadoFolio);
            tblCuentas.AddCell(clTituloTotal);
            tblCuentas.AddCell(clSaldoTotal);
            ColocarTotalEnTablaEnCorreo();
        }

        public void ColocarTotalEnTablaEnCorreo()
        {
            ContenidoCorreo += "<tr>" +
                                    "<td>" + "</td>" +
                                    "<td>" + "</td>" +
                                    "<td>" + "TOTAL" + "</td>" +
                                    //"<td>" + String.Format("{0:N}", totalImporte.ToString())+ "</td>" +
                                    "<td>" + TotalImporte.ToString("C") + "</td>" +
                                "</tr>" +
                            "</table> <br/><br/>";
        }


        public void CrearContenidoCorreo()
        {
            ContenidoCorreo = "Buen día, estimado cliente, le recordamos que las siguientes facturas están por llegar a su plazo de pago,  le agradecemos su ayuda para la programación oportuna del mismo. ";
            ContenidoCorreo += "De igual manera anexamos el estado de cuenta";
            ContenidoCorreo += "<br/><br/>";
            ContenidoCorreo += "<table border=1>" +
                    "<tr>" +
                        "<th>Factura</th>" +
                        "<th>Fecha factura</th>" +
                        "<th>Fecha Vencimiento</th>" +
                        "<th>Importe</th>" +
                    "</tr>";
        }

        public string EnviarCorreo(SqlDataReader leer, string NombreCliente)
        {
            string respuesta = "";
            try
            {
                //string tabla;
                //tabla = "Buen día, estimado cliente, le recordamos que las siguientes facturas están por llegar a su plazo de pago,  le agradecemos su ayuda para la programación oportuna del mismo. ";
                //tabla += "De igual manera anexamos el estado de cuentas";
                //tabla += "<br/><br/>";
                //tabla += "<table border=1>" +
                //        "<tr>" +
                //            "<th>Factura</th>" +
                //            "<th>Fecha factura</th>" +
                //            "<th>Fecha Vencimiento</th>" +
                //            "<th>Importe</th>" +
                //        "</tr>";
                double totalImporte = 0;
                string correoDestino = "";
                MailMessage correo = new MailMessage();
                //while (leer.Read())
                //{
                //    //correoDestino = "finanzas@speedtoner.com.mx";
                //    correoDestino = "cobranza@speedtoner.com.mx";
                //    IdCliente = int.Parse(leer[0].ToString());
                //    //correo.Body += "<br>" + "FACTURAS" + "<br>" +
                //    //    "<table>" +
                //    //        "<tr>" +
                //    //            "<th>Factura</th>" +
                //    //            "<th>Fecha factura</th>" +
                //    //            "<th>Fecha Vencimiento</th>" + 
                //    //        "</tr>" +
                //    //Guardamos los valores obtenidos en variables para darles un formato en especifico a cada una
                //    DateTime FechaFactura = Convert.ToDateTime(leer[6].ToString());
                //    DateTime FechaVencimiento = Convert.ToDateTime(leer[7].ToString());
                //    double importe = double.Parse(leer[5].ToString());
                //    totalImporte += double.Parse(leer[5].ToString());
                //    //tabla += "<tr>" +
                //    //            "<td>" + leer[1] + "</td>" +
                //    //            "<td>" + FechaFactura.ToString("dd-MM-yyyy") + "</td>" +
                //    //            "<td>" + FechaVencimiento.ToString("dd-MM-yyyy") + "</td>" +
                //    //            "<td ALIGN=Right>" + importe.ToString("C") + "</td>" +
                //    //        "</tr>";

                //    //correo.Body += "<br/><br/>" +
                //    //    "Factura: "+leer[0] + "<br/>" + "Fecha factura: "+ FechaFactura.ToString("dd-MM-yyyy") + "<br/>"
                //    //    +"Fecha de vencimiento: "+ FechaVencimiento.ToString("dd-MM-yyyy") + "<br/>" + "Costo: $" + leer[5]; //Mensaje del correo
                //}
                leer.Close();
                correo.From = new MailAddress("cobranza@speedtoner.com.mx", "Departamento de Facturación y Cobranza", System.Text.Encoding.UTF8);//Correo de salida
                                                                                                                                                 //Buscamos en la base de datos todos los correos posibles que pueda tener ese cliente
                SqlDataReader correos = ObtenerCorreosCliente(IdCliente);
                while (correos.Read())
                {
                    correo.To.Add(correos[0].ToString()); //Correo destino?
                }
                //correo.To.Add("sistemas@speedtoner.com.mx");
                correos.Close();

                correo.Subject = "Recordatorio Cobranza para " + NombreCliente; //Asunto

                correo.IsBodyHtml = true;
                //tabla += "<tr>" +
                //                "<td>" + "</td>" +
                //                "<td>" + "</td>" +
                //                "<td>" + "TOTAL" + "</td>" +
                //                //"<td>" + String.Format("{0:N}", totalImporte.ToString())+ "</td>" +
                //                "<td>" + totalImporte.ToString("C") + "</td>" +
                //            "</tr>" +
                //        "</table> <br/><br/>";
                correo.IsBodyHtml = true;
                //tabla += "<h3 style=text-align:center;>SPEED TONER NUEVO LAREDO</h3>";
                //tabla += "<h3 style=text-align:center;>LEONA VICARIO NO.1709 NUEVO LAREDO, TAMPS</h3>";
                //tabla += "<h3 style=text-align:center;>TEL. 867-712-09-64 Y 867-190-0387 cobranza@speedtoner.com.mx sistemas@speedtoner.com.mx</h3>";

                ContenidoCorreo += "<h3 style=text-align:center;>SPEED TONER NUEVO LAREDO</h3>";
                ContenidoCorreo += "<h3 style=text-align:center;>LEONA VICARIO NO.1709 NUEVO LAREDO, TAMPS</h3>";
                ContenidoCorreo += "<h3 style=text-align:center;>TEL. 867-712-09-64 Y 867-190-0387 cobranza@speedtoner.com.mx sistemas@speedtoner.com.mx</h3>";

                correo.Body += ContenidoCorreo;

                //ESTO ES PARA PODER COLOCAR EL PDF EN ARCHIVOS ADJUNTOS DEL CORREO
                Attachment attachment = new Attachment(NombreArchivo);
                correo.Attachments.Add(attachment);

                //AlternateView VISTAHTML = AlternateView.CreateAlternateViewFromString(tabla, null/* TODO Change to default(_) if this is not a reference type */, System.Net.Mime.MediaTypeNames.Text.Html);
                //LinkedResource IMAGEN1 = new LinkedResource(@"C:\Users\DELL PC\Pictures\00_SPEED TONER LOGO JPG", System.Net.Mime.MediaTypeNames.Image.Jpeg);
                //IMAGEN1.ContentId = "IMG1";
                //VISTAHTML.LinkedResources.Add(IMAGEN1);
                //correo.AlternateViews.Add(VISTAHTML);
                AlternateView VISTAHTML = AlternateView.CreateAlternateViewFromString(ContenidoCorreo + "<div width=180px style=text-align:center;><img src=cid:Imagen1 width=180px></div>", null/* TODO Change to default(_) if this is not a reference type */, "text/html");
                string imagePathF = @"C:\Users\DELL PC\Pictures\00_SPEED TONER LOGO JPG.jpg";//aqui la ruta de tu imagen
                LinkedResource face = new LinkedResource(imagePathF, MediaTypeNames.Image.Jpeg);
                face.ContentId = "Imagen1";
                VISTAHTML.LinkedResources.Add(face);

                correo.AlternateViews.Add(VISTAHTML);

                correo.Priority = MailPriority.Normal;

                SmtpClient smtp = new SmtpClient();
                smtp.UseDefaultCredentials = false;
                smtp.Host = "smtp.office365.com"; //Host del servidor de correo
                smtp.Port = 587; //Puerto de salida

                const string quote = "\"";
                string contraseña = "C@rtur5dG5#23*";
                //smtp.Credentials = new System.Net.NetworkCredential("finanzas@speedtoner.com.mx", "C@rTr1dG3#22");//Cuenta de correo
                smtp.Credentials = new System.Net.NetworkCredential("cobranza@speedtoner.com.mx", "C@rtur5dG5#23*");//Cuenta de correo
                ServicePointManager.ServerCertificateValidationCallback = delegate (object s, X509Certificate certificate, X509Chain chain, SslPolicyErrors sslPolicyErrors) { return true; };
                smtp.EnableSsl = true;//True si el servidor de correo permite ssl
                smtp.Send(correo);
                leer.Close();
                return respuesta = "Correo enviado correctamente";
            }
            catch (Exception ex)
            {
                return "Correo no enviado, algo salio mal " + ex.Message;
            }

        }

        public SqlDataReader ObtenerCorreosCliente(int IdCliente)
        {
            SqlDataReader dr;
            comando.Connection = conexion.AbrirConexion();
            comando.CommandText = "BuscarCorreosDestino";
            comando.CommandType = CommandType.StoredProcedure;
            comando.Parameters.AddWithValue("@IdCliente", IdCliente);
            dr = comando.ExecuteReader();
            comando.Parameters.Clear();
            return dr;
        }
        #endregion



    }
}
