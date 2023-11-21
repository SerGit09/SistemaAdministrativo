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


namespace CobranzaSP.Lógica
{
    internal class RemisionLogica
    {
        private CD_Conexion conexion = new CD_Conexion();
        SqlCommand comando = new SqlCommand();
        AccionesLógica NuevaAccion = new AccionesLógica();

        public string AgregarRemision(Remision nuevaRemision, string sp, bool agregandoNueva)
        {
            int respuesta = 0;
            string mensaje = "";
            bool RemisionDuplicada = false;
            comando.Connection = conexion.AbrirConexion();
            //Verificamos si estamos agregando alguna remision
            if (agregandoNueva)
            {
                RemisionDuplicada = NuevaAccion.VerificarDuplicados(nuevaRemision.NumeroFolio, "VerificarRemisionDuplicada");
                if (RemisionDuplicada)
                {
                    mensaje = "Factura duplicada";
                    return mensaje;
                }
            }
            comando.CommandText = sp;
            comando.CommandType = CommandType.StoredProcedure;

            comando.Parameters.AddWithValue("@NumeroFolio", nuevaRemision.NumeroFolio);
            comando.Parameters.AddWithValue("@FechaPago", nuevaRemision.FechaPago);
            comando.Parameters.AddWithValue("@DiasCredito", nuevaRemision.DiasCredito);
            comando.Parameters.AddWithValue("@Cantidad", nuevaRemision.Cantidad);
            comando.Parameters.AddWithValue("@IdCliente", nuevaRemision.IdCliente);

            respuesta = comando.ExecuteNonQuery();
            if (agregandoNueva)
            {
                mensaje = (respuesta > 0) ? "Registro agregado correctamente" : "Algo ha salido mal, no agrego el registro";
            }
            else
            {
                mensaje = (respuesta > 0) ? "Registro modificado correctamente" : "Algo ha salido mal, no se modifico el registro";
            }
            comando.Parameters.Clear();
            conexion.CerrarConexion();
            return mensaje;
        }

        public string RemisionPagada(CuentaPagada nuevaCuentaPagada)
        {
            int respuesta = 0;
            string mensaje = "";
            comando.Connection = conexion.AbrirConexion();
            comando.CommandText = "RemisionPagada";
            comando.CommandType = CommandType.StoredProcedure;
            comando.Parameters.AddWithValue("@NumeroFolio", nuevaCuentaPagada.Factura);
            comando.Parameters.AddWithValue("@IdCliente", nuevaCuentaPagada.IdCliente);
            comando.Parameters.AddWithValue("@Cantidad", nuevaCuentaPagada.Cantidad);
            comando.Parameters.AddWithValue("@FechaFactura", nuevaCuentaPagada.FechaFactura);
            comando.Parameters.AddWithValue("@TipoPago", "REMISION");

            respuesta = comando.ExecuteNonQuery();
            mensaje = (respuesta > 0) ? "Cuenta cobrada correctamente" : "Algo ha salido mal, no se pudo cobrar la remision";

            comando.Parameters.Clear();

            conexion.CerrarConexion();
            return mensaje;
        }

        public string GenerarCorreo(string NombreCliente)
        {
            SqlDataReader leer;
            comando.Connection = conexion.AbrirConexion();
            comando.CommandText = "DatosCorreoRemisiones";
            comando.CommandType = CommandType.StoredProcedure;
            comando.Parameters.AddWithValue("@NombreCliente", NombreCliente);
            leer = comando.ExecuteReader();
            comando.Parameters.Clear();
            return EnviarCorreo(leer, NombreCliente);
        }

        public string EnviarCorreo(SqlDataReader leer, string NombreCliente)
        {
            string respuesta = "";
            try
            {
                string tabla;
                tabla = "Buen día, estimado cliente, anexamos estado de cuenta de remisiones pendientes,  le agradecemos su ayuda para la programación oportuna del mismo. ";
                //tabla += "De igual manera le enviamos la relación de remisiones que presentan retraso en el pago acorde a la fecha de vencimiento.";
                tabla += "<br/><br/>";
                tabla += "<table border=1>" +
                        "<tr>" +
                            "<th>Remision</th>" +
                            "<th>Importe</th>" +
                            "<th>Fecha Pago</th>" +
                        "</tr>";
                double totalImporte = 0;
                int IdCliente = 0;
                string correoDestino = "";
                MailMessage correo = new MailMessage();
                while (leer.Read())
                {
                    //correoDestino = "finanzas@speedtoner.com.mx";
                    correoDestino = "cobranza@speedtoner.com.mx";
                    IdCliente = int.Parse(leer[1].ToString());
                    //correo.Body += "<br>" + "FACTURAS" + "<br>" +
                    //    "<table>" +
                    //        "<tr>" +
                    //            "<th>Factura</th>" +
                    //            "<th>Fecha factura</th>" +
                    //            "<th>Fecha Vencimiento</th>" + 
                    //        "</tr>" +
                    //Guardamos los valores obtenidos en variables para darles un formato en especifico a cada una
                    DateTime FechaPago = Convert.ToDateTime(leer[3].ToString());
                    double importe = double.Parse(leer[2].ToString());
                    totalImporte += double.Parse(leer[2].ToString());
                    tabla += "<tr>" +
                                "<td>" + leer[0] + "</td>" +
                                "<td>" + FechaPago.ToString("dd-MM-yyyy") + "</td>" +
                                "<td ALIGN=Right>" + importe.ToString("C") + "</td>" +
                            "</tr>";

                    //correo.Body += "<br/><br/>" +
                    //    "Factura: "+leer[0] + "<br/>" + "Fecha factura: "+ FechaFactura.ToString("dd-MM-yyyy") + "<br/>"
                    //    +"Fecha de vencimiento: "+ FechaVencimiento.ToString("dd-MM-yyyy") + "<br/>" + "Costo: $" + leer[5]; //Mensaje del correo
                }
                leer.Close();
                correo.From = new MailAddress("cobranza@speedtoner.com.mx", "Departamento de Facturación y Cobranza", System.Text.Encoding.UTF8);//Correo de salida
                                                                                                                                                 //Buscamos en la base de datos todos los correos posibles que pueda tener ese cliente
                SqlDataReader correos = ObtenerCorreosCliente(IdCliente);
                while (correos.Read())
                {
                    correo.To.Add(correos[0].ToString()); //Correo destino?
                }
                correo.To.Add("sistemas@speedtoner.com.mx");
                correos.Close();

                correo.Subject = "Recordatorio Pago de Remisiones para " + NombreCliente; //Asunto

                correo.IsBodyHtml = true;
                tabla += "<tr>" +
                                "<td>" + "</td>" +
                                "<td>" + "TOTAL" + "</td>" +
                                //"<td>" + String.Format("{0:N}", totalImporte.ToString())+ "</td>" +
                                "<td>" + totalImporte.ToString("C") + "</td>" +
                            "</tr>" +
                        "</table> <br/><br/>";
                correo.IsBodyHtml = true;
                tabla += "<h3 style=text-align:center;>SPEED TONER NUEVO LAREDO</h3>";
                tabla += "<h3 style=text-align:center;>LEONA VICARIO NO.1709 NUEVO LAREDO, TAMPS</h3>";
                tabla += "<h3 style=text-align:center;>TEL. 867-712-09-64 Y 867-190-0387 cobranza@speedtoner.com.mx sistemas@speedtoner.com.mx</h3>";
                correo.Body += tabla;

                //AlternateView VISTAHTML = AlternateView.CreateAlternateViewFromString(tabla, null/* TODO Change to default(_) if this is not a reference type */, System.Net.Mime.MediaTypeNames.Text.Html);
                //LinkedResource IMAGEN1 = new LinkedResource(@"C:\Users\DELL PC\Pictures\00_SPEED TONER LOGO JPG", System.Net.Mime.MediaTypeNames.Image.Jpeg);
                //IMAGEN1.ContentId = "IMG1";
                //VISTAHTML.LinkedResources.Add(IMAGEN1);
                //correo.AlternateViews.Add(VISTAHTML);
                AlternateView VISTAHTML = AlternateView.CreateAlternateViewFromString(tabla + "<div width=180px style=text-align:center;><img src=cid:Imagen1 width=180px></div>", null/* TODO Change to default(_) if this is not a reference type */, "text/html");
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
                smtp.Credentials = new System.Net.NetworkCredential("cobranza@speedtoner.com.mx", "C@rtur1dG5#23*");//Cuenta de correo
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
    }
}
