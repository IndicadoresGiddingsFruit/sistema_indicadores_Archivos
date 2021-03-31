using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Mail;
using System.Net.Mime;
using System.Text;
using System.Web;

namespace Sistema_Indicadores.Clases
{
    public class Email
    {
        public void sendmail(string correo_p = "", int region = 0, IEnumerable<ClassProductor> tabla = null, string Cod_Prod = "", string Productor = "", string Campo = "", string estatus = "")
        {
            string destinatario = "", subject= "" , estatus_tabla = "", estatus_header = "",mensaje="",link="";
            int dias = 0;
            if (tabla != null)
            {
                foreach (var item in tabla)
                {
                    if (item.dias != null)
                    {
                        dias = (int)item.dias;
                    }

                    if (item.Estatus != "\n<select class=\"form-control form-control-sm\" id=\"status\" name=\"item.Estatus\"><option>--Seleccione--</option>\n<option value=\"A\">ATENCIÓN A PRODUCTORES</option>\n<option value=\"M\">CIERRE DE MATERIAL</option>\n<option value=\"C\">COBRANZA</option>\n<option value=\"R\">PENDIENTE REVISIÓN</option>\n<option value=\"S\">SALDADO</option>\n<option value=\"T\">TERMINÓ TEMPORADA</option>\n<option value=\"E\">VA A ENTREGAR</option>\n<option value=\"P\">VA A PAGAR</option>\n</select>                                ")
                    {
                        estatus_tabla = item.Estatus;
                        estatus_header = "Estatus";

                        subject = "INFORMATIVO:: VALIDACION DE CARTERA VIGENTE";                      
                        mensaje = "Se han agregado nuevos registros con las siguientes validaciones <br/>";
                    }
                    else {
                        subject = "IMPORTANTE:: VALIDACION DE CARTERA VIGENTE";
                        mensaje = "Se necesita tu apoyo REALIZANDO LA VALIDACION  de la cartera en EL SIGUIENTE ACCESO, gracias. <br/>";
                        link= "http://www.giddingsfruit.mx/SistemaIndicadores/Indicadores/Seguimiento <br/>";
                    }
                }
            }

            MailMessage correo = new MailMessage();
            correo.From = new MailAddress("indicadores.giddingsfruit@gmail.com", "Indicadores GiddingsFruit");
            correo.To.Add(correo_p); //"marholy.martinez@giddingsfruit.mx" correo_p
            
            if (estatus == "A")
            {
                correo.CC.Add("oscar.castillo@giddingsfruit.mx");
                correo.CC.Add("angel.lopez@giddingsfruit.mx");
                subject = "IMPORTANTE:: VALIDACION DE CARTERA VIGENTE";
            }

            else
            {
                if (estatus == "R")
                {
                    subject = "RE: Dias sin respuesta "+dias+ " :: IMPORTANTE:: VALIDACION DE CARTERA VIGENTE";
                }
                
                destinatario = "Ingeniero";

                //correo.CC.Add("maria.orozco@giddingsfruit.mx");
                correo.CC.Add("oscar.castillo@giddingsfruit.mx");
                correo.CC.Add("angel.lopez@giddingsfruit.mx");

                if (correo_p == "aliberth.martinez@giddingsfruit.mx")
                {
                    correo.CC.Add("jose.acosta@giddingsfruit.mx");
                }

                if (region == 2)
                {
                    correo.CC.Add("jose.acosta@giddingsfruit.mx");
                    correo.CC.Add("jose.partida@giddingsfruit.mx");
                }

                else if (region == 3)
                {
                    correo.CC.Add("genaro.morales@giddingsfruit.mx");
                    correo.CC.Add("fernando.fierro@giddingsfruit.mx");
                }

                else if (region == 4)
                {
                    correo.CC.Add("josefina.cervantes@giddingsfruit.mx");
                }

                else if (region == 5)
                {
                    correo.CC.Add("evelyn.caceres@giddingsfruit.mx");
                }
            }

            correo.Subject = subject;
            correo.Body += "Buen dia " + destinatario + " <br/>";
            correo.Body += " <br/>";
            correo.Body += ""+mensaje+" <br/>";
            correo.Body += " <br/>";

            correo.Body += link;
            correo.Body += " <br/>";

            if (estatus != "A")
            {
                //correo.Body += "VIGENCIA DE LA LIGA DE 7 DIAS <br/>";
                //correo.Body += " <br/>";

                if (tabla != null)
                {                    
                    foreach (var item in tabla)
                    {
                        correo.Body += "<table border striped><thead><tr><th colspan='3'></th><th colspan='2'>Cajas entregadas</th></tr>" +
                        "<tr><th> Codigo </th>" +
                        "<th> Productor </th>" +
                        "<th> Campo </th>" +
                        "<th> Sem 27-52 </th>" +
                        "<th> Sem 1-actual </th>" +
                        "<th> "+ estatus_header + " </th>" +
                        "</tr></thead>";

                        correo.Body += "<tbody><tr>" +
                            "<td> " + item.Cod_Prod + " </td>" +
                            "<td> " + item.Productor + "</td>" +
                            "<td> " + item.Campo + " </td>" +
                            "<td> " + item.caja1 + " </td>" +
                            "<td> " + item.caja2 + " </td>" +
                            "<td> " + estatus_tabla + " </td>" +
                            "</tr></tbody>";
                    }
                    correo.Body += "</table>";
                }
                if (Cod_Prod != "")
                {
                    correo.Body += "<table border striped><thead><tr><td>Codigo</td><td> Productor </td><td> Campo </td></tr></thead><tbody><tr>" +
                            "<td> " + Cod_Prod + " </td>" +
                            "<td> " + Productor + "</td>" +
                            "<td> " + Campo + " </td>" +
                            "</tr></tbody>";
                    correo.Body += "</table>";
                }
                correo.Body += " <br/>";
            }

            correo.Body += "Saludos <br/>";
            correo.Body += " <br/>";
            correo.IsBodyHtml = true;
            correo.Priority = MailPriority.High;

            string sSmtpServer = "";
            sSmtpServer = "smtp.gmail.com";

            SmtpClient a = new SmtpClient();
            a.Host = sSmtpServer;
            a.Port = 587;//25
            a.EnableSsl = true;
            a.UseDefaultCredentials = true;
            a.Credentials = new System.Net.NetworkCredential
               ("indicadores.giddingsfruit@gmail.com", "indicadores2019");
            a.Send(correo);

            if (estatus_header == "")
            {
                credenciales(correo_p);
            }
        }

        public void credenciales(string correo_p)
        {
            try
            {
                string img_path = HttpContext.Current.Server.MapPath("/Image/Seguimiento_financ.png");
                MailMessage correo = new MailMessage();
                correo.From = new MailAddress("indicadores.giddingsfruit@gmail.com", "Indicadores GiddingsFruit");
                correo.To.Add(correo_p);//"marholy.martinez@giddingsfruit.mx"
                correo.Subject = "Credenciales";
                correo.Body += "Buen dia <br/>";
                correo.Body += " <br/>";
                correo.Body += "Para acceder al enlace anterior utilice su usuario de SEASON y la clave 123456 <br/>";
                correo.Body += " <br/>";
                correo.Body += "Seleccione un estatus, escribir un comentario, y presionar el boton azul para guardar cada uno <br/>";
                correo.Body += " <br/>";
                correo.Body += "Saludos! <br/>";
                correo.Body += " <br/>";
                correo.Attachments.Add(new Attachment(img_path));
                correo.IsBodyHtml = true;
                correo.Priority = MailPriority.High;
                string sSmtpServer = "";
                sSmtpServer = "smtp.gmail.com";
                SmtpClient a = new SmtpClient();
                a.Host = sSmtpServer;
                a.Port = 587;//25
                a.EnableSsl = true;
                a.UseDefaultCredentials = true;
                a.Credentials = new System.Net.NetworkCredential
                   ("indicadores.giddingsfruit@gmail.com", "indicadores2019");
                a.Send(correo);
            }
            catch (Exception e)
            {
                e.ToString();
            }
        }
    }
}