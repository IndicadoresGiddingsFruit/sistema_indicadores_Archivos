using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Mail;
using System.Web;

namespace Sistema_Indicadores.Clases
{
    public class Email
    {
        public void sendmail(string correo_p, int region, IEnumerable<ClassProductor> tabla = null, string Cod_Prod = "", string Productor = "", string Campo= "")
        {          
            MailMessage correo = new MailMessage();
            correo.From = new MailAddress("indicadores.giddingsfruit@gmail.com", "Indicadores GiddingsFruit");
            correo.To.Add(correo_p);// "marholy.martinez@giddingsfruit.mx" correo_p
            correo.CC.Add("oscar.castillo@giddingsfruit.mx");
            correo.CC.Add("angel.lopez@giddingsfruit.mx");
            correo.CC.Add("maria.orozco@giddingsfruit.mx");

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

            correo.Subject = "IMPORTANTE:: VALIDACION DE CARTERA VIGENTE";
            correo.Body += "Buen dia Ingeniero <br/>";
            correo.Body += " <br/>";
            correo.Body += "Se necesita tu apoyo REALIZANDO LA VALIDACION  de la cartera en EL SIGUIENTE ACCESO, gracias. <br/>";
            correo.Body += " <br/>";
            correo.Body += "http://www.giddingsfruit.mx/SistemaIndicadores/Indicadores/Seguimiento <br/>";
            correo.Body += " <br/>";
            correo.Body += "VIGENCIA DE LA LIGA DE 7 DIAS <br/>";
            correo.Body += " <br/>";
            if (tabla != null)
            {
                correo.Body += "<table border striped><thead>" +
                    "<tr><td> Codigo </td>" +
                    "<td> Productor </td>" +
                    "<td> Campo </td>" +
                    "</tr></thead>";
                foreach (var item in tabla)
                {
                    correo.Body += "<tbody><tr>" +
                        "<td> " + item.Cod_Prod + " </td>" +
                        "<td> " + item.Productor + "</td>" +
                        "<td> " + item.Campo + " </td>" +
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

            credenciales(correo_p);
        }


        public void credenciales(string correo_p)
        {
            MailMessage correo = new MailMessage();
            correo.From = new MailAddress("indicadores.giddingsfruit@gmail.com", "Indicadores GiddingsFruit");
            correo.To.Add(correo_p);// "marholy.martinez@giddingsfruit.mx" correo_p           

            correo.Subject = "Credenciales";
            correo.Body += "Buen dia <br/>";
            correo.Body += " <br/>";
            correo.Body += "Para acceder al enlace anterior utilice su usuario de SEASON y la clave 123456 <br/>";
            correo.Body += " <br/>";
            correo.Body += "Saludos! <br/>";            
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
        }
    }
}