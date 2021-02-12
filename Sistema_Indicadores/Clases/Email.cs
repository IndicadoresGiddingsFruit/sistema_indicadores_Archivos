using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Mail;
using System.Web;

namespace Sistema_Indicadores.Clases
{
    public class Email
    {
        public void sendmail(string correo_p) 
        {
            MailMessage correo = new MailMessage();
            correo.From = new MailAddress("indicadores.giddingsfruit@gmail.com", "Indicadores GiddingsFruit");
            correo.To.Add(correo_p);// "marholy.martinez@giddingsfruit.mx" correo_p
            correo.CC.Add("oscar.castillo@giddingsfruit.mx");
            correo.CC.Add("angel.lopez@giddingsfruit.mx");
            correo.CC.Add("maria.orozco@giddingsfruit.mx");

            correo.Subject = "IMPORTANTE:: VALIDACION DE CARTERA VIGENTE";
            correo.Body += "Buen dia Ingeniero <br/>";
            correo.Body += " <br/>";
            correo.Body += "Se necesita tu apoyo REALIZANDO LA VALIDACION  de la cartera en EL SIGUIENTE ACCESO, gracias. <br/>";
            correo.Body += " <br/>";
            correo.Body += "http://www.giddingsfruit.mx/SistemaIndicadores/Indicadores/Seguimiento <br/>";
            correo.Body += " <br/>";
            correo.Body += "VIGENCIA DE LA LIGA DE 7 DIAS <br/>";
            correo.Body += " <br/>";
            correo.IsBodyHtml = true;
            correo.Priority = MailPriority.Normal;

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