using Sistema_Indicadores.Models;
using System;
using System.Linq;
using System.Net.Mail;
using System.Web.Mvc;

namespace Sistema_Indicadores.Controllers
{
    public class LoginController : Controller
    {
        SeasonSun1213Entities bd = new SeasonSun1213Entities();
        SIPGUsuarios userModel = new SIPGUsuarios();

        // GET: Login
        public ActionResult Index(string mypass, string newpass)
        {
            try
            {
                if (mypass != null)
                {
                    string y = Session["Id"].ToString();
                    int id = Convert.ToInt32(y);
                    string mail = Session["Correo"].ToString();

                    //send_mail(newpass, id);

                    var verdadero = bd.SIPGUsuarios.FirstOrDefault(m => m.Clave == mypass && m.Id == id);
                    if (verdadero != null)
                    {
                        send_mail(newpass, id);
                    }
                    else
                    {
                        MailMessage correo = new MailMessage();
                        correo.From = new MailAddress("indicadores.giddingsfruit@gmail.com", "Indicadores GiddingsFruit");
                        correo.To.Add("marholy.martinez@giddingsfruit.mx");
                        correo.Subject = "Nueva contraseña";
                        correo.Body = "Estimado: " + Session["Nombre"].ToString() + " <br/>";
                        correo.Body += " <br/>";
                        correo.Body += "La clave no se actualizo, verifique porfavor <br/>";
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
                        a.Credentials = new System.Net.NetworkCredential("indicadores.giddingsfruit@gmail.com", "indicadores2019");
                        a.Send(correo);
                    }
                }
            }
            catch (Exception e)
            {
                e.ToString();
            }
            return View();
        }

        //public JsonResult GetPass(string mypass)
        //{
        //    try
        //    {
        //        string y = Session["Id"].ToString();
        //        int id = Convert.ToInt32(y);

        //        bd.Configuration.ProxyCreationEnabled = false;
        //        var result = bd.SIPGUsuarios.FirstOrDefault(m => m.Clave == mypass && m.Id == id);
        //        if (result != null)
        //        {
        //            return Json(result, JsonRequestBehavior.AllowGet);
        //        }
        //    }
        //    catch (Exception e) {
        //        e.ToString();
        //        return null;
        //    }
        //    return null;
        //}

        private void send_mail(string newpass, int id)
        {
            var item = bd.SIPGUsuarios.Where(x => x.Id == id).First();
            item.Clave = newpass;
            bd.SaveChanges();

            var res = bd.SIPGUsuarios.FirstOrDefault(m => m.Id == id);
            string direccion = res.correo;

            MailMessage correo = new MailMessage();
            correo.From = new MailAddress("indicadores.giddingsfruit@gmail.com", "Indicadores GiddingsFruit");
            correo.To.Add(direccion);
            correo.Subject = "Nueva contraseña";
            correo.Body = "Estimado: " + Session["Nombre"].ToString() + " <br/>";
            correo.Body += " <br/>";
            correo.Body += "Su nueva clave es: " + newpass + " <br/>";
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
            a.Credentials = new System.Net.NetworkCredential("indicadores.giddingsfruit@gmail.com", "indicadores2019");
            a.Send(correo);
            Session.Clear();
            Session.Abandon();
        }

        [HttpPost]
        public ActionResult Home(SIPGUsuarios userModel, string logout)
        {
            try
            {
                if (logout == null)
                {
                    var userDetails = bd.SIPGUsuarios.Where(x => x.Nombre == userModel.Nombre && x.Clave == userModel.Clave).FirstOrDefault();

                    if (userDetails == null)
                    {
                        ViewBag.error = "Datos Incorrectos, verifique porfavor";
                        return View("Index", userModel);
                    }
                    else
                    {
                        Session["IdAgen"] = userDetails.IdAgen;
                        Session["Tipo"] = userDetails.Tipo;
                        Session["Nombre"] = userDetails.Completo;
                        Session["Correo"] = userDetails.correo;
                        Session["Id"] = userDetails.Id;

                        if (userDetails.Id == 117)
                        {
                            return RedirectToAction("Zarzamora", "EstimacionBerries");
                        }
                        //oscar
                        else if (userDetails.Id == 391 || userDetails.Id == 188 || userDetails.Id == 44)
                        {
                            return RedirectToAction("Seguimiento", "Indicadores");
                        }                      
                        //activos curva
                        else if (userDetails.Id == 121)
                        {
                            return RedirectToAction("Activos_Curva", "Indicadores");
                        }
                        else if (userDetails.Tipo == null)
                        {
                            ViewBag.error = "Usuario no válido";
                            return View("Index", userModel);
                        }
                        else
                        {
                            if (userDetails.Tipo == "P" && userDetails.IdAgen != 1 && userDetails.IdAgen != 5)
                            {
                                return RedirectToAction("Seguimiento", "Indicadores");
                            }
                            else
                            {
                                return RedirectToAction("SetSolicitud", "Muestreo");
                            }
                        }
                    }
                }
                else
                {
                    System.Web.Security.FormsAuthentication.SignOut();
                    Session.Remove("Id");
                    Session.Remove("Nombre");
                    Session.Clear();
                    Session.Abandon();
                    return RedirectToAction("Index", "Login", false);
                }
            }
            catch (Exception e)
            {
                ViewBag.error = e.ToString();
                return null;
            }
        }
    }
}