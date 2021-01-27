using Sistema_Indicadores.Clases;
using Sistema_Indicadores.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Mail;
using System.Web.Mvc;

namespace Sistema_Indicadores.Controllers
{
    public class MuestreoController : Controller
    {
        SeasonSun1213Entities15 bd = new SeasonSun1213Entities15();
        ProdMuestreo ProdMuestreo = new ProdMuestreo();
        ProdCamposCat ProdCamposCat = new ProdCamposCat();
        ProdProductoresCat ProdProductoresCat = new ProdProductoresCat();
        ProdAnalisis_Residuo ProdAnalisis_Residuo = new ProdAnalisis_Residuo();
        ProdMuestreoSector ProdMuestreoSector = new ProdMuestreoSector();

        string correo_p = "", correo_c = "", correo_i = "", compras_oportunidad = "", subject = "", body_email = "";

        public ActionResult Muestreo()
        {
            if (Session["Nombre"] != null)
            {
                ViewData["Nombre"] = Session["Nombre"].ToString();
            }
            else
            {
                return RedirectToAction("Index", "Login");
            }
            return View();
        }

        //Nuevo muestreo
        [HttpPost]
        public ActionResult Muestreo(ProdMuestreo model)
        {
            try
            {
                ViewData["Nombre"] = Session["Nombre"].ToString();
                var modeloExistente = bd.ProdMuestreo.FirstOrDefault(m => m.Cod_Prod == model.Cod_Prod && m.Cod_Campo == model.Cod_Campo);
                if (modeloExistente == null)
                {
                    var idAgen = bd.ProdCamposCat.Where(x => x.Cod_Prod == model.Cod_Prod && x.Cod_Campo == model.Cod_Campo).First();
                    string tipo = Convert.ToString(Session["Tipo"]);

                    ProdMuestreo.IdAgen = idAgen.IdAgen;
                    ProdMuestreo.Cod_Empresa = 2;
                    ProdMuestreo.Cod_Prod = model.Cod_Prod;
                    ProdMuestreo.Cod_Campo = model.Cod_Campo;
                    ProdMuestreo.Fecha_solicitud = DateTime.Now;
                    ProdMuestreo.Telefono = model.Telefono;
                    ProdMuestreo.Inicio_cosecha = model.Inicio_cosecha;

                    if (model.Compras_Oportunidad == "on")
                    {
                        compras_oportunidad = "S";
                    }
                    else
                    {
                        compras_oportunidad = null;
                    }
                    var camposcat = bd.ProdCamposCat.Where(x => x.Cod_Prod == model.Cod_Prod && x.Cod_Campo == model.Cod_Campo).First();
                    camposcat.Compras_Oportunidad = compras_oportunidad;
                    bd.SaveChanges();

                    if (tipo == "P")
                    {
                        ProdMuestreo.Liberacion = "S";
                    }

                    bd.ProdMuestreo.Add(ProdMuestreo);
                    bd.SaveChanges();

                    var campo = bd.ProdCamposCat.FirstOrDefault(m => m.Cod_Prod == model.Cod_Prod && m.Cod_Campo == model.Cod_Campo);
                    var email_p = bd.SIPGUsuarios.FirstOrDefault(m => m.IdAgen == campo.IdAgen && m.Tipo == "P");
                    var email_c = bd.SIPGUsuarios.FirstOrDefault(m => m.IdAgen == campo.IdAgenC && m.Tipo == "C");
                    var email_i = bd.SIPGUsuarios.FirstOrDefault(m => m.IdAgen == campo.IdAgenI && m.Tipo == "I");

                    correo_p = email_p.correo;

                    if (email_c == null)
                    {
                        var item = bd.ProdCamposCat.Where(x => x.Cod_Prod == model.Cod_Prod && x.Cod_Campo == model.Cod_Campo).First();
                        item.IdAgenC = 167;
                        bd.SaveChanges();
                        correo_c = "mayra.ramirez@giddingsfruit.mx";
                    }
                    else
                    {
                        correo_c = email_c.correo;
                    }

                    if (email_i == null)
                    {
                        var item = bd.ProdCamposCat.Where(x => x.Cod_Prod == model.Cod_Prod && x.Cod_Campo == model.Cod_Campo).First();
                        item.IdAgenI = 205;
                        bd.SaveChanges();
                        correo_i = "jesus.palafox@giddingsfruit.mx";
                    }
                    else
                    {
                        correo_i = email_i.correo;
                    }

                    //correo
                    try
                    {
                        var prod = bd.ProdProductoresCat.FirstOrDefault(x => x.Cod_Prod == campo.Cod_Prod);
                        var Inicio_cosecha = String.Format("{0:d}", model.Inicio_cosecha);

                        MailMessage correo = new MailMessage();
                        correo.From = new MailAddress("indicadores.giddingsfruit@gmail.com", "Indicadores GiddingsFruit");

                        if (Session["Tipo"].ToString() == "P")
                        {
                            correo.To.Add(Session["Correo"].ToString());
                            correo.CC.Add(correo_c);
                            correo.CC.Add(correo_i);
                        }
                        else if (Session["Tipo"].ToString() == "I")
                        {
                            correo.To.Add(Session["Correo"].ToString());
                            correo.CC.Add(correo_c);
                            correo.CC.Add(correo_p);
                        }
                        else if (Session["Tipo"].ToString() == "C")
                        {
                            correo.To.Add(Session["Correo"].ToString());
                            correo.CC.Add(correo_i);
                            correo.CC.Add(correo_p);
                            correo.CC.Add("marco.velazquez@giddingsfruit.mx");
                        }

                        correo.CC.Add("oscar.castillo@giddingsfruit.mx");

                        correo.Subject = "Nuevo Muestreo: " + campo.Cod_Prod;
                        correo.Body = "Solicitado por: " + Session["Nombre"].ToString() + " <br/>";
                        correo.Body += " <br/>";
                        correo.Body += "Productor: " + campo.Cod_Prod + " - " + prod.Nombre + " <br/>";
                        correo.Body += " <br/>";
                        correo.Body += "Campo: " + campo.Cod_Campo + " - " + campo.Descripcion + " <br/>";
                        correo.Body += " <br/>";
                        correo.Body += "Telefono: " + model.Telefono + "<br/>";
                        correo.Body += " <br/>";
                        correo.Body += "Inicio de cosecha: " + Inicio_cosecha + "<br/>";

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
                    catch (Exception e) { e.ToString(); }

                    TempData["sms"] = "Solicitud realizada con éxito, a partir de hoy tendrá 5 días hábiles para realizar muestreo";
                    ViewBag.sms = TempData["sms"].ToString();
                }
                else
                {
                    TempData["error"] = "La solicitud ya fue realizada anteriormente, verifique porfavor";
                    ViewBag.error = TempData["error"].ToString();
                }
            }
            catch (Exception e)
            {
                e.ToString();
                TempData["error"] = e.ToString();
                ViewBag.error = TempData["error"].ToString();
            }
            return View();
        }

        public void email()
        {
            try { }
            catch (Exception e)
            {
                e.ToString();
            }
        }

        //Buscar
        public ActionResult SetSolicitud(string BuscarCod_Prod = "", short agente = 0, string status = "", string pendientes = "")
        {
            if (Session["Nombre"] != null)
            {
                ViewData["Nombre"] = Session["Nombre"].ToString();

                List<SelectListItem> lst_Status = new List<SelectListItem>();
                lst_Status.Add(new SelectListItem() { Text = "--Resultado del analisis--", Value = null });
                lst_Status.Add(new SelectListItem() { Text = "CON RESIDUOS", Value = "R" });
                lst_Status.Add(new SelectListItem() { Text = "EN PROCESO", Value = "P" });
                lst_Status.Add(new SelectListItem() { Text = "FUERA DE LIMITE", Value = "F" });
                lst_Status.Add(new SelectListItem() { Text = "LIBERADO", Value = "L" });

                ViewBag.List_Status = lst_Status;

                List<SelectListItem> lst_Tarjeta = new List<SelectListItem>();
                lst_Tarjeta.Add(new SelectListItem() { Text = "--Tarjeta--", Value = null });
                lst_Tarjeta.Add(new SelectListItem() { Text = "LIBERADA", Value = "S" });
                lst_Tarjeta.Add(new SelectListItem() { Text = "NO LIBERADA", Value = "N" });
                lst_Tarjeta.Add(new SelectListItem() { Text = "PENDIENTE", Value = null });

                ViewBag.List_Tarjeta = lst_Tarjeta;

                List<ProdAgenteCat> agentesI = new List<ProdAgenteCat>();
                agentesI = bd.ProdAgenteCat.Where(x => x.Depto == "I" && x.Activo == true).OrderBy(x => x.Nombre).ToList();
                ViewBag.agentesI = agentesI;

                List<ProdAgenteCat> agentesC = new List<ProdAgenteCat>();
                agentesC = bd.ProdAgenteCat.Where(x => x.Depto == "C" && x.Activo == true && x.Codigo != null).OrderBy(x => x.Nombre).ToList();
                ViewBag.agentesC = agentesC;

                List<ProdAgenteCat> agentesP = new List<ProdAgenteCat>();
                agentesP = bd.ProdAgenteCat.Where(x => x.Depto == "P" && x.Activo == true).OrderBy(x => x.Nombre).ToList();
                ViewBag.agentesP = agentesP;

                short agenteSesion = (short)Session["IdAgen"];
                IQueryable<ClassMuestreo> item = null;

                if (Session["IdAgen"].ToString() == "205" && pendientes == "")
                {

                    item = (from m in (from m in bd.ProdAnalisis_Residuo
                                       group m by new
                                       {
                                           Cod_Empresa = m.Cod_Empresa,
                                           Cod_Prod = m.Cod_Prod,
                                           Cod_Campo = m.Cod_Campo,
                                           IdSector = m.IdSector
                                       } into x
                                       select new
                                       {
                                           Cod_Empresa = x.Key.Cod_Empresa,
                                           Cod_Prod = x.Key.Cod_Prod,
                                           Cod_Campo = x.Key.Cod_Campo,
                                           IdSector = x.Key.IdSector,
                                           Fecha = x.Max(m => m.Fecha)
                                       }) 

                            join an in bd.ProdAnalisis_Residuo on new { m.Cod_Empresa, m.Cod_Prod, m.Cod_Campo, m.IdSector, m.Fecha } equals new { an.Cod_Empresa, an.Cod_Prod, an.Cod_Campo, an.IdSector, an.Fecha } into AnalisisR
                            from an in AnalisisR.DefaultIfEmpty()

                            join c in bd.ProdCamposCat on new { m.Cod_Empresa, m.Cod_Prod, m.Cod_Campo } equals new { c.Cod_Empresa, c.Cod_Prod, c.Cod_Campo } into MuestreoCam
                            from mcam in MuestreoCam.DefaultIfEmpty()

                            join p in bd.ProdProductoresCat on mcam.Cod_Prod equals p.Cod_Prod into MuestreoProd
                            from prod in MuestreoProd.DefaultIfEmpty()

                            join s in bd.ProdMuestreoSector on m.IdSector equals s.id into MuestreoSc
                            from ms in MuestreoSc.DefaultIfEmpty()

                            join r in bd.ProdMuestreo on an.Id_Muestreo equals r.Id into MuestreoAn
                            from man in MuestreoAn.DefaultIfEmpty()

                            join a in bd.ProdAgenteCat on mcam.IdAgen equals a.IdAgen into MuestreoAgentes
                            from ageP in MuestreoAgentes.DefaultIfEmpty()

                            join cf in bd.ProdCalidadMuestreo on man.Id equals cf.Id_Muestreo into MuestreoCa
                            from mc in MuestreoCa.DefaultIfEmpty()

                            join a in bd.ProdAgenteCat on mc.IdAgen equals a.IdAgen into MuestreoAgenC
                            from ageC in MuestreoAgenC.DefaultIfEmpty()

                            join a in bd.ProdAgenteCat on mcam.IdAgenC equals a.IdAgen into MuestreoAgenSC
                            from ageCS in MuestreoAgenSC.DefaultIfEmpty()

                            join l in bd.CatLocalidades on mcam.CodLocalidad equals l.CodLocalidad into MuestreoLoc
                            from loc in MuestreoLoc.DefaultIfEmpty()

                            where an.Estatus != "L" 

                            group m by new
                            {
                                IdAnalisis_Residuo = an.Id,
                                IdMuestreo = man.Id,
                                Asesor = ageP.Abrev,
                                Cod_Prod = m.Cod_Prod,
                                Productor = prod.Nombre,
                                Cod_Campo = m.Cod_Campo,
                                Campo = mcam.Descripcion,
                                Sector = (short)ms.Sector,
                                Ha = mcam.Hectareas,
                                Compras_oportunidad = mcam.Compras_Oportunidad,
                                Fecha_solicitud = (DateTime)man.Fecha_solicitud,
                                Inicio_cosecha = (DateTime)man.Inicio_cosecha,
                                Ubicacion = loc.Descripcion,
                                Telefono = man.Telefono,
                                Liberacion = man.Liberacion,
                                Fecha_ejecucion = (DateTime)man.Fecha_ejecucion,
                                Analisis = an.Estatus,
                                Calidad_fruta = mc.Estatus,
                                IdAgenC = (short)mcam.IdAgenC,
                                AsesorC = ageC.Abrev,
                                AsesorCS = ageCS.Abrev,
                                Tarjeta = man.Tarjeta,
                                IdRegion = ageP.IdRegion,
                                Fecha_analisis = m.Fecha
                            } into x
                            select new ClassMuestreo()
                            {
                                IdAnalisis_Residuo = x.Key.IdAnalisis_Residuo,
                                IdMuestreo = x.Key.IdMuestreo,
                                Asesor = x.Key.Asesor,
                                Cod_Prod = x.Key.Cod_Prod,
                                Productor = x.Key.Productor,
                                Cod_Campo = x.Key.Cod_Campo,
                                Campo = x.Key.Campo,
                                Sector = (short)x.Key.Sector,
                                Ha = x.Key.Ha,
                                Compras_oportunidad = x.Key.Compras_oportunidad,
                                Fecha_solicitud = (DateTime)x.Key.Fecha_solicitud,
                                Inicio_cosecha = (DateTime)x.Key.Inicio_cosecha,
                                Ubicacion = x.Key.Ubicacion,
                                Telefono = x.Key.Telefono,
                                Liberacion = x.Key.Liberacion,
                                Fecha_ejecucion = (DateTime)x.Key.Fecha_ejecucion,
                                Analisis = x.Key.Analisis,
                                Calidad_fruta = x.Key.Calidad_fruta,
                                IdAgenC = (short)x.Key.IdAgenC,
                                AsesorC = x.Key.AsesorC,
                                AsesorCS = x.Key.AsesorCS,
                                Tarjeta = x.Key.Tarjeta,
                                IdRegion = x.Key.IdRegion,
                                Fecha_analisis = x.Key.Fecha_analisis
                            }).Distinct();
                }
                else if (Session["IdAgen"].ToString() == "205" && pendientes != "")
                {
                    item = from m in bd.ProdMuestreo

                           join a in bd.ProdAgenteCat on m.IdAgen equals a.IdAgen into MuestreoAgentes
                           from ageP in MuestreoAgentes.DefaultIfEmpty()

                           join p in bd.ProdProductoresCat on m.Cod_Prod equals p.Cod_Prod into MuestreoProd
                           from prod in MuestreoProd.DefaultIfEmpty()

                           join r in bd.ProdAnalisis_Residuo on m.Id equals r.Id_Muestreo into MuestreoAn
                           from man in MuestreoAn.DefaultIfEmpty()

                           join s in bd.ProdMuestreoSector on m.IdSector equals s.id into MuestreoSc
                           from ms in MuestreoSc.DefaultIfEmpty()

                           join cf in bd.ProdCalidadMuestreo on m.Id equals cf.Id_Muestreo into MuestreoCa
                           from mc in MuestreoCa.DefaultIfEmpty()

                           join c in bd.ProdCamposCat on new { m.Cod_Empresa, m.Cod_Prod, m.Cod_Campo } equals new { c.Cod_Empresa, c.Cod_Prod, c.Cod_Campo } into MuestreoCam
                           from mcam in MuestreoCam.DefaultIfEmpty()

                           join a in bd.ProdAgenteCat on mc.IdAgen equals a.IdAgen into MuestreoAgenC
                           from ageC in MuestreoAgenC.DefaultIfEmpty()

                           join a in bd.ProdAgenteCat on mcam.IdAgenC equals a.IdAgen into MuestreoAgenSC
                           from ageCS in MuestreoAgenSC.DefaultIfEmpty()

                           join l in bd.CatLocalidades on mcam.CodLocalidad equals l.CodLocalidad into MuestreoLoc
                           from loc in MuestreoLoc.DefaultIfEmpty()

                           where man.Fecha == (from c in bd.ProdAnalisis_Residuo where c.Cod_Prod == man.Cod_Prod select c).Max(c => c.Fecha) && man.Estatus == null
                           select new ClassMuestreo
                           {
                               IdMuestreo = m.Id,
                               Asesor = ageP.Abrev,
                               Cod_Prod = m.Cod_Prod,
                               Productor = prod.Nombre,
                               Cod_Campo = m.Cod_Campo,
                               Campo = mcam.Descripcion,
                               Sector = (short)ms.Sector,
                               Ha = mcam.Hectareas,
                               Compras_oportunidad = mcam.Compras_Oportunidad,
                               Fecha_solicitud = (DateTime)m.Fecha_solicitud,
                               Inicio_cosecha = (DateTime)m.Inicio_cosecha,
                               Ubicacion = loc.Descripcion,
                               Telefono = m.Telefono,
                               Liberacion = m.Liberacion,
                               Fecha_ejecucion = (DateTime)m.Fecha_ejecucion,
                               IdAnalisis_Residuo = man.Id,
                               Analisis = man.Estatus,
                               Calidad_fruta = mc.Estatus,
                               IdAgenC = (short)mcam.IdAgenC,
                               AsesorC = ageC.Abrev,
                               AsesorCS = ageCS.Abrev,
                               Tarjeta = m.Tarjeta,
                               IdRegion = ageP.IdRegion,
                               Fecha_analisis = man.Fecha
                           };
                }
                else if (Session["IdAgen"].ToString() == "153" || Session["IdAgen"].ToString() == "281" || Session["IdAgen"].ToString() == "167" || Session["IdAgen"].ToString() == "182")
                {
                    item = (from m in (from m in bd.ProdMuestreo
                                      group m by new
                                           {
                                               Cod_Empresa = m.Cod_Empresa,
                                               Cod_Prod = m.Cod_Prod,
                                               Cod_Campo = m.Cod_Campo,
                                               IdSector = m.IdSector
                                           } into x
                                           select new
                                           {
                                               Cod_Empresa = x.Key.Cod_Empresa,
                                               Cod_Prod = x.Key.Cod_Prod,
                                               Cod_Campo = x.Key.Cod_Campo,
                                               IdSector = x.Key.IdSector,
                                               Fecha_solicitud = x.Max(m => m.Fecha_solicitud)
                                           })

                           join an in bd.ProdMuestreo on new { m.Cod_Empresa, m.Cod_Prod, m.Cod_Campo, m.IdSector, m.Fecha_solicitud } equals new { an.Cod_Empresa, an.Cod_Prod, an.Cod_Campo, an.IdSector, an.Fecha_solicitud } into MuestreoR
                           from an in MuestreoR.DefaultIfEmpty()

                           join c in bd.ProdCamposCat on new { m.Cod_Empresa, m.Cod_Prod, m.Cod_Campo } equals new { c.Cod_Empresa, c.Cod_Prod, c.Cod_Campo } into MuestreoCam
                           from mcam in MuestreoCam.DefaultIfEmpty()

                           join p in bd.ProdProductoresCat on mcam.Cod_Prod equals p.Cod_Prod into MuestreoProd
                           from prod in MuestreoProd.DefaultIfEmpty()

                           join s in bd.ProdMuestreoSector on m.IdSector equals s.id into MuestreoSc
                           from ms in MuestreoSc.DefaultIfEmpty()

                           join r in bd.ProdAnalisis_Residuo on an.Id equals r.Id_Muestreo into MuestreoAn
                           from man in MuestreoAn.DefaultIfEmpty()

                           join a in bd.ProdAgenteCat on an.IdAgen equals a.IdAgen into MuestreoAgentes
                           from ageP in MuestreoAgentes.DefaultIfEmpty()

                           join cf in bd.ProdCalidadMuestreo on an.Id equals cf.Id_Muestreo into MuestreoCa
                           from mc in MuestreoCa.DefaultIfEmpty()                                                    

                           join a in bd.ProdAgenteCat on mc.IdAgen equals a.IdAgen into MuestreoAgenC
                           from ageC in MuestreoAgenC.DefaultIfEmpty()

                           join a in bd.ProdAgenteCat on mcam.IdAgenC equals a.IdAgen into MuestreoAgenSC
                           from ageCS in MuestreoAgenSC.DefaultIfEmpty()

                           join l in bd.CatLocalidades on mcam.CodLocalidad equals l.CodLocalidad into MuestreoLoc
                           from loc in MuestreoLoc.DefaultIfEmpty()

                           where man.Estatus != "L" 
                           select new ClassMuestreo
                           {
                               IdAnalisis_Residuo = man.Id,
                               IdMuestreo = an.Id,
                               Asesor = ageP.Abrev,
                               Cod_Prod = m.Cod_Prod,
                               Productor = prod.Nombre,
                               Cod_Campo = m.Cod_Campo,
                               Campo = mcam.Descripcion,
                               Sector = (short)ms.Sector,
                               Ha = mcam.Hectareas,
                               Compras_oportunidad = mcam.Compras_Oportunidad,
                               Fecha_solicitud = (DateTime)m.Fecha_solicitud,
                               Inicio_cosecha = (DateTime)an.Inicio_cosecha,
                               Ubicacion = loc.Descripcion,
                               Telefono = an.Telefono,
                               Liberacion = an.Liberacion,
                               Fecha_ejecucion = (DateTime)an.Fecha_ejecucion,
                               Analisis = man.Estatus,
                               Calidad_fruta = mc.Estatus,
                               IdAgenC = (short)mcam.IdAgenC,
                               AsesorC = ageC.Abrev,
                               AsesorCS = ageCS.Abrev,
                               Tarjeta = an.Tarjeta,
                               IdRegion = ageP.IdRegion,
                               Fecha_analisis = man.Fecha
                           }).Distinct();
                }
                else if (Session["IdAgen"].ToString() == "1")
                {
                    item = (from m in (from m in bd.ProdMuestreo
                                       group m by new
                                       {
                                           Cod_Empresa = m.Cod_Empresa,
                                           Cod_Prod = m.Cod_Prod,
                                           Cod_Campo = m.Cod_Campo,
                                           IdSector = m.IdSector
                                       } into x
                                       select new
                                       {
                                           Cod_Empresa = x.Key.Cod_Empresa,
                                           Cod_Prod = x.Key.Cod_Prod,
                                           Cod_Campo = x.Key.Cod_Campo,
                                           IdSector = x.Key.IdSector,
                                           Fecha_solicitud = x.Max(m => m.Fecha_solicitud)
                                       })

                            join an in bd.ProdMuestreo on new { m.Cod_Empresa, m.Cod_Prod, m.Cod_Campo, m.IdSector, m.Fecha_solicitud } equals new { an.Cod_Empresa, an.Cod_Prod, an.Cod_Campo, an.IdSector, an.Fecha_solicitud } into MuestreoR
                            from an in MuestreoR.DefaultIfEmpty()

                            join c in bd.ProdCamposCat on new { m.Cod_Empresa, m.Cod_Prod, m.Cod_Campo } equals new { c.Cod_Empresa, c.Cod_Prod, c.Cod_Campo } into MuestreoCam
                            from mcam in MuestreoCam.DefaultIfEmpty()

                            join p in bd.ProdProductoresCat on mcam.Cod_Prod equals p.Cod_Prod into MuestreoProd
                            from prod in MuestreoProd.DefaultIfEmpty()

                            join s in bd.ProdMuestreoSector on m.IdSector equals s.id into MuestreoSc
                            from ms in MuestreoSc.DefaultIfEmpty()

                            join r in bd.ProdAnalisis_Residuo on an.Id equals r.Id_Muestreo into MuestreoAn
                            from man in MuestreoAn.DefaultIfEmpty()

                            join a in bd.ProdAgenteCat on an.IdAgen equals a.IdAgen into MuestreoAgentes
                            from ageP in MuestreoAgentes.DefaultIfEmpty()

                            join cf in bd.ProdCalidadMuestreo on an.Id equals cf.Id_Muestreo into MuestreoCa
                            from mc in MuestreoCa.DefaultIfEmpty()

                            join a in bd.ProdAgenteCat on mc.IdAgen equals a.IdAgen into MuestreoAgenC
                            from ageC in MuestreoAgenC.DefaultIfEmpty()

                            join a in bd.ProdAgenteCat on mcam.IdAgenC equals a.IdAgen into MuestreoAgenSC
                            from ageCS in MuestreoAgenSC.DefaultIfEmpty()

                            join l in bd.CatLocalidades on mcam.CodLocalidad equals l.CodLocalidad into MuestreoLoc
                            from loc in MuestreoLoc.DefaultIfEmpty()

                            where man.Estatus != "L" && (ageP.IdRegion == 1 || ageP.IdRegion == 3 || ageP.IdRegion == 4 || ageP.IdRegion == 5)
                            select new ClassMuestreo
                            {
                                IdAnalisis_Residuo = man.Id,
                                IdMuestreo = an.Id,
                                Asesor = ageP.Abrev,
                                Cod_Prod = m.Cod_Prod,
                                Productor = prod.Nombre,
                                Cod_Campo = m.Cod_Campo,
                                Campo = mcam.Descripcion,
                                Sector = (short)ms.Sector,
                                Ha = mcam.Hectareas,
                                Compras_oportunidad = mcam.Compras_Oportunidad,
                                Fecha_solicitud = (DateTime)m.Fecha_solicitud,
                                Inicio_cosecha = (DateTime)an.Inicio_cosecha,
                                Ubicacion = loc.Descripcion,
                                Telefono = an.Telefono,
                                Liberacion = an.Liberacion,
                                Fecha_ejecucion = (DateTime)an.Fecha_ejecucion,
                                Analisis = man.Estatus,
                                Calidad_fruta = mc.Estatus,
                                IdAgenC = (short)mcam.IdAgenC,
                                AsesorC = ageC.Abrev,
                                AsesorCS = ageCS.Abrev,
                                Tarjeta = an.Tarjeta,
                                IdRegion = ageP.IdRegion,
                                Fecha_analisis = man.Fecha
                            }).Distinct();                  
                }
                else if (Session["IdAgen"].ToString() == "5")
                {
                    item = (from m in (from m in bd.ProdMuestreo
                                       group m by new
                                       {
                                           Cod_Empresa = m.Cod_Empresa,
                                           Cod_Prod = m.Cod_Prod,
                                           Cod_Campo = m.Cod_Campo,
                                           IdSector = m.IdSector
                                       } into x
                                       select new
                                       {
                                           Cod_Empresa = x.Key.Cod_Empresa,
                                           Cod_Prod = x.Key.Cod_Prod,
                                           Cod_Campo = x.Key.Cod_Campo,
                                           IdSector = x.Key.IdSector,
                                           Fecha_solicitud = x.Max(m => m.Fecha_solicitud)
                                       })

                            join an in bd.ProdMuestreo on new { m.Cod_Empresa, m.Cod_Prod, m.Cod_Campo, m.IdSector, m.Fecha_solicitud } equals new { an.Cod_Empresa, an.Cod_Prod, an.Cod_Campo, an.IdSector, an.Fecha_solicitud } into MuestreoR
                            from an in MuestreoR.DefaultIfEmpty()

                            join c in bd.ProdCamposCat on new { m.Cod_Empresa, m.Cod_Prod, m.Cod_Campo } equals new { c.Cod_Empresa, c.Cod_Prod, c.Cod_Campo } into MuestreoCam
                            from mcam in MuestreoCam.DefaultIfEmpty()

                            join p in bd.ProdProductoresCat on mcam.Cod_Prod equals p.Cod_Prod into MuestreoProd
                            from prod in MuestreoProd.DefaultIfEmpty()

                            join s in bd.ProdMuestreoSector on m.IdSector equals s.id into MuestreoSc
                            from ms in MuestreoSc.DefaultIfEmpty()

                            join r in bd.ProdAnalisis_Residuo on an.Id equals r.Id_Muestreo into MuestreoAn
                            from man in MuestreoAn.DefaultIfEmpty()

                            join a in bd.ProdAgenteCat on an.IdAgen equals a.IdAgen into MuestreoAgentes
                            from ageP in MuestreoAgentes.DefaultIfEmpty()

                            join cf in bd.ProdCalidadMuestreo on an.Id equals cf.Id_Muestreo into MuestreoCa
                            from mc in MuestreoCa.DefaultIfEmpty()

                            join a in bd.ProdAgenteCat on mc.IdAgen equals a.IdAgen into MuestreoAgenC
                            from ageC in MuestreoAgenC.DefaultIfEmpty()

                            join a in bd.ProdAgenteCat on mcam.IdAgenC equals a.IdAgen into MuestreoAgenSC
                            from ageCS in MuestreoAgenSC.DefaultIfEmpty()

                            join l in bd.CatLocalidades on mcam.CodLocalidad equals l.CodLocalidad into MuestreoLoc
                            from loc in MuestreoLoc.DefaultIfEmpty()

                            where man.Estatus != "L" && ageP.IdRegion == 2
                            select new ClassMuestreo
                            {
                                IdAnalisis_Residuo = man.Id,
                                IdMuestreo = an.Id,
                                Asesor = ageP.Abrev,
                                Cod_Prod = m.Cod_Prod,
                                Productor = prod.Nombre,
                                Cod_Campo = m.Cod_Campo,
                                Campo = mcam.Descripcion,
                                Sector = (short)ms.Sector,
                                Ha = mcam.Hectareas,
                                Compras_oportunidad = mcam.Compras_Oportunidad,
                                Fecha_solicitud = (DateTime)m.Fecha_solicitud,
                                Inicio_cosecha = (DateTime)an.Inicio_cosecha,
                                Ubicacion = loc.Descripcion,
                                Telefono = an.Telefono,
                                Liberacion = an.Liberacion,
                                Fecha_ejecucion = (DateTime)an.Fecha_ejecucion,
                                Analisis = man.Estatus,
                                Calidad_fruta = mc.Estatus,
                                IdAgenC = (short)mcam.IdAgenC,
                                AsesorC = ageC.Abrev,
                                AsesorCS = ageCS.Abrev,
                                Tarjeta = an.Tarjeta,
                                IdRegion = ageP.IdRegion,
                                Fecha_analisis = man.Fecha
                            }).Distinct();                 
                }
                else
                {
                    item = (from m in (from m in bd.ProdMuestreo
                                       group m by new
                                       {
                                           Cod_Empresa = m.Cod_Empresa,
                                           Cod_Prod = m.Cod_Prod,
                                           Cod_Campo = m.Cod_Campo,
                                           IdSector = m.IdSector
                                       } into x
                                       select new
                                       {
                                           Cod_Empresa = x.Key.Cod_Empresa,
                                           Cod_Prod = x.Key.Cod_Prod,
                                           Cod_Campo = x.Key.Cod_Campo,
                                           IdSector = x.Key.IdSector,
                                           Fecha_solicitud = x.Max(m => m.Fecha_solicitud)
                                       })

                            join an in bd.ProdMuestreo on new { m.Cod_Empresa, m.Cod_Prod, m.Cod_Campo, m.IdSector, m.Fecha_solicitud } equals new { an.Cod_Empresa, an.Cod_Prod, an.Cod_Campo, an.IdSector, an.Fecha_solicitud } into MuestreoR
                            from an in MuestreoR.DefaultIfEmpty()

                            join c in bd.ProdCamposCat on new { m.Cod_Empresa, m.Cod_Prod, m.Cod_Campo } equals new { c.Cod_Empresa, c.Cod_Prod, c.Cod_Campo } into MuestreoCam
                            from mcam in MuestreoCam.DefaultIfEmpty()

                            join p in bd.ProdProductoresCat on mcam.Cod_Prod equals p.Cod_Prod into MuestreoProd
                            from prod in MuestreoProd.DefaultIfEmpty()

                            join s in bd.ProdMuestreoSector on m.IdSector equals s.id into MuestreoSc
                            from ms in MuestreoSc.DefaultIfEmpty()

                            join r in bd.ProdAnalisis_Residuo on an.Id equals r.Id_Muestreo into MuestreoAn
                            from man in MuestreoAn.DefaultIfEmpty()

                            join a in bd.ProdAgenteCat on an.IdAgen equals a.IdAgen into MuestreoAgentes
                            from ageP in MuestreoAgentes.DefaultIfEmpty()

                            join cf in bd.ProdCalidadMuestreo on an.Id equals cf.Id_Muestreo into MuestreoCa
                            from mc in MuestreoCa.DefaultIfEmpty()

                            join a in bd.ProdAgenteCat on mc.IdAgen equals a.IdAgen into MuestreoAgenC
                            from ageC in MuestreoAgenC.DefaultIfEmpty()

                            join a in bd.ProdAgenteCat on mcam.IdAgenC equals a.IdAgen into MuestreoAgenSC
                            from ageCS in MuestreoAgenSC.DefaultIfEmpty()

                            join l in bd.CatLocalidades on mcam.CodLocalidad equals l.CodLocalidad into MuestreoLoc
                            from loc in MuestreoLoc.DefaultIfEmpty()

                            where  man.Estatus != "L" && (mcam.IdAgen == agenteSesion || mcam.IdAgenC == agenteSesion || mcam.IdAgenI == agenteSesion) //&& (cf.Estatus != "1" || cf.Estatus != "2") && m.Tarjeta != "S"                           

                            select new ClassMuestreo
                            {
                                IdAnalisis_Residuo = man.Id,
                                IdMuestreo = an.Id,
                                Asesor = ageP.Abrev,
                                Cod_Prod = m.Cod_Prod,
                                Productor = prod.Nombre,
                                Cod_Campo = m.Cod_Campo,
                                Campo = mcam.Descripcion,
                                Sector = (short)ms.Sector,
                                Ha = mcam.Hectareas,
                                Compras_oportunidad = mcam.Compras_Oportunidad,
                                Fecha_solicitud = (DateTime)m.Fecha_solicitud,
                                Inicio_cosecha = (DateTime)an.Inicio_cosecha,
                                Ubicacion = loc.Descripcion,
                                Telefono = an.Telefono,
                                Liberacion = an.Liberacion,
                                Fecha_ejecucion = (DateTime)an.Fecha_ejecucion,
                                Analisis = man.Estatus,
                                Calidad_fruta = mc.Estatus,
                                IdAgenC = (short)mcam.IdAgenC,
                                AsesorC = ageC.Abrev,
                                AsesorCS = ageCS.Abrev,
                                Tarjeta = an.Tarjeta,
                                IdRegion = ageP.IdRegion,
                                Fecha_analisis = man.Fecha
                            }).Distinct();
                }

                if (!String.IsNullOrEmpty(BuscarCod_Prod))
                {
                    item = item.Where(s => s.Cod_Prod.Contains(BuscarCod_Prod));
                }

                else if (agente != 0)
                {
                    item = item.Where(s => s.IdAgenC == agente);
                }

                else if (!String.IsNullOrEmpty(status))
                {
                    item = item.Where(s => s.Analisis == status);
                }

                //else if (agenteC != 0)
                //{
                //    item = item.Where(s => s.IdAgenC == agenteC);
                //}

                return View(item.ToList());
            }
            else
            {
                return RedirectToAction("Index", "Login");
            }
        }

        //public JsonResult Solicitudes_List()
        //{
        //    List<ClassMuestreo> resultados = bd.Database.SqlQuery<ClassMuestreo>("SELECT M.Id as IdMuestreo, A.Nombre as Ingeniero, M.Cod_Prod, P.Nombre as Productor, C.Descripcion as Campo, M.Fecha_solicitud, M.Fecha_ejecucion, isnull(L.Descripcion, '') as Ubicacion, M.Telefono, isnull(M.Liberacion, '') as Liberacion, " +
        //        "isnull((case when R.Estatus = 'R' then 'CON RESIDUOS' else (case when R.Estatus = 'P' then 'EN PROCESO' else (case when R.Estatus = 'F' then 'FUERA DE LIMITE' else (case when R.Estatus = 'L' then 'LIBERADO' end ) end) end) end),'') as Estatus, " +
        //        "isnull(M.Calidad_fruta, 'P') as Calidad_fruta, isnull(M.Tarjeta, 'P') as Tarjeta " +
        //        "FROM ProdMuestreo M " +
        //        "LEFT JOIN ProdAgenteCat A on M.IdAgen = A.IdAgen " +
        //        "LEFT JOIN ProdProductoresCat P on M.Cod_Prod = P.Cod_Prod " +
        //        "LEFT JOIN ProdCamposCat C on M.Cod_Prod = C.Cod_Prod and M.Cod_Campo = C.Cod_Campo " +
        //        "LEFT JOIN CatLocalidades L on C.CodLocalidad = L.CodLocalidad " +
        //        "LEFT JOIN ProdAnalisis_Residuo R ON M.Id = R.Id_Muestreo " +
        //        "order by M.Fecha_solicitud desc").ToList();
        //    return Json(resultados, JsonRequestBehavior.AllowGet);
        //}

        //Añadir fecha de ejecucion 
        public ActionResult EditSetSolicitud(int id)
        {
            if (Session["Nombre"] != null)
            {
                ViewData["Nombre"] = Session["Nombre"].ToString();
                var item = bd.ProdMuestreo.Where(x => x.Id == id).First();
                return View(item);
            }
            else
            {
                return RedirectToAction("Index", "Login");
            }
        }

        public JsonResult Sectores(int IdMuestreo = 0, int IdAnalisis = 0, int IdSector = 0, short Sector = 0, string Fecha_ejecucion = "")
        {
            List<ClassMuestreo> resultados = new List<ClassMuestreo>();
            try
            {
                var item = bd.ProdMuestreo.Where(x => x.Id == IdMuestreo).First();

                if (IdSector != 0)
                {
                    var model_sector = bd.ProdMuestreoSector.Find(IdSector);
                    if (model_sector != null)
                    {
                        var muestreo = bd.ProdMuestreo.FirstOrDefault(x => x.Id == IdMuestreo && x.IdSector == IdSector);
                        if (muestreo != null)
                        {
                            muestreo.IdSector = null;
                            bd.SaveChanges();
                        }
                        bd.ProdMuestreoSector.Remove(model_sector);
                        bd.SaveChanges();

                        var busca_sector = bd.ProdMuestreoSector.OrderByDescending(x => x.Cod_Prod == muestreo.Cod_Prod && x.Cod_Campo == muestreo.Cod_Campo).First();
                        if (busca_sector != null)
                        {
                            var prodMuestreoObj = bd.ProdMuestreo.SingleOrDefault(x => x.Id == IdMuestreo);
                            if (prodMuestreoObj != null)
                            {
                                prodMuestreoObj.IdSector = busca_sector.id;
                                bd.SaveChanges();
                            }
                        }
                    }
                }
                if (Sector != 0 && Fecha_ejecucion != "")
                {
                    item.Fecha_ejecucion = Convert.ToDateTime(Fecha_ejecucion);
                    item.IdAgenI = (short)Session["IdAgen"];

                    var valida_sector = bd.ProdMuestreoSector.FirstOrDefault(x => x.Cod_Prod == item.Cod_Prod && x.Cod_Campo == item.Cod_Campo && x.Sector == Sector);
                    if (valida_sector == null)
                    {
                        ProdMuestreoSector.Cod_Prod = item.Cod_Prod;
                        ProdMuestreoSector.Cod_Campo = item.Cod_Campo;
                        ProdMuestreoSector.Sector = Sector;
                        bd.ProdMuestreoSector.Add(ProdMuestreoSector);
                    }
                    bd.SaveChanges();

                    var IdMuestreoSector = ProdMuestreoSector.id;
                    var prodMuestreoObj = bd.ProdMuestreo.SingleOrDefault(x => x.Id == IdMuestreo);
                    if (prodMuestreoObj != null)
                    {
                        prodMuestreoObj.IdSector = IdMuestreoSector;
                        bd.SaveChanges();
                    }
                }
                if (Sector != 0)
                {
                    var valida_sector = bd.ProdMuestreoSector.FirstOrDefault(x => x.Cod_Prod == item.Cod_Prod && x.Cod_Campo == item.Cod_Campo && x.Sector == Sector);
                    if (valida_sector == null)
                    {
                        ProdMuestreoSector.Cod_Prod = item.Cod_Prod;
                        ProdMuestreoSector.Cod_Campo = item.Cod_Campo;
                        ProdMuestreoSector.Sector = Sector;
                        bd.ProdMuestreoSector.Add(ProdMuestreoSector);
                    }
                    bd.SaveChanges();

                    var IdMuestreoSector = ProdMuestreoSector.id;
                    var prodMuestreoObj = bd.ProdMuestreo.SingleOrDefault(x => x.Id == IdMuestreo);
                    if (prodMuestreoObj != null)
                    {
                        prodMuestreoObj.IdSector = IdMuestreoSector;
                        bd.SaveChanges();
                    }
                }
            }
            catch (Exception e)
            {
                e.ToString();
            }

            if (IdMuestreo != 0)
            {
                resultados = bd.Database.SqlQuery<ClassMuestreo>("Select S.id as IdSector, S.Sector, M.Fecha_ejecucion from ProdMuestreoSector S " +
                     "left join ProdMuestreo M on S.Cod_Prod=M.Cod_Prod and S.Cod_Campo=M.Cod_Campo where M.Id=" + IdMuestreo + "order by Sector").ToList();
            }
            else if (IdAnalisis != 0)
            {
                resultados = bd.Database.SqlQuery<ClassMuestreo>("Select S.id as IdSector, S.Sector, M.Fecha_ejecucion from ProdMuestreoSector S " +
                    "left join ProdAnalisis_Residuo A on S.Cod_Prod = A.Cod_Prod and S.Cod_Campo = A.Cod_Campo " +
                    "left join ProdMuestreo M on S.Cod_Prod = M.Cod_Prod and S.Cod_Campo = M.Cod_Campo where A.Id = " + IdAnalisis + " order by Sector").ToList();
            }
            return Json(resultados, JsonRequestBehavior.AllowGet);
        }

        //Liberar muestreo       
        [HttpPost]
        public ActionResult SetSolicitudL(ProdMuestreo model, string liberacion)
        {
            if (Session["Nombre"] != null)
            {
                ViewData["Nombre"] = Session["Nombre"].ToString();
            }
            else
            {
                return RedirectToAction("Index", "Login");
            }
            try
            {
                if (model.Id > 0 && liberacion != null)
                {
                    var item = bd.ProdMuestreo.Where(x => x.Id == model.Id).First();
                    item.Liberacion = liberacion;
                    item.IdAgen = (short)Session["IdAgen"];
                    bd.SaveChanges();

                    string productor, correo_p, correo_c, correo_i;

                    var campo = bd.ProdCamposCat.FirstOrDefault(m => m.Cod_Prod == item.Cod_Prod && m.Cod_Campo == item.Cod_Campo);
                    var email_p = bd.SIPGUsuarios.FirstOrDefault(m => m.IdAgen == campo.IdAgen);
                    correo_p = email_p.correo;
                    var email_c = bd.SIPGUsuarios.FirstOrDefault(m => m.IdAgen == campo.IdAgenC);
                    correo_c = email_c.correo;
                    var email_i = bd.SIPGUsuarios.FirstOrDefault(m => m.IdAgen == campo.IdAgenI);
                    correo_i = email_i.correo;
                    var prod = bd.ProdProductoresCat.FirstOrDefault(x => x.Cod_Prod == item.Cod_Prod);
                    productor = prod.Nombre;

                    MailMessage correo = new MailMessage();
                    correo.From = new MailAddress("indicadores.giddingsfruit@gmail.com", "Indicadores GiddingsFruit");
                    correo.To.Add(Session["Correo"].ToString());//correo_p
                    correo.CC.Add(correo_c);
                    correo.CC.Add(correo_i);
                    correo.Subject = "Muestreo Liberado: " + item.Cod_Prod;
                    correo.Body = "Liberado por: " + Session["Nombre"].ToString() + " <br/>";
                    correo.Body += " <br/>";
                    correo.Body += "Productor: " + item.Cod_Prod + " - " + productor + " <br/>";
                    correo.Body += " <br/>";
                    correo.Body += "Campo: " + item.Cod_Campo + " - " + campo.Descripcion + " <br/>";
                    correo.Body += " <br/>";
                    correo.Body += "Ubicacion: " + campo.Ubicacion + " <br/>";
                    correo.Body += " <br/>";
                    correo.Body += "Telefono: " + item.Telefono + "<br/>";
                    correo.Body += " <br/>";

                    correo.IsBodyHtml = true;
                    correo.Priority = MailPriority.Normal;

                    string sSmtpServer = "";
                    sSmtpServer = "smtp.gmail.com";

                    SmtpClient a = new SmtpClient();
                    a.Host = sSmtpServer;
                    a.Port = 587;//25
                    a.EnableSsl = true;
                    a.UseDefaultCredentials = false;// true;
                    a.Credentials = new System.Net.NetworkCredential("indicadores.giddingsfruit@gmail.com", "indicadores2019");
                    a.Send(correo);
                }
            }
            catch (Exception e) { e.ToString(); }
            return RedirectToAction("SetSolicitud", "Muestreo");

        }

        //Liberar tarjeta       
        [HttpPost]
        public ActionResult SetSolicitudLT(ProdMuestreo model, string tarjeta, string liberar_Tarjeta)
        {
            if (Session["Nombre"] != null)
            {
                ViewData["Nombre"] = Session["Nombre"].ToString();
            }
            else
            {
                return RedirectToAction("Index", "Login");
            }
            try
            {
                if (model.Id > 0 && tarjeta != null)
                {
                    var item = bd.ProdMuestreo.Where(x => x.Id == model.Id).First();
                    item.Tarjeta = tarjeta;
                    item.IdAgen_Tarjeta = (short)Session["IdAgen"];
                    item.Liberar_Tarjeta = liberar_Tarjeta;
                    bd.SaveChanges();

                    var muestreo_Calidad = bd.ProdCalidadMuestreo.Where(x => x.Id_Muestreo == model.Id).First();
                    var analisis_r = bd.ProdAnalisis_Residuo.Where(x => x.Id_Muestreo == model.Id).First();

                    string correo_p, correo_c, correo_i;

                    var campo = bd.ProdCamposCat.FirstOrDefault(m => m.Cod_Prod == item.Cod_Prod && m.Cod_Campo == item.Cod_Campo);
                    var email_p = bd.SIPGUsuarios.FirstOrDefault(m => m.IdAgen == campo.IdAgen);
                    correo_p = email_p.correo;
                    var email_c = bd.SIPGUsuarios.FirstOrDefault(m => m.IdAgen == campo.IdAgenC);
                    correo_c = email_c.correo;
                    var email_i = bd.SIPGUsuarios.FirstOrDefault(m => m.IdAgen == campo.IdAgenI);
                    correo_i = email_i.correo;
                    var prod = bd.ProdProductoresCat.FirstOrDefault(x => x.Cod_Prod == item.Cod_Prod);

                    MailMessage correo = new MailMessage();
                    correo.From = new MailAddress("indicadores.giddingsfruit@gmail.com", "Indicadores GiddingsFruit");
                    correo.To.Add(Session["Correo"].ToString());
                    correo.CC.Add(correo_p);
                    correo.CC.Add(correo_c);
                    correo.CC.Add(correo_i);
                    correo.Subject = "Entrega de tarjeta: " + item.Cod_Prod;
                    correo.Body = "Autorizado por: " + Session["Nombre"].ToString() + " <br/>";
                    correo.Body += " <br/>";
                    correo.Body += "Se ha autorizado la liberación para entrega de tarjeta del productor: " + item.Cod_Prod + " - " + prod.Nombre + " <br/>";
                    correo.Body += " <br/>";
                    correo.Body += "campo: " + item.Cod_Campo + " - " + campo.Descripcion + " ubicado en " + campo.Ubicacion + " <br/>";
                    correo.Body += " <br/>";
                    if (muestreo_Calidad.Estatus == "1")
                    {
                        correo.Body += "Evaluacion de calidad: APTA <br/>";
                    }
                    else if (muestreo_Calidad.Estatus == "2")
                    {
                        correo.Body += "Evaluacion de calidad: APTA CON CONDICIONES<br/>";
                    }
                    else if (muestreo_Calidad.Estatus == "3")
                    {
                        correo.Body += "Evaluacion de calidad: PENDIENTE <br/>";
                    }
                    correo.Body += " <br/>";
                    if (muestreo_Calidad != null)
                    {
                        correo.Body += "Incidencias encontradas: " + muestreo_Calidad.Incidencia + " <br/>";
                        correo.Body += " <br/>";
                        if (muestreo_Calidad.Propuesta != "")
                        {
                            correo.Body += "Propuesta: " + muestreo_Calidad.Propuesta + " <br/>";
                            correo.Body += " <br/>";
                        }
                    }

                    if (liberar_Tarjeta != "")
                    {
                        correo.Body += "Justificación: " + liberar_Tarjeta + "<br/>";
                        correo.Body += " <br/>";
                    }
                    if (analisis_r != null)
                    {
                        if (analisis_r.Estatus == "R")
                        {
                            correo.Body += "Resultado del analisis: CON RESIDUOS <br/>";
                        }
                        else if (analisis_r.Estatus == "P")
                        {
                            correo.Body += "Resultado del analisis: EN PROCESO <br/>";
                        }
                        else if (analisis_r.Estatus == "F")
                        {
                            correo.Body += "Resultado del analisis: FUERA DE LIMITE <br/>";
                        }
                        else if (analisis_r.Estatus == "L")
                        {
                            correo.Body += "Resultado del analisis: LIBERADO <br/>";
                        }
                        correo.Body += " <br/>";
                    }

                    correo.IsBodyHtml = true;
                    correo.Priority = MailPriority.Normal;

                    string sSmtpServer = "";
                    sSmtpServer = "smtp.gmail.com";

                    SmtpClient a = new SmtpClient();
                    a.Host = sSmtpServer;
                    a.Port = 587;//25
                    a.EnableSsl = true;
                    a.UseDefaultCredentials = false;// true;
                    a.Credentials = new System.Net.NetworkCredential("indicadores.giddingsfruit@gmail.com", "indicadores2019");
                    a.Send(correo);
                }
            }
            catch (Exception e) { e.ToString(); }
            return RedirectToAction("SetSolicitud", "Muestreo");

        }

        //Calidad fruta
        [HttpPost]
        public ActionResult SetSolicitudCF(ProdMuestreo model, string estatus, string incidencia = "", string propuesta = "")
        {
            if (Session["Nombre"] != null)
            {
                ViewData["Nombre"] = Session["Nombre"].ToString();
            }
            if (model.Id > 0)
            {
                ProdCalidadMuestreo ProdCalidadMuestreo = new ProdCalidadMuestreo();
                ProdCalidadMuestreo.Estatus = estatus;
                ProdCalidadMuestreo.Fecha = DateTime.Now;
                if (incidencia != "")
                {
                    ProdCalidadMuestreo.Incidencia = incidencia;
                    if (propuesta != "")
                    {
                        ProdCalidadMuestreo.Propuesta = propuesta;
                    }
                }
                ProdCalidadMuestreo.IdAgen = (short)Session["IdAgen"];
                ProdCalidadMuestreo.Id_Muestreo = model.Id;
                bd.ProdCalidadMuestreo.Add(ProdCalidadMuestreo);
                bd.SaveChanges();

                var item = bd.ProdMuestreo.Where(x => x.Id == model.Id).First();
                if (estatus != null)
                {
                    if (estatus == "1" || estatus == "2")
                    {
                        item.Tarjeta = "S";
                    }
                    if (estatus == "3")
                    {
                        item.Tarjeta = "N";
                    }
                    bd.SaveChanges();

                    var analisis = bd.ProdAnalisis_Residuo.FirstOrDefault(x => x.Id_Muestreo == model.Id);

                    try
                    {
                        string correo_p, correo_c, correo_i;

                        var campo = bd.ProdCamposCat.FirstOrDefault(m => m.Cod_Prod == item.Cod_Prod && m.Cod_Campo == item.Cod_Campo);
                        var email_p = bd.SIPGUsuarios.FirstOrDefault(m => m.IdAgen == campo.IdAgen);
                        correo_p = email_p.correo;
                        var email_c = bd.SIPGUsuarios.FirstOrDefault(m => m.IdAgen == campo.IdAgenC);
                        correo_c = email_c.correo;
                        var email_i = bd.SIPGUsuarios.FirstOrDefault(m => m.IdAgen == campo.IdAgenI);
                        correo_i = email_i.correo;
                        var prod = bd.ProdProductoresCat.FirstOrDefault(x => x.Cod_Prod == item.Cod_Prod);

                        MailMessage correo = new MailMessage();
                        correo.From = new MailAddress("indicadores.giddingsfruit@gmail.com", "Indicadores GiddingsFruit");
                        correo.To.Add(Session["Correo"].ToString());//correo_c
                        correo.CC.Add(correo_p);
                        correo.CC.Add(correo_i);
                        correo.Subject = "Calidad de fruta evaluada: " + item.Cod_Prod;
                        correo.Body += "Evaluado por: " + Session["Nombre"].ToString() + " <br/>";
                        correo.Body += " <br/>";
                        correo.Body += "Productor: " + item.Cod_Prod + " - " + prod.Nombre + " <br/>";
                        correo.Body += " <br/>";
                        correo.Body += "Campo: " + item.Cod_Campo + " - " + campo.Descripcion + " <br/>";
                        correo.Body += " <br/>";
                        if (estatus == "1")
                        {
                            correo.Body += "Estatus: APTA <br/>";
                            correo.Body += " <br/>";
                            if (analisis != null)
                            {
                                if (analisis.Estatus == "L")
                                {
                                    correo.Body += "Entregar tarjeta <br/>";
                                    correo.Body += " <br/>";
                                }
                            }
                        }
                        else if (estatus == "2")
                        {
                            correo.Body += "Estatus: APTA CON CONDICIONES<br/>";
                            correo.Body += " <br/>";
                            if (analisis.Estatus == "L")
                            {
                                correo.Body += "Entregar tarjeta <br/>";
                                correo.Body += " <br/>";
                            }
                        }
                        else if (estatus == "3")
                        {
                            correo.Body += "Estatus: PENDIENTE <br/>";
                            correo.Body += " <br/>";
                            correo.Body += "No entregar tarjeta <br/>";
                            correo.Body += " <br/>";
                        }
                        if (incidencia != "")
                        {
                            correo.Body += "Incidencia: " + incidencia + " <br/>";
                            correo.Body += " <br/>";
                            if (propuesta != "")
                            {
                                correo.Body += "Propuesta: " + propuesta + " <br/>";
                                correo.Body += " <br/>";
                            }
                        }
                        correo.IsBodyHtml = true;
                        correo.Priority = MailPriority.Normal;

                        string sSmtpServer = "";
                        sSmtpServer = "smtp.gmail.com";

                        SmtpClient a = new SmtpClient();
                        a.Host = sSmtpServer;
                        a.Port = 587;//25
                        a.EnableSsl = true;
                        a.UseDefaultCredentials = false;// true;
                        a.Credentials = new System.Net.NetworkCredential("indicadores.giddingsfruit@gmail.com", "indicadores2019");
                        a.Send(correo);
                    }
                    catch (Exception e) { e.ToString(); }
                }
            }
            return RedirectToAction("SetSolicitud", "Muestreo");
        }

        //Re-asignar a otro ing
        [HttpPost]
        public ActionResult SetReasignarIng(ProdMuestreo model, short idAgen)
        {
            if (Session["Nombre"] != null)
            {
                ViewData["Nombre"] = Session["Nombre"].ToString();
            }
            if (model.Id > 0)
            {
                string correo, productor, agente;

                var item = bd.ProdMuestreo.Where(x => x.Id == model.Id).First();
                var campo = bd.ProdCamposCat.FirstOrDefault(x => x.Cod_Prod == item.Cod_Prod && x.Cod_Campo == item.Cod_Campo);
                var user = bd.SIPGUsuarios.FirstOrDefault(x => x.IdAgen == idAgen);
                correo = user.correo;
                agente = user.Completo;

                var prod = bd.ProdProductoresCat.FirstOrDefault(x => x.Cod_Prod == item.Cod_Prod);
                productor = prod.Nombre;

                if (Session["Tipo"].ToString() == "P")
                {
                    campo.IdAgen = idAgen;
                }
                else if (Session["Tipo"].ToString() == "C")
                {
                    campo.IdAgenC = idAgen;
                }
                else if (Session["Tipo"].ToString() == "I")
                {
                    campo.IdAgenI = idAgen;
                }
                bd.SaveChanges();

                try
                {
                    MailMessage mailMessage = new MailMessage();
                    mailMessage.From = new MailAddress("indicadores.giddingsfruit@gmail.com", "Indicadores GiddingsFruit");
                    mailMessage.To.Add(Session["Correo"].ToString());
                    mailMessage.CC.Add(correo);
                    mailMessage.CC.Add("oscar.castillo@giddingsfruit.mx");
                    mailMessage.Subject = "Reasignación de código: " + item.Cod_Prod;
                    mailMessage.Body += "El productor: " + item.Cod_Prod + " - " + productor + " con campo: " + item.Cod_Campo + " - " + campo.Descripcion + " ha sido reasignado a " + agente + " por " + Session["Nombre"].ToString() + " <br/>";
                    //mailMessage.Body += " <br/>";
                    //mailMessage.Body += "Tiene solicitud de muestreo pendiente!<br/>";
                    //mailMessage.Body += " <br/>";
                    mailMessage.IsBodyHtml = true;
                    mailMessage.Priority = MailPriority.Normal;

                    string sSmtpServer = "";
                    sSmtpServer = "smtp.gmail.com";

                    SmtpClient a = new SmtpClient();
                    a.Host = sSmtpServer;
                    a.Port = 587;//25
                    a.EnableSsl = true;
                    a.UseDefaultCredentials = false;// true;
                    a.Credentials = new System.Net.NetworkCredential("indicadores.giddingsfruit@gmail.com", "indicadores2019");
                    a.Send(mailMessage);
                }
                catch (Exception e)
                {
                    e.ToString();
                }
            }
            return RedirectToAction("SetSolicitud", "Muestreo");
        }

        public ActionResult Nuevo_Analisis_Residuo(ProdAnalisis_Residuo model)
        {
            if (Session["Nombre"] != null)
            {
                ViewData["Nombre"] = Session["Nombre"].ToString();
                List<ClassMuestreo> lst_zonas = (from d in bd.ProdZonasRastreoCat
                                                 select new ClassMuestreo
                                                 {
                                                     CodZona = d.Codigo,
                                                     DescZona = d.DescZona
                                                 }).Where(x => x.CodZona != "LS").OrderBy(x => x.DescZona).ToList();

                List<SelectListItem> zonas = lst_zonas.ConvertAll(d =>
                {
                    return new SelectListItem()
                    {
                        Text = d.DescZona.ToString(),
                        Value = d.CodZona.ToString(),
                        Selected = false
                    };
                });
                ViewBag.zonas = zonas;
            }
            else
            {
                return RedirectToAction("Index", "Login");
            }
            return View();
        }

        [HttpPost]
        public ActionResult Nuevo_Analisis_Residuo(ProdAnalisis_Residuo model, short sector = 0, int LiberacionUSA = 0, int LiberacionEU = 0)
        {
            try
            {
                if (Session["Nombre"] != null)
                {
                    ViewData["Nombre"] = Session["Nombre"].ToString();
                    List<ClassMuestreo> lst_zonas = (from d in bd.ProdZonasRastreoCat
                                                     select new ClassMuestreo
                                                     {
                                                         CodZona = d.Codigo,
                                                         DescZona = d.DescZona
                                                     }).Where(x => x.CodZona != "LS").OrderBy(x => x.DescZona).ToList();

                    List<SelectListItem> zonas = lst_zonas.ConvertAll(d =>
                    {
                        return new SelectListItem()
                        {
                            Text = d.DescZona.ToString(),
                            Value = d.CodZona.ToString(),
                            Selected = false
                        };
                    });
                    ViewBag.zonas = zonas;

                    int idSector = 0;

                    if (sector == 0)
                    {
                        sector = 1;
                    }

                    var model_sector = bd.ProdMuestreoSector.FirstOrDefault(m => m.Cod_Prod == model.Cod_Prod && m.Cod_Campo == model.Cod_Campo && m.Sector == sector);
                    if (model_sector == null)
                    {
                        ProdMuestreoSector.Cod_Prod = model.Cod_Prod;
                        ProdMuestreoSector.Cod_Campo = model.Cod_Campo;
                        ProdMuestreoSector.Sector = sector;
                        bd.ProdMuestreoSector.Add(ProdMuestreoSector);
                        bd.SaveChanges();
                        idSector = ProdMuestreoSector.id;
                    }

                    var modeloExistente = bd.ProdAnalisis_Residuo.FirstOrDefault(m => m.Cod_Prod == model.Cod_Prod && m.Cod_Campo == model.Cod_Campo && m.Estatus == model.Estatus);

                    if (modeloExistente == null)
                    {
                        DateTime fechaLiberacionUSA = DateTime.Now, fechaLiberacionEU = DateTime.Now;
                        ProdAnalisis_Residuo.Cod_Empresa = 2;
                        ProdAnalisis_Residuo.Cod_Prod = model.Cod_Prod;
                        ProdAnalisis_Residuo.Cod_Campo = model.Cod_Campo;
                        ProdAnalisis_Residuo.CodZona = model.CodZona;
                        ProdAnalisis_Residuo.Fecha_envio = model.Fecha_envio;
                        ProdAnalisis_Residuo.Fecha_entrega = model.Fecha_entrega;
                        ProdAnalisis_Residuo.Estatus = model.Estatus;
                        ProdAnalisis_Residuo.Num_analisis = model.Num_analisis;
                        ProdAnalisis_Residuo.Laboratorio = model.Laboratorio;
                        ProdAnalisis_Residuo.Comentarios = model.Comentarios;
                        ProdAnalisis_Residuo.IdAgen = (short)Session["IdAgen"];
                        ProdAnalisis_Residuo.Fecha = DateTime.Now;
                        ProdAnalisis_Residuo.Folio = model.Folio;
                        if (idSector == 0)
                        {
                            ProdAnalisis_Residuo.IdSector = model_sector.id;
                        }
                        else
                        {
                            ProdAnalisis_Residuo.IdSector = idSector;
                        }

                        if (model.Estatus == "F")
                        {
                            fechaLiberacionUSA = Convert.ToDateTime(model.Fecha_envio).AddDays(LiberacionUSA);
                            fechaLiberacionEU = Convert.ToDateTime(model.Fecha_envio).AddDays(LiberacionEU);

                            ProdAnalisis_Residuo.LiberacionUSA = fechaLiberacionUSA;
                            ProdAnalisis_Residuo.LiberacionEU = fechaLiberacionEU;
                        }

                        if (model.Estatus == "L" && model.Comentarios != null)
                        {
                            TempData["sms"] = "Si el Status es LIBERADO no escriba comentarios porfavor!";
                            ViewBag.error = TempData["sms"].ToString();
                        }
                        else
                        {
                            bd.ProdAnalisis_Residuo.Add(ProdAnalisis_Residuo);
                            bd.SaveChanges();

                            var IdProdAnalisis = ProdAnalisis_Residuo.Id;

                            var prodSectoresObj = bd.ProdMuestreoSector.SingleOrDefault(x => x.Cod_Prod == model.Cod_Prod && x.Cod_Campo == model.Cod_Campo && x.Sector == sector);
                            if (prodSectoresObj != null)
                            {
                                prodSectoresObj.Cod_Prod = model.Cod_Prod;
                                prodSectoresObj.Cod_Campo = model.Cod_Campo;
                                prodSectoresObj.Sector = sector;
                                bd.SaveChanges();

                                var prodAnalisisObj = bd.ProdAnalisis_Residuo.SingleOrDefault(x => x.Id == IdProdAnalisis);
                                if (prodAnalisisObj != null)
                                {
                                    prodAnalisisObj.IdSector = prodSectoresObj.id;
                                    bd.SaveChanges();
                                }
                            }

                            TempData["sms"] = "Datos guardados con éxito";
                            ViewBag.sms = TempData["sms"].ToString();
                        }
                    }
                    else
                    {
                        TempData["sms"] = "Informacion duplicada, verifique porfavor";
                        ViewBag.error = TempData["sms"].ToString();
                    }
                }
                else
                {
                    return RedirectToAction("Index", "Login");
                }

            }
            catch (Exception e)
            {
                TempData["error"] = e.ToString();
                ViewBag.error = TempData["error"].ToString();
            }
            return View();
        }

        public ActionResult Analisis_Residuos(ProdAnalisis_Residuo model)
        {
            if (Session["Nombre"] != null)
            {
                ViewData["Nombre"] = Session["Nombre"].ToString();
                List<ClassMuestreo> lst_zonas = (from d in bd.ProdZonasRastreoCat
                                                 select new ClassMuestreo
                                                 {
                                                     CodZona = d.Codigo,
                                                     DescZona = d.DescZona
                                                 }).OrderBy(x => x.DescZona).ToList();

                List<SelectListItem> zonas = lst_zonas.ConvertAll(d =>
                {
                    return new SelectListItem()
                    {
                        Text = d.DescZona.ToString(),
                        Value = d.CodZona.ToString(),
                        Selected = false
                    };
                });
                ViewBag.zonas = zonas;
            }
            else
            {
                return RedirectToAction("Index", "Login");
            }
            return View();
        }
        public JsonResult Analisis_ResiduosData(string Cod_Prod = "", short Cod_Campo = 0, short Sector = 0, string CodZona = "", string Fecha_envio = "", string Fecha_entrega = "", string Estatus = "", int Num_analisis = 0, string Laboratorio = "", string Comentarios = "", int LiberacionUSA = 0, int LiberacionEU = 0, string Folio="")
        {
            try
            {
                if (Session["Nombre"] != null)
                {
                    if (Estatus != "")
                    {
                        var muestreo = bd.ProdMuestreo.FirstOrDefault(m => m.Cod_Prod == Cod_Prod && m.Cod_Campo == Cod_Campo);

                        int idsector = 0;
                        var modeloExistente = bd.ProdAnalisis_Residuo.FirstOrDefault(m => m.Cod_Prod == Cod_Prod && m.Cod_Campo == Cod_Campo && m.Estatus == Estatus && m.Num_analisis== Num_analisis);
                        
                        if (modeloExistente == null)
                        {
                            var sectores = bd.ProdMuestreoSector.FirstOrDefault(m => m.Cod_Prod == Cod_Prod && m.Cod_Campo == Cod_Campo);
                            idsector = sectores.id;

                            DateTime fechaLiberacionUSA = DateTime.Now, fechaLiberacionEU = DateTime.Now;
                            ProdAnalisis_Residuo.Cod_Empresa = 2;
                            ProdAnalisis_Residuo.Cod_Prod = Cod_Prod;
                            ProdAnalisis_Residuo.Cod_Campo = Cod_Campo;
                            ProdAnalisis_Residuo.CodZona = CodZona;
                            ProdAnalisis_Residuo.Fecha_envio = Convert.ToDateTime(Fecha_envio); ;
                            ProdAnalisis_Residuo.Fecha_entrega = Convert.ToDateTime(Fecha_entrega); ;
                            ProdAnalisis_Residuo.Estatus = Estatus;
                            ProdAnalisis_Residuo.Num_analisis = Num_analisis;
                            ProdAnalisis_Residuo.Laboratorio = Laboratorio;
                            ProdAnalisis_Residuo.Comentarios = Comentarios;
                            ProdAnalisis_Residuo.IdAgen = (short)Session["IdAgen"];
                            ProdAnalisis_Residuo.Fecha = DateTime.Now;
                            if (muestreo != null)
                            {
                                ProdAnalisis_Residuo.Id_Muestreo = muestreo.Id;
                            }
                            ProdAnalisis_Residuo.IdSector = idsector;
                            ProdAnalisis_Residuo.Folio = Folio;

                            if (Estatus == "F")
                            {
                                fechaLiberacionUSA = Convert.ToDateTime(Fecha_envio).AddDays(LiberacionUSA);
                                fechaLiberacionEU = Convert.ToDateTime(Fecha_envio).AddDays(LiberacionEU);

                                ProdAnalisis_Residuo.LiberacionUSA = fechaLiberacionUSA;
                                ProdAnalisis_Residuo.LiberacionEU = fechaLiberacionEU;
                            }

                            bd.ProdAnalisis_Residuo.Add(ProdAnalisis_Residuo);
                            bd.SaveChanges();

                            //var IdProdMuestreo = ProdAnalisis_Residuo.Id_Muestreo;
                            //var IdProdAnalisis = ProdAnalisis_Residuo.Id;

                            ////var prodMuestreoObj = bd.ProdMuestreo.SingleOrDefault(x => x.Id == IdProdMuestreo);
                            ////if (prodMuestreoObj != null)
                            ////{
                            ////    prodMuestreoObj.Id_Analisis_Residuo = IdProdAnalisis;
                            ////    bd.SaveChanges();
                            ////}

                            //var prodSectoresObj = bd.ProdMuestreoSector.SingleOrDefault(x => x.Cod_Prod == Cod_Prod && x.Cod_Campo == Cod_Campo && x.Sector== Sector);
                            //if (prodSectoresObj != null)
                            //{
                            //    prodSectoresObj.Cod_Prod = Cod_Prod;
                            //    prodSectoresObj.Cod_Campo = Cod_Campo;
                            //    prodSectoresObj.Sector = Sector;
                            //    bd.SaveChanges();

                            //    var prodAnalisisObj = bd.ProdAnalisis_Residuo.SingleOrDefault(x => x.Id == IdProdAnalisis);
                            //    if (prodAnalisisObj != null)
                            //    {
                            //        prodAnalisisObj.IdSector = prodSectoresObj.id;
                            //        bd.SaveChanges();
                            //    }
                            //}

                            TempData["sms"] = "Datos guardados con éxito";
                            ViewBag.sms = TempData["sms"].ToString();

                            if (CodZona == "LS")
                            {
                                try
                                {
                                    string estatus = "", correo_p, correo_c, correo_i;

                                    var campo = bd.ProdCamposCat.FirstOrDefault(m => m.Cod_Prod == Cod_Prod && m.Cod_Campo == Cod_Campo);
                                    var email_p = bd.SIPGUsuarios.FirstOrDefault(m => m.IdAgen == campo.IdAgen);
                                    correo_p = email_p.correo;
                                    var email_c = bd.SIPGUsuarios.FirstOrDefault(m => m.IdAgen == campo.IdAgenC);
                                    correo_c = email_c.correo;
                                    var email_i = bd.SIPGUsuarios.FirstOrDefault(m => m.IdAgen == campo.IdAgenI);
                                    correo_i = email_i.correo;
                                    var prod = bd.ProdProductoresCat.FirstOrDefault(x => x.Cod_Prod == Cod_Prod);

                                    if (Estatus == "R")
                                    {
                                        estatus = "Estatus: CON RESIDUOS";
                                    }
                                    else if (Estatus == "P")
                                    {
                                        estatus = "Estatus: EN PROCESO";
                                    }
                                    else if (Estatus == "F")
                                    {
                                        estatus = "Estatus: FUERA DEL LIMITE";
                                    }
                                    else if (Estatus == "L")
                                    {
                                        estatus = "Estatus: LIBERADO";
                                    }

                                    MailMessage correo = new MailMessage();
                                    correo.From = new MailAddress("indicadores.giddingsfruit@gmail.com", "Indicadores GiddingsFruit");
                                    correo.To.Add(Session["Correo"].ToString());//"marholy.martinez@giddingsfruit.mx"
                                    correo.CC.Add(correo_p);
                                    correo.CC.Add(correo_c);
                                    if (correo_i != "jesus.palafox@giddingsfruit.mx")
                                    {
                                        correo.CC.Add(correo_i);
                                    }
                                    correo.Subject = "Analisis de residuos: " + campo.Cod_Prod + " - " + estatus;
                                    correo.Body += "Productor: " + campo.Cod_Prod + " - " + prod.Nombre + " <br/>";
                                    correo.Body += " <br/>";
                                    correo.Body += "Campo: " + campo.Cod_Campo + " - " + campo.Descripcion + " <br/>";
                                    correo.Body += " <br/>";
                                    correo.Body += "Fecha de envio: " + Fecha_envio + "<br/>";
                                    correo.Body += " <br/>";
                                    correo.Body += "Fecha de entrega: " + Fecha_entrega + "<br/>";
                                    correo.Body += " <br/>";
                                    correo.Body += estatus + "<br/>";
                                    correo.Body += " <br/>";
                                    if (muestreo.Tarjeta != null)
                                    {
                                        if (Estatus == "L" && muestreo.Tarjeta == "S")
                                        {
                                            correo.Body += "Entregar tarjeta <br/>";
                                            correo.Body += " <br/>";
                                        }
                                    }
                                    if (Estatus == "F")
                                    {
                                        if (LiberacionUSA != 0)
                                        {
                                            correo.Body += "Fecha de liberacion para USA: " + fechaLiberacionEU.ToString("yyyy-MM-dd") + "<br/>";
                                            correo.Body += " <br/>";
                                        }
                                        if (LiberacionEU != 0)
                                        {
                                            correo.Body += "Fecha de liberacion para EUROPA: " + fechaLiberacionUSA.ToString("yyyy-MM-dd") + "<br/>";
                                            correo.Body += " <br/>";
                                        }
                                    }
                                    correo.Body += "Numero de analisis: " + Num_analisis + "<br/>";
                                    correo.Body += " <br/>";
                                    correo.Body += "Laboratorio: " + Laboratorio + "<br/>";
                                    correo.Body += " <br/>";
                                    if (Comentarios != null)
                                    {
                                        correo.Body += "Comentarios: " + Comentarios + "<br/>";
                                        correo.Body += " <br/>";
                                    }
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
                                catch (Exception e)
                                {
                                    e.ToString();
                                }
                            }
                        }
                        else
                        {
                            TempData["sms"] = "Informacion duplicada, verifique porfavor";
                            ViewBag.error = TempData["sms"].ToString();
                        }
                    }
                }
            }
            catch (Exception e)
            {
                e.ToString();
            }

            ClassMuestreo resultados = bd.Database.SqlQuery<ClassMuestreo>("select S.Cod_Prod, P.Nombre as Productor, S.Cod_Campo, isnull(STUFF((SELECT ',' + cast(Sector as varchar) as SectorList FROM ProdMuestreoSector " +
                "where cod_prod='" + Cod_Prod + "' and Cod_Campo=" + Cod_Campo + " FOR XML PATH(''), TYPE).value('.', 'NVARCHAR(MAX)'), 1, 1, ''),0) as SectorList, " +
                "T.Descripcion as Tipo, Pr.Descripcion as Producto, A.Folio, A.CodZona, isnull(A.Num_analisis, 0) as Num_analisis " +
                "from ProdMuestreoSector S " +
                "Left Join ProdAnalisis_Residuo A on S.id = A.IdSector " +
                "Left Join ProdCamposCat C on S.Cod_Prod = C.Cod_Prod AND S.Cod_Campo = C.Cod_Campo " +
                "Left Join ProdProductoresCat P on S.Cod_Prod = P.Cod_Prod " +
                "Left Join CatTiposProd T on C.Tipo = T.Tipo " +
                "Left Join CatProductos Pr on C.Tipo = Pr.Tipo and C.Producto = Pr.Producto " +
                "where S.Cod_Prod = '" + Cod_Prod + "' and S.Cod_Campo = " + Cod_Campo + "").First();
            return Json(resultados, JsonRequestBehavior.AllowGet);
        }
        public ActionResult Resultados_Analisis(string status = "", string c_zona = "", string cod_prod = "")
        {
            if (Session["Nombre"] != null)
            {
                ViewData["Nombre"] = Session["Nombre"].ToString();

                List<SelectListItem> lst_Status = new List<SelectListItem>();
                lst_Status.Add(new SelectListItem() { Text = "--Resultado del analisis--", Value = null });
                lst_Status.Add(new SelectListItem() { Text = "CON RESIDUOS", Value = "R" });
                lst_Status.Add(new SelectListItem() { Text = "EN PROCESO", Value = "P" });
                lst_Status.Add(new SelectListItem() { Text = "FUERA DE LIMITE", Value = "F" });
                lst_Status.Add(new SelectListItem() { Text = "LIBERADO", Value = "L" });

                ViewBag.List_Status = lst_Status;

                List<ProdZonasRastreoCat> lst_Zona = new List<ProdZonasRastreoCat>();
                lst_Zona = bd.ProdZonasRastreoCat.OrderBy(x => x.DescZona).ToList();
                ViewBag.List_Zona = lst_Zona;

                if (status == "--Resultado del analisis--")
                {
                    status = "";
                }

                short agenteSesion = (short)Session["IdAgen"];
                IQueryable<ClassMuestreo> item = null;
                if (Session["IdAgen"].ToString() == "1" || Session["IdAgen"].ToString() == "205" || Session["IdAgen"].ToString() == "153" || Session["IdAgen"].ToString() == "281" || Session["IdAgen"].ToString() == "167" || Session["IdAgen"].ToString() == "182")
                {
                    item = from r in bd.ProdAnalisis_Residuo
                           join c in bd.ProdCamposCat on new { r.Cod_Empresa, r.Cod_Prod, r.Cod_Campo } equals new { c.Cod_Empresa, c.Cod_Prod, c.Cod_Campo } into MuestreoCam
                           from mcam in MuestreoCam.DefaultIfEmpty()
                           join s in bd.ProdMuestreoSector on r.IdSector equals s.id into MuestreoSc
                           from ms in MuestreoSc.DefaultIfEmpty()
                           join p in bd.ProdProductoresCat on r.Cod_Prod equals p.Cod_Prod into MuestreoProd
                           from prod in MuestreoProd.DefaultIfEmpty()
                           join l in bd.ProdZonasRastreoCat on mcam.IdZona equals l.IdZona into Zonas
                           from z in Zonas.DefaultIfEmpty()
                           join t in bd.CatTiposProd on mcam.Tipo equals t.Tipo into Tipos
                           from tp in Tipos.DefaultIfEmpty()
                           join pr in bd.CatProductos on new { mcam.Tipo, mcam.Producto } equals new { pr.Tipo, pr.Producto } into Productos
                           from prd in Productos.DefaultIfEmpty()
                           orderby r.Fecha
                           select new ClassMuestreo
                           {
                               Cod_Prod = r.Cod_Prod,
                               Cod_Campo = r.Cod_Campo,
                               Sector = (short)ms.Sector,
                               Productor = prod.Nombre,
                               Tipo = tp.Descripcion,
                               Producto = prd.Descripcion,
                               CodZona = r.CodZona,
                               DescZona = z.DescZona,
                               Fecha_envio = (DateTime)r.Fecha_envio,
                               Fecha_entrega = (DateTime)r.Fecha_entrega,
                               Analisis = r.Estatus,
                               Num_analisis = r.Num_analisis,
                               Laboratorio = r.Laboratorio,
                               LiberacionUSA = (DateTime)r.LiberacionUSA,
                               LiberacionEU = (DateTime)r.LiberacionEU,
                               Comentarios = r.Comentarios,
                               IdAgen = mcam.IdAgen,
                               IdAgenC = (short?)(int)mcam.IdAgenC,
                               IdAgenI = mcam.IdAgenI
                           };
                }
                else
                {
                    item = from r in bd.ProdAnalisis_Residuo
                           join c in bd.ProdCamposCat on new { r.Cod_Empresa, r.Cod_Prod, r.Cod_Campo } equals new { c.Cod_Empresa, c.Cod_Prod, c.Cod_Campo } into MuestreoCam
                           from mcam in MuestreoCam.DefaultIfEmpty()
                           join s in bd.ProdMuestreoSector on r.IdSector equals s.id into MuestreoSc
                           from ms in MuestreoSc.DefaultIfEmpty()
                           join p in bd.ProdProductoresCat on r.Cod_Prod equals p.Cod_Prod into MuestreoProd
                           from prod in MuestreoProd.DefaultIfEmpty()
                           join l in bd.ProdZonasRastreoCat on mcam.IdZona equals l.IdZona into Zonas
                           from z in Zonas.DefaultIfEmpty()
                           join t in bd.CatTiposProd on mcam.Tipo equals t.Tipo into Tipos
                           from tp in Tipos.DefaultIfEmpty()
                           join pr in bd.CatProductos on new { mcam.Tipo, mcam.Producto } equals new { pr.Tipo, pr.Producto } into Productos
                           from prd in Productos.DefaultIfEmpty()
                           where mcam.IdAgen == agenteSesion || mcam.IdAgen == agenteSesion || mcam.IdAgenI == agenteSesion
                           orderby r.Fecha
                           select new ClassMuestreo
                           {
                               Cod_Prod = r.Cod_Prod,
                               Cod_Campo = r.Cod_Campo,
                               Sector = (short)ms.Sector,
                               Productor = prod.Nombre,
                               Tipo = tp.Descripcion,
                               Producto = prd.Descripcion,
                               CodZona = z.Codigo,
                               DescZona = z.DescZona,
                               Fecha_envio = (DateTime)r.Fecha_envio,
                               Fecha_entrega = (DateTime)r.Fecha_entrega,
                               Analisis = r.Estatus,
                               Num_analisis = r.Num_analisis,
                               Laboratorio = r.Laboratorio,
                               LiberacionUSA = (DateTime)r.LiberacionUSA,
                               LiberacionEU = (DateTime)r.LiberacionEU,
                               Comentarios = r.Comentarios
                           };
                }

                //if (Session["IdAgen"].ToString() != "205" || Session["IdAgen"].ToString() != "153" || Session["IdAgen"].ToString() != "281" || Session["IdAgen"].ToString() != "167" || Session["IdAgen"].ToString() != "182")
                //{
                //    item = item.Where(r => r.IdAgen==agenteSesion || r.IdAgenC == agenteSesion || r.IdAgenI == agenteSesion);
                //}

                if (!String.IsNullOrEmpty(status))
                {
                    item = item.Where(r => r.Analisis.Contains(status));
                }

                else if (!String.IsNullOrEmpty(c_zona))
                {
                    item = item.Where(s => s.CodZona.Contains(c_zona));
                }

                else if (!String.IsNullOrEmpty(cod_prod))
                {
                    item = item.Where(s => s.Cod_Prod.Contains(cod_prod));
                }
                return View(item.ToList());
            }
            else
            {
                return RedirectToAction("Index", "Login");
            }
            //return View();
        }
        public JsonResult Resultados_Analisis_List(string estatus = "", string c_zona = "", string cod_prod = "")
        {
            List<ClassMuestreo> resultados = null;
            if ((short)Session["IdAgen"] == 205 || Session["IdAgen"].ToString() == "153" || Session["IdAgen"].ToString() == "281")
            {
                if (estatus != "")
                {
                    resultados = bd.Database.SqlQuery<ClassMuestreo>("select R.Cod_Prod, R.Cod_Campo, ISNULL(S.Sector,0) as Sector, P.Nombre as Productor, T.Descripcion as Tipo, V.Descripcion as Producto, isnull(L.DescZona,'') as Zona, R.Fecha_envio, R.Fecha_entrega, " +
                    "R.Estatus, R.Num_analisis, R.Laboratorio, isnull(R.LiberacionUSA,''), R.LiberacionEU, UPPER(isnull(R.Comentarios, '')) as Comentarios " +
                    "FROM ProdAnalisis_Residuo R " +
                    "LEFT JOIN ProdMuestreoSector S ON R.IdSector = S.Id " +
                    "Left Join ProdProductoresCat P on R.Cod_Prod = P.Cod_Prod " +
                    "Left Join ProdCamposCat C on R.Cod_Prod = C.Cod_Prod and R.Cod_Campo = C.Cod_Campo " +
                    "Left Join CatTiposProd T on C.Tipo = T.Tipo " +
                    "Left Join CatProductos V on C.Tipo = V.Tipo AND C.Producto = V.Producto " +
                    "Left Join ProdZonasRastreoCat L on R.CodZona = L.Codigo " +
                    "where R.Cod_Prod is not null  and R.Estatus='" + estatus + "' " +
                    "order by R.Fecha, R.Cod_Prod, R.Cod_Campo, S.Sector, R.Num_analisis desc").ToList();
                }

                else if (c_zona != "")
                {
                    resultados = bd.Database.SqlQuery<ClassMuestreo>("select R.Cod_Prod, R.Cod_Campo, ISNULL(S.Sector,0) as Sector, P.Nombre as Productor, T.Descripcion as Tipo, V.Descripcion as Producto, isnull(L.DescZona,'') as Zona, R.Fecha_envio, R.Fecha_entrega, " +
                     "R.Estatus, R.Num_analisis, R.Laboratorio, R.LiberacionUSA, R.LiberacionEU, UPPER(isnull(R.Comentarios, '')) as Comentarios " +
                     "FROM ProdAnalisis_Residuo R " +
                     "LEFT JOIN ProdMuestreoSector S ON R.IdSector = S.Id " +
                     "Left Join ProdProductoresCat P on R.Cod_Prod = P.Cod_Prod " +
                     "Left Join ProdCamposCat C on R.Cod_Prod = C.Cod_Prod and R.Cod_Campo = C.Cod_Campo " +
                     "Left Join CatTiposProd T on C.Tipo = T.Tipo " +
                     "Left Join CatProductos V on C.Tipo = V.Tipo AND C.Producto = V.Producto " +
                     "Left Join ProdZonasRastreoCat L on R.CodZona = L.Codigo " +
                     "where R.Cod_Prod is not null  and R.CodZona='" + c_zona + "' " +
                     "order by R.Fecha, R.Cod_Prod, R.Cod_Campo, S.Sector, R.Num_analisis desc").ToList();
                }

                else if (cod_prod != "")
                {
                    resultados = bd.Database.SqlQuery<ClassMuestreo>("select R.Cod_Prod, R.Cod_Campo, ISNULL(S.Sector,0) as Sector, P.Nombre as Productor, T.Descripcion as Tipo, V.Descripcion as Producto, isnull(L.DescZona,'') as Zona, R.Fecha_envio, R.Fecha_entrega, " +
                        "R.Estatus, R.Num_analisis, R.Laboratorio, R.LiberacionUSA, R.LiberacionEU, UPPER(isnull(R.Comentarios, '')) as Comentarios " +
                        "FROM ProdAnalisis_Residuo R " +
                        "LEFT JOIN ProdMuestreoSector S ON R.IdSector = S.Id " +
                        "Left Join ProdProductoresCat P on R.Cod_Prod = P.Cod_Prod " +
                        "Left Join ProdCamposCat C on R.Cod_Prod = C.Cod_Prod and R.Cod_Campo = C.Cod_Campo " +
                        "Left Join CatTiposProd T on C.Tipo = T.Tipo " +
                        "Left Join CatProductos V on C.Tipo = V.Tipo AND C.Producto = V.Producto " +
                        "Left Join ProdZonasRastreoCat L on R.CodZona = L.Codigo " +
                        "where R.Cod_Prod is not null and R.Cod_Prod='" + cod_prod + "' " +
                        "order by R.Fecha, R.Cod_Prod, R.Cod_Campo, S.Sector, R.Num_analisis desc").ToList();
                }

                else
                {
                    resultados = bd.Database.SqlQuery<ClassMuestreo>("select R.Cod_Prod, R.Cod_Campo, ISNULL(S.Sector,0) as Sector, P.Nombre as Productor, T.Descripcion as Tipo, V.Descripcion as Producto, isnull(L.DescZona,'') as Zona, R.Fecha_envio, R.Fecha_entrega, " +
                        "R.Estatus, R.Num_analisis, R.Laboratorio, R.LiberacionUSA, R.LiberacionEU, UPPER(isnull(R.Comentarios, '')) as Comentarios " +
                        "FROM ProdAnalisis_Residuo R " +
                        "LEFT JOIN ProdMuestreoSector S ON R.IdSector = S.Id " +
                        "Left Join ProdProductoresCat P on R.Cod_Prod = P.Cod_Prod " +
                        "Left Join ProdCamposCat C on R.Cod_Prod = C.Cod_Prod and R.Cod_Campo = C.Cod_Campo " +
                        "Left Join CatTiposProd T on C.Tipo = T.Tipo " +
                        "Left Join CatProductos V on C.Tipo = V.Tipo AND C.Producto = V.Producto " +
                        "Left Join ProdZonasRastreoCat L on R.CodZona = L.Codigo " +
                        "where R.Cod_Prod is not null " +
                        "order by R.Fecha, R.Cod_Prod, R.Cod_Campo, S.Sector, R.Num_analisis desc").ToList();
                }
            }
            else if ((short)Session["IdAgen"] == 1)
            {
                if (estatus != "")
                {
                    resultados = bd.Database.SqlQuery<ClassMuestreo>("select R.Cod_Prod, R.Cod_Campo, ISNULL(S.Sector,0) as Sector, P.Nombre as Productor, T.Descripcion as Tipo, V.Descripcion as Producto, isnull(L.DescZona,'') as Zona, R.Fecha_envio, R.Fecha_entrega, " +
                      "R.Estatus, R.Num_analisis, R.Laboratorio, isnull(convert(varchar, R.LiberacionUSA, 103), '') as LiberacionUSA, isnull(convert(varchar, R.LiberacionEU, 103), '') as LiberacionEU, UPPER(isnull(R.Comentarios, '')) as Comentarios " +
                      "FROM ProdAnalisis_Residuo R " +
                      "LEFT JOIN ProdMuestreoSector S ON R.IdSector = S.Id " +
                      "Left Join ProdProductoresCat P on R.Cod_Prod = P.Cod_Prod " +
                      "Left Join ProdCamposCat C on R.Cod_Prod = C.Cod_Prod and R.Cod_Campo = C.Cod_Campo " +
                      "Left Join CatTiposProd T on C.Tipo = T.Tipo " +
                      "Left Join CatProductos V on C.Tipo = V.Tipo AND C.Producto = V.Producto " +
                      "Left Join ProdZonasRastreoCat L on R.CodZona = L.Codigo " +
                      "where R.Cod_Prod is not null and R.Estatus='" + estatus + "' and C.IdZona in (1,2,3,4,7,8,10,11,12,13,14,15) " +
                      "order by R.Fecha, R.Cod_Prod, R.Cod_Campo, S.Sector, R.Num_analisis desc").ToList();
                }

                else if (c_zona != "")
                {
                    resultados = bd.Database.SqlQuery<ClassMuestreo>("select R.Cod_Prod, R.Cod_Campo, ISNULL(S.Sector,0) as Sector, P.Nombre as Productor, T.Descripcion as Tipo, V.Descripcion as Producto, isnull(L.DescZona,'') as Zona, R.Fecha_envio, R.Fecha_entrega, " +
                     "R.Estatus, R.Num_analisis, R.Laboratorio, isnull(convert(varchar, R.LiberacionUSA, 103), '') as LiberacionUSA, isnull(convert(varchar, R.LiberacionEU, 103), '') as LiberacionEU, UPPER(isnull(R.Comentarios, '')) as Comentarios " +
                     "FROM ProdAnalisis_Residuo R " +
                     "LEFT JOIN ProdMuestreoSector S ON R.IdSector = S.Id " +
                     "Left Join ProdProductoresCat P on R.Cod_Prod = P.Cod_Prod " +
                     "Left Join ProdCamposCat C on R.Cod_Prod = C.Cod_Prod and R.Cod_Campo = C.Cod_Campo " +
                     "Left Join CatTiposProd T on C.Tipo = T.Tipo " +
                     "Left Join CatProductos V on C.Tipo = V.Tipo AND C.Producto = V.Producto " +
                     "Left Join ProdZonasRastreoCat L on R.CodZona = L.Codigo " +
                     "where R.Cod_Prod is not null and R.CodZona='" + c_zona + "' and C.IdZona in (1,2,3,4,7,8,10,11,12,13,14,15) " +
                     "order by R.Fecha, R.Cod_Prod, R.Cod_Campo, S.Sector, R.Num_analisis desc").ToList();
                }

                else if (cod_prod != "")
                {
                    resultados = bd.Database.SqlQuery<ClassMuestreo>("select R.Cod_Prod, R.Cod_Campo, ISNULL(S.Sector,0) as Sector, P.Nombre as Productor, T.Descripcion as Tipo, V.Descripcion as Producto, isnull(L.DescZona,'') as Zona, R.Fecha_envio, R.Fecha_entrega, " +
                    "R.Estatus, R.Num_analisis, R.Laboratorio, isnull(convert(varchar, R.LiberacionUSA, 103), '') as LiberacionUSA, isnull(convert(varchar, R.LiberacionEU, 103), '') as LiberacionEU, UPPER(isnull(R.Comentarios, '')) as Comentarios " +
                    "FROM ProdAnalisis_Residuo R " +
                    "LEFT JOIN ProdMuestreoSector S ON R.IdSector = S.Id " +
                    "Left Join ProdProductoresCat P on R.Cod_Prod = P.Cod_Prod " +
                    "Left Join ProdCamposCat C on R.Cod_Prod = C.Cod_Prod and R.Cod_Campo = C.Cod_Campo " +
                    "Left Join CatTiposProd T on C.Tipo = T.Tipo " +
                    "Left Join CatProductos V on C.Tipo = V.Tipo AND C.Producto = V.Producto " +
                    "Left Join ProdZonasRastreoCat L on R.CodZona = L.Codigo " +
                    "where R.Cod_Prod is not null and R.Cod_Prod='" + cod_prod + "' and C.IdZona in (1,2,3,4,7,8,10,11,12,13,14,15) " +
                    "order by R.Fecha, R.Cod_Prod, R.Cod_Campo, S.Sector, R.Num_analisis desc").ToList();
                }

                else
                {
                    resultados = bd.Database.SqlQuery<ClassMuestreo>("select R.Cod_Prod, R.Cod_Campo, ISNULL(S.Sector,0) as Sector, P.Nombre as Productor, T.Descripcion as Tipo, V.Descripcion as Producto, isnull(L.DescZona,'') as Zona, R.Fecha_envio, R.Fecha_entrega, " +
                    "R.Estatus, R.Num_analisis, R.Laboratorio, isnull(convert(varchar, R.LiberacionUSA, 103), '') as LiberacionUSA, isnull(convert(varchar, R.LiberacionEU, 103), '') as LiberacionEU, UPPER(isnull(R.Comentarios, '')) as Comentarios " +
                    "FROM ProdAnalisis_Residuo R " +
                    "LEFT JOIN ProdMuestreoSector S ON R.IdSector = S.Id " +
                    "Left Join ProdProductoresCat P on R.Cod_Prod = P.Cod_Prod " +
                    "Left Join ProdCamposCat C on R.Cod_Prod = C.Cod_Prod and R.Cod_Campo = C.Cod_Campo " +
                    "Left Join CatTiposProd T on C.Tipo = T.Tipo " +
                    "Left Join CatProductos V on C.Tipo = V.Tipo AND C.Producto = V.Producto " +
                    "Left Join ProdZonasRastreoCat L on R.CodZona = L.Codigo " +
                    "where R.Cod_Prod is not null and C.IdZona in (1,2,3,4,7,8,10,11,12,13,14,15) " +
                    "order by R.Fecha, R.Cod_Prod, R.Cod_Campo, S.Sector, R.Num_analisis desc").ToList();
                }
            }
            else if ((short)Session["IdAgen"] == 5)
            {
                if (estatus != "")
                {
                    resultados = bd.Database.SqlQuery<ClassMuestreo>("select R.Cod_Prod, R.Cod_Campo, ISNULL(S.Sector,0) as Sector, P.Nombre as Productor, T.Descripcion as Tipo, V.Descripcion as Producto, isnull(L.DescZona,'') as Zona, R.Fecha_envio, R.Fecha_entrega, " +
                      "R.Estatus, R.Num_analisis, R.Laboratorio, isnull(convert(varchar, R.LiberacionUSA, 103), '') as LiberacionUSA, isnull(convert(varchar, R.LiberacionEU, 103), '') as LiberacionEU, UPPER(isnull(R.Comentarios, '')) as Comentarios " +
                      "FROM ProdAnalisis_Residuo R " +
                      "LEFT JOIN ProdMuestreoSector S ON R.IdSector = S.Id " +
                      "Left Join ProdProductoresCat P on R.Cod_Prod = P.Cod_Prod " +
                      "Left Join ProdCamposCat C on R.Cod_Prod = C.Cod_Prod and R.Cod_Campo = C.Cod_Campo " +
                      "Left Join CatTiposProd T on C.Tipo = T.Tipo " +
                      "Left Join CatProductos V on C.Tipo = V.Tipo AND C.Producto = V.Producto " +
                      "Left Join ProdZonasRastreoCat L on R.CodZona = L.Codigo " +
                      "where R.Cod_Prod is not null and R.Estatus='" + estatus + "'and C.IdZona in (5,6,9)" +
                      "order by R.Fecha, R.Cod_Prod, R.Cod_Campo, S.Sector, R.Num_analisis desc").ToList();
                }

                else if (c_zona != "")
                {
                    resultados = bd.Database.SqlQuery<ClassMuestreo>("select R.Cod_Prod, R.Cod_Campo, ISNULL(S.Sector,0) as Sector, P.Nombre as Productor, T.Descripcion as Tipo, V.Descripcion as Producto, isnull(L.DescZona,'') as Zona, R.Fecha_envio, R.Fecha_entrega, " +
                    "R.Estatus, R.Num_analisis, R.Laboratorio, isnull(convert(varchar, R.LiberacionUSA, 103), '') as LiberacionUSA, isnull(convert(varchar, R.LiberacionEU, 103), '') as LiberacionEU, UPPER(isnull(R.Comentarios, '')) as Comentarios " +
                    "FROM ProdAnalisis_Residuo R " +
                    "LEFT JOIN ProdMuestreoSector S ON R.IdSector = S.Id " +
                    "Left Join ProdProductoresCat P on R.Cod_Prod = P.Cod_Prod " +
                    "Left Join ProdCamposCat C on R.Cod_Prod = C.Cod_Prod and R.Cod_Campo = C.Cod_Campo " +
                    "Left Join CatTiposProd T on C.Tipo = T.Tipo " +
                    "Left Join CatProductos V on C.Tipo = V.Tipo AND C.Producto = V.Producto " +
                    "Left Join ProdZonasRastreoCat L on R.CodZona = L.Codigo " +
                    "where R.Cod_Prod is not null and R.CodZona='" + c_zona + "' and C.IdZona in (5,6,9)" +
                    "order by R.Fecha, R.Cod_Prod, R.Cod_Campo, S.Sector, R.Num_analisis desc").ToList();
                }

                else if (cod_prod != "")
                {
                    resultados = bd.Database.SqlQuery<ClassMuestreo>("select R.Cod_Prod, R.Cod_Campo, ISNULL(S.Sector,0) as Sector, P.Nombre as Productor, T.Descripcion as Tipo, V.Descripcion as Producto, isnull(L.DescZona,'') as Zona, R.Fecha_envio, R.Fecha_entrega, " +
                   "R.Estatus, R.Num_analisis, R.Laboratorio, isnull(convert(varchar, R.LiberacionUSA, 103), '') as LiberacionUSA, isnull(convert(varchar, R.LiberacionEU, 103), '') as LiberacionEU, UPPER(isnull(R.Comentarios, '')) as Comentarios " +
                   "FROM ProdAnalisis_Residuo R " +
                   "LEFT JOIN ProdMuestreoSector S ON R.IdSector = S.Id " +
                   "Left Join ProdProductoresCat P on R.Cod_Prod = P.Cod_Prod " +
                   "Left Join ProdCamposCat C on R.Cod_Prod = C.Cod_Prod and R.Cod_Campo = C.Cod_Campo " +
                   "Left Join CatTiposProd T on C.Tipo = T.Tipo " +
                   "Left Join CatProductos V on C.Tipo = V.Tipo AND C.Producto = V.Producto " +
                   "Left Join ProdZonasRastreoCat L on R.CodZona = L.Codigo " +
                   "where R.Cod_Prod is not null and R.Cod_Prod='" + cod_prod + "' and C.IdZona in (5,6,9)" +
                   "order by R.Fecha, R.Cod_Prod, R.Cod_Campo, S.Sector, R.Num_analisis desc").ToList();
                }

                else
                {
                    resultados = bd.Database.SqlQuery<ClassMuestreo>("select R.Cod_Prod, R.Cod_Campo, ISNULL(S.Sector,0) as Sector, P.Nombre as Productor, T.Descripcion as Tipo, V.Descripcion as Producto, isnull(L.DescZona,'') as Zona, R.Fecha_envio, R.Fecha_entrega, " +
                    "R.Estatus, R.Num_analisis, R.Laboratorio, isnull(convert(varchar, R.LiberacionUSA, 103), '') as LiberacionUSA, isnull(convert(varchar, R.LiberacionEU, 103), '') as LiberacionEU, UPPER(isnull(R.Comentarios, '')) as Comentarios " +
                    "FROM ProdAnalisis_Residuo R " +
                    "LEFT JOIN ProdMuestreoSector S ON R.IdSector = S.Id " +
                    "Left Join ProdProductoresCat P on R.Cod_Prod = P.Cod_Prod " +
                    "Left Join ProdCamposCat C on R.Cod_Prod = C.Cod_Prod and R.Cod_Campo = C.Cod_Campo " +
                    "Left Join CatTiposProd T on C.Tipo = T.Tipo " +
                    "Left Join CatProductos V on C.Tipo = V.Tipo AND C.Producto = V.Producto " +
                    "Left Join ProdZonasRastreoCat L on R.CodZona = L.Codigo " +
                    "where R.Cod_Prod is not null and C.IdZona in (5,6,9)" +
                    "order by R.Fecha, R.Cod_Prod, R.Cod_Campo, S.Sector, R.Num_analisis desc").ToList();
                }
            }
            else
            {
                if (estatus != "")
                {
                    resultados = bd.Database.SqlQuery<ClassMuestreo>("select R.Cod_Prod, R.Cod_Campo, ISNULL(S.Sector,0) as Sector, P.Nombre as Productor, T.Descripcion as Tipo, V.Descripcion as Producto, isnull(L.DescZona,'') as Zona, R.Fecha_envio, R.Fecha_entrega, " +
                        "R.Estatus, R.Num_analisis, R.Laboratorio, isnull(convert(varchar, R.LiberacionUSA, 103), '') as LiberacionUSA, isnull(convert(varchar, R.LiberacionEU, 103), '') as LiberacionEU, UPPER(isnull(R.Comentarios, '')) as Comentarios " +
                        "FROM ProdAnalisis_Residuo R " +
                        "LEFT JOIN ProdMuestreoSector S ON R.IdSector = S.Id " +
                        "Left Join ProdProductoresCat P on R.Cod_Prod = P.Cod_Prod " +
                        "Left Join ProdCamposCat C on R.Cod_Prod = C.Cod_Prod and R.Cod_Campo = C.Cod_Campo " +
                        "Left Join CatTiposProd T on C.Tipo = T.Tipo " +
                        "Left Join CatProductos V on C.Tipo = V.Tipo AND C.Producto = V.Producto " +
                        "Left Join ProdZonasRastreoCat L on R.CodZona = L.Codigo " +
                        "where R.Cod_Prod is not null and R.Estatus='" + estatus + "' and (C.IdAgen=" + (short)Session["IdAgen"] + " or C.IdAgenC=" + (short)Session["IdAgen"] + " or C.IdAgenI=" + (short)Session["IdAgen"] + ") " +
                        "order by R.Fecha, R.Cod_Prod, R.Cod_Campo, S.Sector, R.Num_analisis desc").ToList();
                }

                else if (c_zona != "")
                {
                    resultados = bd.Database.SqlQuery<ClassMuestreo>("select R.Cod_Prod, R.Cod_Campo, ISNULL(S.Sector,0) as Sector, P.Nombre as Productor, T.Descripcion as Tipo, V.Descripcion as Producto, isnull(L.DescZona,'') as Zona, R.Fecha_envio, R.Fecha_entrega, " +
                      "R.Estatus, R.Num_analisis, R.Laboratorio, isnull(convert(varchar, R.LiberacionUSA, 103), '') as LiberacionUSA, isnull(convert(varchar, R.LiberacionEU, 103), '') as LiberacionEU, UPPER(isnull(R.Comentarios, '')) as Comentarios " +
                      "FROM ProdAnalisis_Residuo R " +
                      "LEFT JOIN ProdMuestreoSector S ON R.IdSector = S.Id " +
                      "Left Join ProdProductoresCat P on R.Cod_Prod = P.Cod_Prod " +
                      "Left Join ProdCamposCat C on R.Cod_Prod = C.Cod_Prod and R.Cod_Campo = C.Cod_Campo " +
                      "Left Join CatTiposProd T on C.Tipo = T.Tipo " +
                      "Left Join CatProductos V on C.Tipo = V.Tipo AND C.Producto = V.Producto " +
                      "Left Join ProdZonasRastreoCat L on R.CodZona = L.Codigo " +
                      "where R.Cod_Prod is not null and R.CodZona='" + c_zona + "' and (C.IdAgen=" + (short)Session["IdAgen"] + " or C.IdAgenC=" + (short)Session["IdAgen"] + " or C.IdAgenI=" + (short)Session["IdAgen"] + ") " +
                      "order by R.Fecha, R.Cod_Prod, R.Cod_Campo, S.Sector, R.Num_analisis desc").ToList();
                }

                else if (cod_prod != "")
                {
                    resultados = bd.Database.SqlQuery<ClassMuestreo>("select R.Cod_Prod, R.Cod_Campo, ISNULL(S.Sector,0) as Sector, P.Nombre as Productor, T.Descripcion as Tipo, V.Descripcion as Producto, isnull(L.DescZona,'') as Zona, R.Fecha_envio, R.Fecha_entrega, " +
                     "R.Estatus, R.Num_analisis, R.Laboratorio, isnull(convert(varchar, R.LiberacionUSA, 103), '') as LiberacionUSA, isnull(convert(varchar, R.LiberacionEU, 103), '') as LiberacionEU, UPPER(isnull(R.Comentarios, '')) as Comentarios " +
                     "FROM ProdAnalisis_Residuo R " +
                     "LEFT JOIN ProdMuestreoSector S ON R.IdSector = S.Id " +
                     "Left Join ProdProductoresCat P on R.Cod_Prod = P.Cod_Prod " +
                     "Left Join ProdCamposCat C on R.Cod_Prod = C.Cod_Prod and R.Cod_Campo = C.Cod_Campo " +
                     "Left Join CatTiposProd T on C.Tipo = T.Tipo " +
                     "Left Join CatProductos V on C.Tipo = V.Tipo AND C.Producto = V.Producto " +
                     "Left Join ProdZonasRastreoCat L on R.CodZona = L.Codigo " +
                     "where R.Cod_Prod is not null and R.Cod_Prod='" + cod_prod + "' and (C.IdAgen=" + (short)Session["IdAgen"] + " or C.IdAgenC=" + (short)Session["IdAgen"] + " or C.IdAgenI=" + (short)Session["IdAgen"] + ") " +
                     "order by R.Fecha, R.Cod_Prod, R.Cod_Campo, S.Sector, R.Num_analisis desc").ToList();
                }

                else
                {
                    resultados = bd.Database.SqlQuery<ClassMuestreo>("select R.Cod_Prod, R.Cod_Campo, ISNULL(S.Sector,0) as Sector, P.Nombre as Productor, T.Descripcion as Tipo, V.Descripcion as Producto, isnull(L.DescZona,'') as Zona, R.Fecha_envio, R.Fecha_entrega, " +
                    "R.Estatus, R.Num_analisis, R.Laboratorio, isnull(convert(varchar, R.LiberacionUSA, 103), '') as LiberacionUSA, isnull(convert(varchar, R.LiberacionEU, 103), '') as LiberacionEU, UPPER(isnull(R.Comentarios, '')) as Comentarios " +
                    "FROM ProdAnalisis_Residuo R " +
                    "LEFT JOIN ProdMuestreoSector S ON R.IdSector = S.Id " +
                    "Left Join ProdProductoresCat P on R.Cod_Prod = P.Cod_Prod " +
                    "Left Join ProdCamposCat C on R.Cod_Prod = C.Cod_Prod and R.Cod_Campo = C.Cod_Campo " +
                    "Left Join CatTiposProd T on C.Tipo = T.Tipo " +
                    "Left Join CatProductos V on C.Tipo = V.Tipo AND C.Producto = V.Producto " +
                    "Left Join ProdZonasRastreoCat L on R.CodZona = L.Codigo " +
                    "where R.Cod_Prod is not null and (C.IdAgen=" + (short)Session["IdAgen"] + " or C.IdAgenC=" + (short)Session["IdAgen"] + " or C.IdAgenI=" + (short)Session["IdAgen"] + ") " +
                    "order by R.Fecha, R.Cod_Prod, R.Cod_Campo, S.Sector, R.Num_analisis desc").ToList();
                }
            }
            return Json(resultados, JsonRequestBehavior.AllowGet);
        }

        public ActionResult Evaluacion()
        {
            if (Session["Nombre"] != null)
            {
                ViewData["Nombre"] = Session["Nombre"].ToString();
                List<SelectListItem> lst_Status = new List<SelectListItem>();
                lst_Status.Add(new SelectListItem() { Text = "--Resultado del analisis--", Value = null });
                lst_Status.Add(new SelectListItem() { Text = "CON RESIDUOS", Value = "R" });
                lst_Status.Add(new SelectListItem() { Text = "EN PROCESO", Value = "P" });
                lst_Status.Add(new SelectListItem() { Text = "FUERA DE LIMITE", Value = "F" });
                lst_Status.Add(new SelectListItem() { Text = "LIBERADO", Value = "L" });

                ViewBag.List_Status = lst_Status;
            }
            else
            {
                return RedirectToAction("Index", "Login");
            }
            return View();
        }

        public JsonResult EvaluacionList(string estatus = "", string cod_prod = "")
        {
            List<ClassMuestreo> resultados = null;
            if ((short)Session["IdAgen"] == 205 || Session["IdAgen"].ToString() == "153" || Session["IdAgen"].ToString() == "281")
            {
                if (estatus != "")
                {
                    resultados = bd.Database.SqlQuery<ClassMuestreo>("Select C.Asesor, C.Cod_Prod, C.Productor, C.Campo, C.Ubicacion, C.Fecha_solicitud, C.Inicio_cosecha, C.Fecha_real, C.Fecha_analisis, " +
                        "(case when C.Dias <> 0 then 'ENTREGANDO' ELSE 'EN ESPERA' END) AS Estatus ,C.Dias, isnull((case when C.Estatus='R' then 'CON RESIDUOS' else(case when C.Estatus='P' then 'EN PROCESO' else(case when C.Estatus='F' then 'FUERA DE LIMITE' else(case when C.Estatus='L' then 'LIBERADO' end) end) end) end),'') as Analisis " +
                        "from(" +
                        "SELECT A.Nombre as Asesor, M.Cod_Prod, P.Nombre as Productor, C.Descripcion as Campo, isnull(L.Descripcion, '') as Ubicacion, M.Fecha_solicitud, M.Inicio_cosecha, isnull(convert(varchar, E.Fecha, 101), '') as Fecha_real, R1.FechaA as Fecha_analisis, isnull(DATEDIFF(day, R1.FechaA,E.Fecha), 0) as Dias, R.Estatus, C.IdAgen, C.IdAgenC, C.IdAgenI " +
                        "FROM ProdMuestreo M " +
                        "LEFT JOIN ProdAgenteCat A on M.IdAgen = A.IdAgen " +
                        "LEFT JOIN ProdProductoresCat P on M.Cod_Prod = P.Cod_Prod " +
                        "LEFT JOIN ProdCamposCat C on M.Cod_Prod = C.Cod_Prod and M.Cod_Campo = C.Cod_Campo " +
                        "LEFT JOIN CatLocalidades L on C.CodLocalidad = L.CodLocalidad " +
                        "LEFT JOIN (Select distinct Id_Muestreo, max(Fecha) as FechaA FROM ProdAnalisis_Residuo group by Id_Muestreo)R1 ON M.Id=R1.Id_Muestreo " +
                        "LEFT JOIN (Select Estatus, Fecha FROM ProdAnalisis_Residuo)R ON R.Fecha=R1.FechaA " +
                        "LEFT JOIN(Select Cod_Prod, Cod_Campo, min(Fecha) as Fecha, Temporada from UV_ProdRecepcion " +
                        "where Temporada = (Select temporada from CatSemanas where getdate() between Inicio and Fin) " +
                        "group by Cod_Prod, Cod_Campo, Temporada)E on M.Cod_Prod = E.Cod_Prod and M.Cod_Campo = E.Cod_Campo " +
                        "where R.Estatus <> '')C " +
                        "Where C.Estatus='" + estatus + "'" +
                        "order by C.Fecha_solicitud desc").ToList();
                }
                else if (cod_prod != "")
                {
                    resultados = bd.Database.SqlQuery<ClassMuestreo>("Select C.Asesor, C.Cod_Prod, C.Productor, C.Campo, C.Ubicacion, C.Fecha_solicitud, C.Inicio_cosecha, C.Fecha_real, C.Fecha_analisis, " +
                        "(case when C.Dias <> 0 then 'ENTREGANDO' ELSE 'EN ESPERA' END) AS Estatus ,C.Dias, isnull((case when C.Estatus='R' then 'CON RESIDUOS' else(case when C.Estatus='P' then 'EN PROCESO' else(case when C.Estatus='F' then 'FUERA DE LIMITE' else(case when C.Estatus='L' then 'LIBERADO' end) end) end) end),'') as Analisis " +
                        "from(" +
                        "SELECT A.Nombre as Asesor, M.Cod_Prod, P.Nombre as Productor, C.Descripcion as Campo, isnull(L.Descripcion, '') as Ubicacion, M.Fecha_solicitud, M.Inicio_cosecha, isnull(convert(varchar, E.Fecha, 101), '') as Fecha_real, R1.FechaA as Fecha_analisis, isnull(DATEDIFF(day, R1.FechaA,E.Fecha), 0) as Dias, R.Estatus, C.IdAgen, C.IdAgenC, C.IdAgenI " +
                        "FROM ProdMuestreo M " +
                        "LEFT JOIN ProdAgenteCat A on M.IdAgen = A.IdAgen " +
                        "LEFT JOIN ProdProductoresCat P on M.Cod_Prod = P.Cod_Prod " +
                        "LEFT JOIN ProdCamposCat C on M.Cod_Prod = C.Cod_Prod and M.Cod_Campo = C.Cod_Campo " +
                        "LEFT JOIN CatLocalidades L on C.CodLocalidad = L.CodLocalidad " +
                        "LEFT JOIN (Select distinct Id_Muestreo, max(Fecha) as FechaA FROM ProdAnalisis_Residuo group by Id_Muestreo)R1 ON M.Id=R1.Id_Muestreo " +
                        "LEFT JOIN (Select Estatus, Fecha FROM ProdAnalisis_Residuo)R ON R.Fecha=R1.FechaA " +
                        "LEFT JOIN(Select Cod_Prod, Cod_Campo, min(Fecha) as Fecha, Temporada from UV_ProdRecepcion " +
                        "where Temporada = (Select temporada from CatSemanas where getdate() between Inicio and Fin) " +
                        "group by Cod_Prod, Cod_Campo, Temporada)E on M.Cod_Prod = E.Cod_Prod and M.Cod_Campo = E.Cod_Campo " +
                        "where R.Estatus <> '')C " +
                        "where C.Cod_Prod='" + cod_prod + "' " +
                        "order by C.Fecha_solicitud desc").ToList();
                }
                else
                {
                    resultados = bd.Database.SqlQuery<ClassMuestreo>("Select C.Asesor, C.Cod_Prod, C.Productor, C.Campo, C.Ubicacion, C.Fecha_solicitud, C.Inicio_cosecha, C.Fecha_real, C.Fecha_analisis, " +
                        "(case when C.Dias <> 0 then 'ENTREGANDO' ELSE 'EN ESPERA' END) AS Estatus ,C.Dias, isnull((case when C.Estatus='R' then 'CON RESIDUOS' else(case when C.Estatus='P' then 'EN PROCESO' else(case when C.Estatus='F' then 'FUERA DE LIMITE' else(case when C.Estatus='L' then 'LIBERADO' end) end) end) end),'') as Analisis " +
                       "from(" +
                       "SELECT A.Nombre as Asesor, M.Cod_Prod, P.Nombre as Productor, C.Descripcion as Campo, isnull(L.Descripcion, '') as Ubicacion, M.Fecha_solicitud, M.Inicio_cosecha, isnull(convert(varchar, E.Fecha, 101), '') as Fecha_real, R1.FechaA as Fecha_analisis, isnull(DATEDIFF(day, R1.FechaA,E.Fecha), 0) as Dias, R.Estatus, C.IdAgen, C.IdAgenC, C.IdAgenI " +
                       "FROM ProdMuestreo M " +
                       "LEFT JOIN ProdAgenteCat A on M.IdAgen = A.IdAgen " +
                       "LEFT JOIN ProdProductoresCat P on M.Cod_Prod = P.Cod_Prod " +
                       "LEFT JOIN ProdCamposCat C on M.Cod_Prod = C.Cod_Prod and M.Cod_Campo = C.Cod_Campo " +
                       "LEFT JOIN CatLocalidades L on C.CodLocalidad = L.CodLocalidad " +
                       "LEFT JOIN (Select distinct Id_Muestreo, max(Fecha) as FechaA FROM ProdAnalisis_Residuo group by Id_Muestreo)R1 ON M.Id=R1.Id_Muestreo " +
                       "LEFT JOIN (Select Estatus, Fecha FROM ProdAnalisis_Residuo)R ON R.Fecha=R1.FechaA " +
                       "LEFT JOIN(Select Cod_Prod, Cod_Campo, min(Fecha) as Fecha, Temporada from UV_ProdRecepcion " +
                       "where Temporada = (Select temporada from CatSemanas where getdate() between Inicio and Fin) " +
                       "group by Cod_Prod, Cod_Campo, Temporada)E on M.Cod_Prod = E.Cod_Prod and M.Cod_Campo = E.Cod_Campo " +
                       "where R.Estatus <> '')C " +
                       "order by C.Fecha_solicitud desc").ToList();
                }
            }
            else if ((short)Session["IdAgen"] == 1)
            {
                if (estatus != "")
                {
                    resultados = bd.Database.SqlQuery<ClassMuestreo>("Select C.Asesor, C.Cod_Prod, C.Productor, C.Campo, C.Ubicacion, C.Fecha_solicitud, C.Inicio_cosecha, C.Fecha_real, C.Fecha_analisis, " +
                        "(case when C.Dias <> 0 then 'ENTREGANDO' ELSE 'EN ESPERA' END) AS Estatus ,C.Dias, isnull((case when C.Estatus='R' then 'CON RESIDUOS' else(case when C.Estatus='P' then 'EN PROCESO' else(case when C.Estatus='F' then 'FUERA DE LIMITE' else(case when C.Estatus='L' then 'LIBERADO' end) end) end) end),'') as Analisis " +
                      "from(" +
                      "SELECT A.Nombre as Asesor, M.Cod_Prod, P.Nombre as Productor, C.Descripcion as Campo, isnull(L.Descripcion, '') as Ubicacion, M.Fecha_solicitud, M.Inicio_cosecha, isnull(convert(varchar, E.Fecha, 101), '') as Fecha_real, R1.FechaA as Fecha_analisis, isnull(DATEDIFF(day, R1.FechaA,E.Fecha), 0) as Dias, R.Estatus, C.IdAgen, C.IdAgenC, C.IdAgenI " +
                      "FROM ProdMuestreo M " +
                        "LEFT JOIN ProdAgenteCat A on M.IdAgen = A.IdAgen " +
                        "LEFT JOIN ProdProductoresCat P on M.Cod_Prod = P.Cod_Prod " +
                        "LEFT JOIN ProdCamposCat C on M.Cod_Prod = C.Cod_Prod and M.Cod_Campo = C.Cod_Campo " +
                        "LEFT JOIN CatLocalidades L on C.CodLocalidad = L.CodLocalidad " +
                         "LEFT JOIN (Select distinct Id_Muestreo, max(Fecha) as FechaA FROM ProdAnalisis_Residuo group by Id_Muestreo)R1 ON M.Id=R1.Id_Muestreo " +
                        "LEFT JOIN (Select Estatus, Fecha FROM ProdAnalisis_Residuo)R ON R.Fecha=R1.FechaA " +
                        "LEFT JOIN(Select Cod_Prod, Cod_Campo, min(Fecha) as Fecha, Temporada from UV_ProdRecepcion " +
                        "where Temporada = (Select temporada from CatSemanas where getdate() between Inicio and Fin) " +
                        "group by Cod_Prod, Cod_Campo, Temporada)E on M.Cod_Prod = E.Cod_Prod and M.Cod_Campo = E.Cod_Campo " +
                        "where R.Estatus <> '' and C.IdZona in (1,2,3,4,7,8,10,11,12,13,14,15))C " +
                        "Where C.Estatus='" + estatus + "'" +
                        "order by C.Fecha_solicitud desc").ToList();
                }
                else if (cod_prod != "")
                {
                    resultados = bd.Database.SqlQuery<ClassMuestreo>("Select C.Asesor, C.Cod_Prod, C.Productor, C.Campo, C.Ubicacion, C.Fecha_solicitud, C.Inicio_cosecha, C.Fecha_real, C.Fecha_analisis, " +
                        "(case when C.Dias <> 0 then 'ENTREGANDO' ELSE 'EN ESPERA' END) AS Estatus ,C.Dias, isnull((case when C.Estatus='R' then 'CON RESIDUOS' else(case when C.Estatus='P' then 'EN PROCESO' else(case when C.Estatus='F' then 'FUERA DE LIMITE' else(case when C.Estatus='L' then 'LIBERADO' end) end) end) end),'') as Analisis " +
                       "from(" +
                       "SELECT A.Nombre as Asesor, M.Cod_Prod, P.Nombre as Productor, C.Descripcion as Campo, isnull(L.Descripcion, '') as Ubicacion, M.Fecha_solicitud, M.Inicio_cosecha, isnull(convert(varchar, E.Fecha, 101), '') as Fecha_real, R1.FechaA as Fecha_analisis, isnull(DATEDIFF(day, R1.FechaA,E.Fecha), 0) as Dias, R.Estatus, C.IdAgen, C.IdAgenC, C.IdAgenI " +
                        "FROM ProdMuestreo M " +
                        "LEFT JOIN ProdAgenteCat A on M.IdAgen = A.IdAgen " +
                        "LEFT JOIN ProdProductoresCat P on M.Cod_Prod = P.Cod_Prod " +
                        "LEFT JOIN ProdCamposCat C on M.Cod_Prod = C.Cod_Prod and M.Cod_Campo = C.Cod_Campo " +
                        "LEFT JOIN CatLocalidades L on C.CodLocalidad = L.CodLocalidad " +
                        "LEFT JOIN (Select distinct Id_Muestreo, max(Fecha) as FechaA FROM ProdAnalisis_Residuo group by Id_Muestreo)R1 ON M.Id=R1.Id_Muestreo " +
                        "LEFT JOIN (Select Estatus, Fecha FROM ProdAnalisis_Residuo)R ON R.Fecha=R1.FechaA " +
                        "LEFT JOIN(Select Cod_Prod, Cod_Campo, min(Fecha) as Fecha, Temporada from UV_ProdRecepcion " +
                        "where Temporada = (Select temporada from CatSemanas where getdate() between Inicio and Fin) " +
                        "group by Cod_Prod, Cod_Campo, Temporada)E on M.Cod_Prod = E.Cod_Prod and M.Cod_Campo = E.Cod_Campo " +
                        "where R.Estatus <> '' and C.IdZona in (1,2,3,4,7,8,10,11,12,13,14,15))C " +
                        "where C.Cod_Prod='" + cod_prod + "' " +
                        "order by C.Fecha_solicitud desc").ToList();
                }
                else
                {
                    resultados = bd.Database.SqlQuery<ClassMuestreo>("Select C.Asesor, C.Cod_Prod, C.Productor, C.Campo, C.Ubicacion, C.Fecha_solicitud, C.Inicio_cosecha, C.Fecha_real, C.Fecha_analisis, " +
                        "(case when C.Dias <> 0 then 'ENTREGANDO' ELSE 'EN ESPERA' END) AS Estatus ,C.Dias, isnull((case when C.Estatus='R' then 'CON RESIDUOS' else(case when C.Estatus='P' then 'EN PROCESO' else(case when C.Estatus='F' then 'FUERA DE LIMITE' else(case when C.Estatus='L' then 'LIBERADO' end) end) end) end),'') as Analisis " +
                       "from(" +
                       "SELECT A.Nombre as Asesor, M.Cod_Prod, P.Nombre as Productor, C.Descripcion as Campo, isnull(L.Descripcion, '') as Ubicacion, M.Fecha_solicitud, M.Inicio_cosecha, isnull(convert(varchar, E.Fecha, 101), '') as Fecha_real, R1.FechaA as Fecha_analisis, isnull(DATEDIFF(day, R1.FechaA,E.Fecha), 0) as Dias, R.Estatus, C.IdAgen, C.IdAgenC, C.IdAgenI " +
                        "FROM ProdMuestreo M " +
                       "LEFT JOIN ProdAgenteCat A on M.IdAgen = A.IdAgen " +
                       "LEFT JOIN ProdProductoresCat P on M.Cod_Prod = P.Cod_Prod " +
                       "LEFT JOIN ProdCamposCat C on M.Cod_Prod = C.Cod_Prod and M.Cod_Campo = C.Cod_Campo " +
                       "LEFT JOIN CatLocalidades L on C.CodLocalidad = L.CodLocalidad " +
                       "LEFT JOIN (Select distinct Id_Muestreo, max(Fecha) as FechaA FROM ProdAnalisis_Residuo group by Id_Muestreo)R1 ON M.Id=R1.Id_Muestreo " +
                        "LEFT JOIN (Select Estatus, Fecha FROM ProdAnalisis_Residuo)R ON R.Fecha=R1.FechaA " +
                       "LEFT JOIN(Select Cod_Prod, Cod_Campo, min(Fecha) as Fecha, Temporada from UV_ProdRecepcion " +
                       "where Temporada = (Select temporada from CatSemanas where getdate() between Inicio and Fin) " +
                       "group by Cod_Prod, Cod_Campo, Temporada)E on M.Cod_Prod = E.Cod_Prod and M.Cod_Campo = E.Cod_Campo " +
                       "where R.Estatus <> '' and C.IdZona in (1,2,3,4,7,8,10,11,12,13,14,15))C " +
                       "order by C.Fecha_solicitud desc").ToList();
                }

            }
            else if ((short)Session["IdAgen"] == 5)
            {
                if (estatus != "")
                {
                    resultados = bd.Database.SqlQuery<ClassMuestreo>("Select C.Asesor, C.Cod_Prod, C.Productor, C.Campo, C.Ubicacion, C.Fecha_solicitud, C.Inicio_cosecha, C.Fecha_real, C.Fecha_analisis, " +
                        "(case when C.Dias <> 0 then 'ENTREGANDO' ELSE 'EN ESPERA' END) AS Estatus ,C.Dias, isnull((case when C.Estatus='R' then 'CON RESIDUOS' else(case when C.Estatus='P' then 'EN PROCESO' else(case when C.Estatus='F' then 'FUERA DE LIMITE' else(case when C.Estatus='L' then 'LIBERADO' end) end) end) end),'') as Analisis " +
                       "from(" +
                       "SELECT A.Nombre as Asesor, M.Cod_Prod, P.Nombre as Productor, C.Descripcion as Campo, isnull(L.Descripcion, '') as Ubicacion, M.Fecha_solicitud, M.Inicio_cosecha, isnull(convert(varchar, E.Fecha, 101), '') as Fecha_real, R1.FechaA as Fecha_analisis, isnull(DATEDIFF(day, R1.FechaA,E.Fecha), 0) as Dias, R.Estatus, C.IdAgen, C.IdAgenC, C.IdAgenI " +
                       "FROM ProdMuestreo M " +
                        "LEFT JOIN ProdAgenteCat A on M.IdAgen = A.IdAgen " +
                        "LEFT JOIN ProdProductoresCat P on M.Cod_Prod = P.Cod_Prod " +
                        "LEFT JOIN ProdCamposCat C on M.Cod_Prod = C.Cod_Prod and M.Cod_Campo = C.Cod_Campo " +
                        "LEFT JOIN CatLocalidades L on C.CodLocalidad = L.CodLocalidad " +
                        "LEFT JOIN (Select distinct Id_Muestreo, max(Fecha) as FechaA FROM ProdAnalisis_Residuo group by Id_Muestreo)R1 ON M.Id=R1.Id_Muestreo " +
                        "LEFT JOIN (Select Estatus, Fecha FROM ProdAnalisis_Residuo)R ON R.Fecha=R1.FechaA " +
                        "LEFT JOIN(Select Cod_Prod, Cod_Campo, min(Fecha) as Fecha, Temporada from UV_ProdRecepcion " +
                        "where Temporada = (Select temporada from CatSemanas where getdate() between Inicio and Fin) " +
                        "group by Cod_Prod, Cod_Campo, Temporada)E on M.Cod_Prod = E.Cod_Prod and M.Cod_Campo = E.Cod_Campo " +
                        "where R.Estatus <> '' and C.IdZona in (5,6,9))C " +
                        "Where C.Estatus='" + estatus + "'" +
                        "order by C.Fecha_solicitud desc").ToList();
                }
                else if (cod_prod != "")
                {
                    resultados = bd.Database.SqlQuery<ClassMuestreo>("Select C.Asesor, C.Cod_Prod, C.Productor, C.Campo, C.Ubicacion, C.Fecha_solicitud, C.Inicio_cosecha, C.Fecha_real, C.Fecha_analisis, " +
                        "(case when C.Dias <> 0 then 'ENTREGANDO' ELSE 'EN ESPERA' END) AS Estatus ,C.Dias, isnull((case when C.Estatus='R' then 'CON RESIDUOS' else(case when C.Estatus='P' then 'EN PROCESO' else(case when C.Estatus='F' then 'FUERA DE LIMITE' else(case when C.Estatus='L' then 'LIBERADO' end) end) end) end),'') as Analisis " +
                       "from(" +
                       "SELECT A.Nombre as Asesor, M.Cod_Prod, P.Nombre as Productor, C.Descripcion as Campo, isnull(L.Descripcion, '') as Ubicacion, M.Fecha_solicitud, M.Inicio_cosecha, isnull(convert(varchar, E.Fecha, 101), '') as Fecha_real, R1.FechaA as Fecha_analisis, isnull(DATEDIFF(day, R1.FechaA,E.Fecha), 0) as Dias, R.Estatus, C.IdAgen, C.IdAgenC, C.IdAgenI " +
                        "FROM ProdMuestreo M " +
                        "LEFT JOIN ProdAgenteCat A on M.IdAgen = A.IdAgen " +
                        "LEFT JOIN ProdProductoresCat P on M.Cod_Prod = P.Cod_Prod " +
                        "LEFT JOIN ProdCamposCat C on M.Cod_Prod = C.Cod_Prod and M.Cod_Campo = C.Cod_Campo " +
                        "LEFT JOIN CatLocalidades L on C.CodLocalidad = L.CodLocalidad " +
                        "LEFT JOIN (Select distinct Id_Muestreo, max(Fecha) as FechaA FROM ProdAnalisis_Residuo group by Id_Muestreo)R1 ON M.Id=R1.Id_Muestreo " +
                        "LEFT JOIN (Select Estatus, Fecha FROM ProdAnalisis_Residuo)R ON R.Fecha=R1.FechaA " +
                        "LEFT JOIN(Select Cod_Prod, Cod_Campo, min(Fecha) as Fecha, Temporada from UV_ProdRecepcion " +
                        "where Temporada = (Select temporada from CatSemanas where getdate() between Inicio and Fin) " +
                        "group by Cod_Prod, Cod_Campo, Temporada)E on M.Cod_Prod = E.Cod_Prod and M.Cod_Campo = E.Cod_Campo " +
                        "where R.Estatus <> '' and C.IdZona in (5,6,9))C " +
                        "where C.Cod_Prod='" + cod_prod + "' " +
                        "order by C.Fecha_solicitud desc").ToList();
                }
                else
                {
                    resultados = bd.Database.SqlQuery<ClassMuestreo>("Select C.Asesor, C.Cod_Prod, C.Productor, C.Campo, C.Ubicacion, C.Fecha_solicitud, C.Inicio_cosecha, C.Fecha_real, C.Fecha_analisis, " +
                        "(case when C.Dias <> 0 then 'ENTREGANDO' ELSE 'EN ESPERA' END) AS Estatus ,C.Dias, isnull((case when C.Estatus='R' then 'CON RESIDUOS' else(case when C.Estatus='P' then 'EN PROCESO' else(case when C.Estatus='F' then 'FUERA DE LIMITE' else(case when C.Estatus='L' then 'LIBERADO' end) end) end) end),'') as Analisis " +
                       "from(" +
                       "SELECT A.Nombre as Asesor, M.Cod_Prod, P.Nombre as Productor, C.Descripcion as Campo, isnull(L.Descripcion, '') as Ubicacion, M.Fecha_solicitud, M.Inicio_cosecha, isnull(convert(varchar, E.Fecha, 101), '') as Fecha_real, R1.FechaA as Fecha_analisis, isnull(DATEDIFF(day, R1.FechaA,E.Fecha), 0) as Dias, R.Estatus, C.IdAgen, C.IdAgenC, C.IdAgenI " +
                       "FROM ProdMuestreo M " +
                       "LEFT JOIN ProdAgenteCat A on M.IdAgen = A.IdAgen " +
                       "LEFT JOIN ProdProductoresCat P on M.Cod_Prod = P.Cod_Prod " +
                       "LEFT JOIN ProdCamposCat C on M.Cod_Prod = C.Cod_Prod and M.Cod_Campo = C.Cod_Campo " +
                       "LEFT JOIN CatLocalidades L on C.CodLocalidad = L.CodLocalidad " +
                       "LEFT JOIN (Select distinct Id_Muestreo, max(Fecha) as FechaA FROM ProdAnalisis_Residuo group by Id_Muestreo)R1 ON M.Id=R1.Id_Muestreo " +
                       "LEFT JOIN (Select Estatus, Fecha FROM ProdAnalisis_Residuo)R ON R.Fecha=R1.FechaA " +
                       "LEFT JOIN(Select Cod_Prod, Cod_Campo, min(Fecha) as Fecha, Temporada from UV_ProdRecepcion " +
                       "where Temporada = (Select temporada from CatSemanas where getdate() between Inicio and Fin) " +
                       "group by Cod_Prod, Cod_Campo, Temporada)E on M.Cod_Prod = E.Cod_Prod and M.Cod_Campo = E.Cod_Campo " +
                       "where R.Estatus <> '' and C.IdZona in (5,6,9))C " +
                       "order by C.Fecha_solicitud desc").ToList();
                }
            }
            else
            {
                if (estatus != "")
                {
                    resultados = bd.Database.SqlQuery<ClassMuestreo>("Select C.Asesor, C.Cod_Prod, C.Productor, C.Campo, C.Ubicacion, C.Fecha_solicitud, C.Inicio_cosecha, C.Fecha_real, C.Fecha_analisis, " +
                        "(case when C.Dias <> 0 then 'ENTREGANDO' ELSE 'EN ESPERA' END) AS Estatus ,C.Dias, isnull((case when C.Estatus='R' then 'CON RESIDUOS' else(case when C.Estatus='P' then 'EN PROCESO' else(case when C.Estatus='F' then 'FUERA DE LIMITE' else(case when C.Estatus='L' then 'LIBERADO' end) end) end) end),'') as Analisis " +
                       "from(" +
                       "SELECT A.Nombre as Asesor, M.Cod_Prod, P.Nombre as Productor, C.Descripcion as Campo, isnull(L.Descripcion, '') as Ubicacion, M.Fecha_solicitud, M.Inicio_cosecha, isnull(convert(varchar, E.Fecha, 101), '') as Fecha_real, R1.FechaA as Fecha_analisis, isnull(DATEDIFF(day, R1.FechaA,E.Fecha), 0) as Dias, R.Estatus, C.IdAgen, C.IdAgenC, C.IdAgenI " +
                        "FROM ProdMuestreo M " +
                        "LEFT JOIN ProdAgenteCat A on M.IdAgen = A.IdAgen " +
                        "LEFT JOIN ProdProductoresCat P on M.Cod_Prod = P.Cod_Prod " +
                        "LEFT JOIN ProdCamposCat C on M.Cod_Prod = C.Cod_Prod and M.Cod_Campo = C.Cod_Campo " +
                        "LEFT JOIN CatLocalidades L on C.CodLocalidad = L.CodLocalidad " +
                        "LEFT JOIN (Select distinct Id_Muestreo, max(Fecha) as FechaA FROM ProdAnalisis_Residuo group by Id_Muestreo)R1 ON M.Id=R1.Id_Muestreo " +
                        "LEFT JOIN (Select Estatus, Fecha FROM ProdAnalisis_Residuo)R ON R.Fecha=R1.FechaA " +
                        "LEFT JOIN(Select Cod_Prod, Cod_Campo, min(Fecha) as Fecha, Temporada from UV_ProdRecepcion " +
                        "where Temporada = (Select temporada from CatSemanas where getdate() between Inicio and Fin) " +
                        "group by Cod_Prod, Cod_Campo, Temporada)E on M.Cod_Prod = E.Cod_Prod and M.Cod_Campo = E.Cod_Campo " +
                        "where R.Estatus <> '')C " +
                        "Where C.Estatus='" + estatus + "' and C.IdAgen=" + (short)Session["IdAgen"] + " or C.IdAgenC=" + (short)Session["IdAgen"] + " or C.IdAgenI=" + (short)Session["IdAgen"] + "" +
                        "order by C.Fecha_solicitud desc").ToList();
                }
                else if (cod_prod != "")
                {
                    resultados = bd.Database.SqlQuery<ClassMuestreo>("Select C.Asesor, C.Cod_Prod, C.Productor, C.Campo, C.Ubicacion, C.Fecha_solicitud, C.Inicio_cosecha, C.Fecha_real, C.Fecha_analisis, " +
                        "(case when C.Dias <> 0 then 'ENTREGANDO' ELSE 'EN ESPERA' END) AS Estatus ,C.Dias, isnull((case when C.Estatus='R' then 'CON RESIDUOS' else(case when C.Estatus='P' then 'EN PROCESO' else(case when C.Estatus='F' then 'FUERA DE LIMITE' else(case when C.Estatus='L' then 'LIBERADO' end) end) end) end),'') as Analisis " +
                      "from(" +
                      "SELECT A.Nombre as Asesor, M.Cod_Prod, P.Nombre as Productor, C.Descripcion as Campo, isnull(L.Descripcion, '') as Ubicacion, M.Fecha_solicitud, M.Inicio_cosecha, isnull(convert(varchar, E.Fecha, 101), '') as Fecha_real, R1.FechaA as Fecha_analisis, isnull(DATEDIFF(day, R1.FechaA,E.Fecha), 0) as Dias, R.Estatus, C.IdAgen, C.IdAgenC, C.IdAgenI " +
                      "FROM ProdMuestreo M " +
                        "LEFT JOIN ProdAgenteCat A on M.IdAgen = A.IdAgen " +
                        "LEFT JOIN ProdProductoresCat P on M.Cod_Prod = P.Cod_Prod " +
                        "LEFT JOIN ProdCamposCat C on M.Cod_Prod = C.Cod_Prod and M.Cod_Campo = C.Cod_Campo " +
                        "LEFT JOIN CatLocalidades L on C.CodLocalidad = L.CodLocalidad " +
                        "LEFT JOIN (Select distinct Id_Muestreo, max(Fecha) as FechaA FROM ProdAnalisis_Residuo group by Id_Muestreo)R1 ON M.Id=R1.Id_Muestreo " +
                        "LEFT JOIN (Select Estatus, Fecha FROM ProdAnalisis_Residuo)R ON R.Fecha=R1.FechaA " +
                        "LEFT JOIN(Select Cod_Prod, Cod_Campo, min(Fecha) as Fecha, Temporada from UV_ProdRecepcion " +
                        "where Temporada = (Select temporada from CatSemanas where getdate() between Inicio and Fin) " +
                        "group by Cod_Prod, Cod_Campo, Temporada)E on M.Cod_Prod = E.Cod_Prod and M.Cod_Campo = E.Cod_Campo " +
                        "where R.Estatus <> '')C " +
                        "where C.Cod_Prod='" + cod_prod + "' and C.IdAgen=" + (short)Session["IdAgen"] + " or C.IdAgenC=" + (short)Session["IdAgen"] + " or C.IdAgenI=" + (short)Session["IdAgen"] + "" +
                        "order by C.Fecha_solicitud desc").ToList();
                }
                else
                {
                    resultados = bd.Database.SqlQuery<ClassMuestreo>("Select C.Asesor, C.Cod_Prod, C.Productor, C.Campo, C.Ubicacion, C.Fecha_solicitud, C.Inicio_cosecha, C.Fecha_real, C.Fecha_analisis, " +
                        "(case when C.Dias <> 0 then 'ENTREGANDO' ELSE 'EN ESPERA' END) AS Estatus ,C.Dias, isnull((case when C.Estatus='R' then 'CON RESIDUOS' else(case when C.Estatus='P' then 'EN PROCESO' else(case when C.Estatus='F' then 'FUERA DE LIMITE' else(case when C.Estatus='L' then 'LIBERADO' end) end) end) end),'') as Analisis " +
                       "from(" +
                       "SELECT A.Nombre as Asesor, M.Cod_Prod, P.Nombre as Productor, C.Descripcion as Campo, isnull(L.Descripcion, '') as Ubicacion, M.Fecha_solicitud, M.Inicio_cosecha, isnull(convert(varchar, E.Fecha, 101), '') as Fecha_real, R1.FechaA as Fecha_analisis, isnull(DATEDIFF(day, R1.FechaA,E.Fecha), 0) as Dias, R.Estatus, C.IdAgen, C.IdAgenC, C.IdAgenI " +
                       "FROM ProdMuestreo M " +
                       "LEFT JOIN ProdAgenteCat A on M.IdAgen = A.IdAgen " +
                       "LEFT JOIN ProdProductoresCat P on M.Cod_Prod = P.Cod_Prod " +
                       "LEFT JOIN ProdCamposCat C on M.Cod_Prod = C.Cod_Prod and M.Cod_Campo = C.Cod_Campo " +
                       "LEFT JOIN CatLocalidades L on C.CodLocalidad = L.CodLocalidad " +
                       "LEFT JOIN (Select distinct Id_Muestreo, max(Fecha) as FechaA FROM ProdAnalisis_Residuo group by Id_Muestreo)R1 ON M.Id=R1.Id_Muestreo " +
                       "LEFT JOIN (Select Estatus, Fecha FROM ProdAnalisis_Residuo)R ON R.Fecha=R1.FechaA " +
                       "LEFT JOIN(Select Cod_Prod, Cod_Campo, min(Fecha) as Fecha, Temporada from UV_ProdRecepcion " +
                       "where Temporada = (Select temporada from CatSemanas where getdate() between Inicio and Fin) " +
                       "group by Cod_Prod, Cod_Campo, Temporada)E on M.Cod_Prod = E.Cod_Prod and M.Cod_Campo = E.Cod_Campo " +
                       "where R.Estatus <> '')C " +
                       "where C.IdAgen=" + (short)Session["IdAgen"] + " or C.IdAgenC=" + (short)Session["IdAgen"] + " or C.IdAgenI=" + (short)Session["IdAgen"] + "" +
                       "order by C.Fecha_solicitud desc").ToList();
                }
            }
            return Json(resultados, JsonRequestBehavior.AllowGet);
        }

        public JsonResult EvaluacionListF(string param = "")
        {
            string query = "";
            List<ClassMuestreo> resultados = null;

            if (param != "")
            {
                if (param == "N")
                {
                    query = ">5";
                }
                else if (param == "R")
                {
                    query = "between 3 and 5";
                }
                else if (param == "A")
                {
                    query = "between 2 and 3";
                }
                else if (param == "V")
                {
                    query = "<2";
                }

                if ((short)Session["IdAgen"] == 205 || Session["IdAgen"].ToString() == "153" || Session["IdAgen"].ToString() == "281")
                {
                    resultados = bd.Database.SqlQuery<ClassMuestreo>("Select C.Asesor, C.Cod_Prod, C.Productor, C.Campo, C.Ubicacion, C.Fecha_solicitud, C.Inicio_cosecha, C.Fecha_real, C.Fecha_analisis, " +
                        "(case when C.Dias <> 0 then 'ENTREGANDO' ELSE 'EN ESPERA' END) AS Estatus ,C.Dias, isnull((case when C.Estatus='R' then 'CON RESIDUOS' else(case when C.Estatus='P' then 'EN PROCESO' else(case when C.Estatus='F' then 'FUERA DE LIMITE' else(case when C.Estatus='L' then 'LIBERADO' end) end) end) end),'') as Analisis " +
                    "from(" +
                    "SELECT A.Nombre as Asesor, M.Cod_Prod, P.Nombre as Productor, C.Descripcion as Campo, isnull(L.Descripcion, '') as Ubicacion, M.Fecha_solicitud, M.Inicio_cosecha, isnull(convert(varchar, E.Fecha, 101), '') as Fecha_real, R1.FechaA as Fecha_analisis, isnull(DATEDIFF(day, R1.FechaA,E.Fecha), 0) as Dias, R.Estatus, C.IdAgen, C.IdAgenC, C.IdAgenI " +
                    "FROM ProdMuestreo M " +
                    "LEFT JOIN ProdAgenteCat A on M.IdAgen = A.IdAgen " +
                    "LEFT JOIN ProdProductoresCat P on M.Cod_Prod = P.Cod_Prod " +
                    "LEFT JOIN ProdCamposCat C on M.Cod_Prod = C.Cod_Prod and M.Cod_Campo = C.Cod_Campo " +
                    "LEFT JOIN CatLocalidades L on C.CodLocalidad = L.CodLocalidad " +
                    "LEFT JOIN (Select distinct Id_Muestreo, max(Fecha) as FechaA FROM ProdAnalisis_Residuo group by Id_Muestreo)R1 ON M.Id=R1.Id_Muestreo " +
                    "LEFT JOIN (Select Estatus, Fecha FROM ProdAnalisis_Residuo)R ON R.Fecha=R1.FechaA " +
                    "LEFT JOIN(Select Cod_Prod, Cod_Campo, min(Fecha) as Fecha, Temporada from UV_ProdRecepcion " +
                    "where Temporada = (Select temporada from CatSemanas where getdate() between Inicio and Fin) " +
                    "group by Cod_Prod, Cod_Campo, Temporada)E on M.Cod_Prod = E.Cod_Prod and M.Cod_Campo = E.Cod_Campo " +
                    "where R.Estatus <> '')C " +
                    "Where C.Dias " + query + " " +
                    "order by C.Fecha_solicitud desc").ToList();

                }
                else if ((short)Session["IdAgen"] == 1)
                {
                    resultados = bd.Database.SqlQuery<ClassMuestreo>("Select C.Asesor, C.Cod_Prod, C.Productor, C.Campo, C.Ubicacion, C.Fecha_solicitud, C.Inicio_cosecha, C.Fecha_real, C.Fecha_analisis, " +
                        "(case when C.Dias <> 0 then 'ENTREGANDO' ELSE 'EN ESPERA' END) AS Estatus ,C.Dias, isnull((case when C.Estatus='R' then 'CON RESIDUOS' else(case when C.Estatus='P' then 'EN PROCESO' else(case when C.Estatus='F' then 'FUERA DE LIMITE' else(case when C.Estatus='L' then 'LIBERADO' end) end) end) end),'') as Analisis " +
                     "from(" +
                     "SELECT A.Nombre as Asesor, M.Cod_Prod, P.Nombre as Productor, C.Descripcion as Campo, isnull(L.Descripcion, '') as Ubicacion, M.Fecha_solicitud, M.Inicio_cosecha, isnull(convert(varchar, E.Fecha, 101), '') as Fecha_real, R1.FechaA as Fecha_analisis, isnull(DATEDIFF(day, R1.FechaA,E.Fecha), 0) as Dias, R.Estatus, C.IdAgen, C.IdAgenC, C.IdAgenI " +
                     "FROM ProdMuestreo M " +
                     "LEFT JOIN ProdAgenteCat A on M.IdAgen = A.IdAgen " +
                     "LEFT JOIN ProdProductoresCat P on M.Cod_Prod = P.Cod_Prod " +
                     "LEFT JOIN ProdCamposCat C on M.Cod_Prod = C.Cod_Prod and M.Cod_Campo = C.Cod_Campo " +
                     "LEFT JOIN CatLocalidades L on C.CodLocalidad = L.CodLocalidad " +
                     "LEFT JOIN (Select distinct Id_Muestreo, max(Fecha) as FechaA FROM ProdAnalisis_Residuo group by Id_Muestreo)R1 ON M.Id=R1.Id_Muestreo " +
                     "LEFT JOIN (Select Estatus, Fecha FROM ProdAnalisis_Residuo)R ON R.Fecha=R1.FechaA " +
                     "LEFT JOIN(Select Cod_Prod, Cod_Campo, min(Fecha) as Fecha, Temporada from UV_ProdRecepcion " +
                     "where Temporada = (Select temporada from CatSemanas where getdate() between Inicio and Fin) " +
                     "group by Cod_Prod, Cod_Campo, Temporada)E on M.Cod_Prod = E.Cod_Prod and M.Cod_Campo = E.Cod_Campo " +
                     "where R.Estatus <> '' and C.IdZona in (1,2,3,4,7,8,10,11,12,13,14,15))C " +
                     "Where C.Dias " + query + " " +
                     "order by C.Fecha_solicitud desc").ToList();
                }
                else if ((short)Session["IdAgen"] == 5)
                {
                    resultados = bd.Database.SqlQuery<ClassMuestreo>("Select C.Asesor, C.Cod_Prod, C.Productor, C.Campo, C.Ubicacion, C.Fecha_solicitud, C.Inicio_cosecha, C.Fecha_real, C.Fecha_analisis, " +
                        "(case when C.Dias <> 0 then 'ENTREGANDO' ELSE 'EN ESPERA' END) AS Estatus ,C.Dias, isnull((case when C.Estatus='R' then 'CON RESIDUOS' else(case when C.Estatus='P' then 'EN PROCESO' else(case when C.Estatus='F' then 'FUERA DE LIMITE' else(case when C.Estatus='L' then 'LIBERADO' end) end) end) end),'') as Analisis " +
                     "from(" +
                     "SELECT A.Nombre as Asesor, M.Cod_Prod, P.Nombre as Productor, C.Descripcion as Campo, isnull(L.Descripcion, '') as Ubicacion, M.Fecha_solicitud, M.Inicio_cosecha, isnull(convert(varchar, E.Fecha, 101), '') as Fecha_real, R1.FechaA as Fecha_analisis, isnull(DATEDIFF(day, R1.FechaA,E.Fecha), 0) as Dias, R.Estatus, C.IdAgen, C.IdAgenC, C.IdAgenI " +
                     "FROM ProdMuestreo M " +
                     "LEFT JOIN ProdAgenteCat A on M.IdAgen = A.IdAgen " +
                     "LEFT JOIN ProdProductoresCat P on M.Cod_Prod = P.Cod_Prod " +
                     "LEFT JOIN ProdCamposCat C on M.Cod_Prod = C.Cod_Prod and M.Cod_Campo = C.Cod_Campo " +
                     "LEFT JOIN CatLocalidades L on C.CodLocalidad = L.CodLocalidad " +
                     "LEFT JOIN (Select distinct Id_Muestreo, max(Fecha) as FechaA FROM ProdAnalisis_Residuo group by Id_Muestreo)R1 ON M.Id=R1.Id_Muestreo " +
                     "LEFT JOIN (Select Estatus, Fecha FROM ProdAnalisis_Residuo)R ON R.Fecha=R1.FechaA " +
                     "LEFT JOIN(Select Cod_Prod, Cod_Campo, min(Fecha) as Fecha, Temporada from UV_ProdRecepcion " +
                     "where Temporada = (Select temporada from CatSemanas where getdate() between Inicio and Fin) " +
                     "group by Cod_Prod, Cod_Campo, Temporada)E on M.Cod_Prod = E.Cod_Prod and M.Cod_Campo = E.Cod_Campo " +
                     "where R.Estatus <> '' and C.IdZona in (5,6,9))C " +
                     "Where C.Dias " + query + " " +
                     "order by C.Fecha_solicitud desc").ToList();
                }
                else
                {
                    resultados = bd.Database.SqlQuery<ClassMuestreo>("Select C.Asesor, C.Cod_Prod, C.Productor, C.Campo, C.Ubicacion, C.Fecha_solicitud, C.Inicio_cosecha, C.Fecha_real, C.Fecha_analisis, " +
                        "(case when C.Dias <> 0 then 'ENTREGANDO' ELSE 'EN ESPERA' END) AS Estatus ,C.Dias, isnull((case when C.Estatus='R' then 'CON RESIDUOS' else(case when C.Estatus='P' then 'EN PROCESO' else(case when C.Estatus='F' then 'FUERA DE LIMITE' else(case when C.Estatus='L' then 'LIBERADO' end) end) end) end),'') as Analisis " +
                      "from(" +
                      "SELECT A.Nombre as Asesor, M.Cod_Prod, P.Nombre as Productor, C.Descripcion as Campo, isnull(L.Descripcion, '') as Ubicacion, M.Fecha_solicitud, M.Inicio_cosecha, isnull(convert(varchar, E.Fecha, 101), '') as Fecha_real, R1.FechaA as Fecha_analisis, isnull(DATEDIFF(day, R1.FechaA,E.Fecha), 0) as Dias, R.Estatus, C.IdAgen, C.IdAgenC, C.IdAgenI " +
                      "FROM ProdMuestreo M " +
                      "LEFT JOIN ProdAgenteCat A on M.IdAgen = A.IdAgen " +
                      "LEFT JOIN ProdProductoresCat P on M.Cod_Prod = P.Cod_Prod " +
                      "LEFT JOIN ProdCamposCat C on M.Cod_Prod = C.Cod_Prod and M.Cod_Campo = C.Cod_Campo " +
                      "LEFT JOIN CatLocalidades L on C.CodLocalidad = L.CodLocalidad " +
                      "LEFT JOIN (Select distinct Id_Muestreo, max(Fecha) as FechaA FROM ProdAnalisis_Residuo group by Id_Muestreo)R1 ON M.Id=R1.Id_Muestreo " +
                      "LEFT JOIN (Select Estatus, Fecha FROM ProdAnalisis_Residuo)R ON R.Fecha=R1.FechaA " +
                      "LEFT JOIN(Select Cod_Prod, Cod_Campo, min(Fecha) as Fecha, Temporada from UV_ProdRecepcion " +
                      "where Temporada = (Select temporada from CatSemanas where getdate() between Inicio and Fin) " +
                      "group by Cod_Prod, Cod_Campo, Temporada)E on M.Cod_Prod = E.Cod_Prod and M.Cod_Campo = E.Cod_Campo " +
                      "where R.Estatus <> '')C " +
                      "Where C.Dias " + query + " and (C.IdAgen=" + (short)Session["IdAgen"] + " or C.IdAgenC=" + (short)Session["IdAgen"] + " or C.IdAgenI=" + (short)Session["IdAgen"] + ") " +
                      "order by C.Fecha_solicitud desc").ToList();
                }
            }
            return Json(resultados, JsonRequestBehavior.AllowGet);
        }

        public ActionResult Liberar_Estatus()
        {
            if (Session["Nombre"] != null)
            {
                ViewData["Nombre"] = Session["Nombre"].ToString();

                var item = from r in bd.ProdAnalisis_Residuo
                           join s in bd.ProdMuestreoSector on r.IdSector equals s.id into MuestreoSc
                           from ms in MuestreoSc.DefaultIfEmpty()
                           join c in bd.ProdZonasRastreoCat on r.CodZona equals c.Codigo into Zona
                           from z in Zona.DefaultIfEmpty()
                           where r.Estatus == "F" && (r.LiberacionUSA <= DateTime.Now || r.LiberacionEU <= DateTime.Now)
                           select new ClassMuestreo
                           {
                               IdAnalisis_Residuo = r.Id,
                               Cod_Prod = r.Cod_Prod,
                               Cod_Campo = r.Cod_Campo,
                               Sector = (short)ms.Sector,
                               DescZona = z.DescZona,
                               Fecha_envio = (DateTime)r.Fecha_envio,
                               Fecha_entrega = (DateTime)r.Fecha_entrega,
                               Analisis = r.Estatus,
                               Num_analisis = r.Num_analisis,
                               Laboratorio = r.Laboratorio,
                               Comentarios = r.Comentarios,
                               LiberacionUSA = (DateTime)r.LiberacionUSA,
                               LiberacionEU = (DateTime)r.LiberacionEU
                           };

                return View(item.ToList());
            }
            else
            {
                return RedirectToAction("Index", "Login");
            }
        }

        public JsonResult Liberar_EstatusList(string param = "")
        {
            List<ClassMuestreo> resultados = bd.Database.SqlQuery<ClassMuestreo>("select * from(select A.Id as IdAnalisis_Residuo, A.Fecha, isnull(A.Cod_Prod,'') as Cod_Prod, isnull(A.Cod_Campo,'') as Cod_Campo, isnull(S.Sector,'') as Sector, " +
                "isnull(Z.DescZona, '') as DescZona, A.Fecha_envio, A.Fecha_entrega, (case when A.Estatus = 'F' THEN 'FUERA DE LIMITE' END) as Estatus, A.Num_analisis, A.Laboratorio, isnull(A.Comentarios, '') as Comentarios, " +
                "isnull(convert(varchar, A.LiberacionUSA, 105), '') as LiberacionUSA, isnull(convert(varchar, A.LiberacionEU, 105), '') AS LiberacionEU, " +
                "DATEDIFF(day, LiberacionUSA, getdate()) as diasUSA, DATEDIFF(day, LiberacionEU, getdate()) as diasEU " +
                "from ProdAnalisis_Residuo A " +
                "left join ProdMuestreoSector S on A.IdSector = S.Id " +
                "left join ProdZonasRastreoCat Z on A.CodZona = Z.Codigo " +
                "where A.Estatus = 'F')A WHERE A.diasUSA >= 0 OR A.diasEU >= 0 " +
                "order by A.Fecha").ToList();
            return Json(resultados, JsonRequestBehavior.AllowGet);
        }

        [HttpPost]
        public ActionResult Liberar_Analisis(int id_analisis)
        {
            if (Session["Nombre"] != null)
            {
                ViewData["Nombre"] = Session["Nombre"].ToString();

                if (id_analisis > 0)
                {
                    var item = bd.ProdAnalisis_Residuo.Where(x => x.Id == id_analisis).First();
                    item.Estatus = "L";
                    bd.SaveChanges();

                    try
                    {
                        string estatus = "", correo_p, correo_c, correo_i;

                        var campo = bd.ProdCamposCat.FirstOrDefault(m => m.Cod_Prod == item.Cod_Prod && m.Cod_Campo == item.Cod_Campo);
                        var email_p = bd.SIPGUsuarios.FirstOrDefault(m => m.IdAgen == campo.IdAgen && m.Tipo == "P");
                        var email_c = bd.SIPGUsuarios.FirstOrDefault(m => m.IdAgen == campo.IdAgenC && m.Tipo == "C");
                        var email_i = bd.SIPGUsuarios.FirstOrDefault(m => m.IdAgen == campo.IdAgenI && m.Tipo == "I");
                        var prod = bd.ProdProductoresCat.FirstOrDefault(x => x.Cod_Prod == item.Cod_Prod);
                        correo_p = email_p.correo;

                        if (email_c == null)
                        {
                            //var item = bd.ProdCamposCat.Where(x => x.Cod_Prod == model.Cod_Prod && x.Cod_Campo == model.Cod_Campo).First();
                            campo.IdAgenC = 167;
                            bd.SaveChanges();
                            correo_c = "mayra.ramirez@giddingsfruit.mx";
                        }
                        else
                        {
                            correo_c = email_c.correo;
                        }

                        if (email_i == null)
                        {
                            //var item = bd.ProdCamposCat.Where(x => x.Cod_Prod == model.Cod_Prod && x.Cod_Campo == model.Cod_Campo).First();
                            campo.IdAgenI = 205;
                            bd.SaveChanges();
                            correo_i = "jesus.palafox@giddingsfruit.mx";
                        }
                        else
                        {
                            correo_i = email_i.correo;
                        }

                        MailMessage correo = new MailMessage();
                        correo.From = new MailAddress("indicadores.giddingsfruit@gmail.com", "Indicadores GiddingsFruit");
                        correo.To.Add(Session["Correo"].ToString());
                        correo.CC.Add(correo_p);
                        correo.CC.Add(correo_c);
                        correo.CC.Add(correo_i);
                        
                        correo.Subject = "Analisis liberado: " + campo.Cod_Prod + " - " + estatus;
                        correo.Body += "El codigo: " + campo.Cod_Prod + " - " + prod.Nombre + " <br/>";
                        correo.Body += " <br/>";
                        correo.Body += "Campo: " + campo.Cod_Campo + " - " + campo.Descripcion + " <br/>";
                        correo.Body += " <br/>";
                        correo.Body += "Ha cambiado su estatus de: Fuera de Limite a Liberado";
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
                    catch (Exception e)
                    {
                        e.ToString();
                    }
                }
            }
            return RedirectToAction("Liberar_Estatus", "Muestreo");
        }

        public JsonResult SectoresList(int IdMuestreo)
        {
            ClassMuestreo resultados = bd.Database.SqlQuery<ClassMuestreo>("SELECT STUFF((SELECT ',' + cast(Sector as varchar) as SectorList FROM ProdMuestreoSector " +
                "where IdMuestreo = " + IdMuestreo + " FOR XML PATH(''), TYPE).value('.', 'NVARCHAR(MAX)'), 1, 1, '') as SectorList").First();
            return Json(resultados, JsonRequestBehavior.AllowGet);
        }

        public ActionResult Muestreos_Liberados()
        {
            if (Session["Nombre"] != null)
            {
                ViewData["Nombre"] = Session["Nombre"].ToString();

                short agenteSesion = (short)Session["IdAgen"];
                IQueryable<ClassMuestreo> item = null;

                if (Session["IdAgen"].ToString() == "205" || Session["IdAgen"].ToString() == "153" || Session["IdAgen"].ToString() == "281" || Session["IdAgen"].ToString() == "167" || Session["IdAgen"].ToString() == "182" || Session["IdAgen"].ToString() == "1")
                {
                   item= (from m in (from m in bd.ProdAnalisis_Residuo
                                group m by new
                                {
                                    Cod_Empresa = m.Cod_Empresa,
                                    Cod_Prod = m.Cod_Prod,
                                    Cod_Campo = m.Cod_Campo,
                                    IdSector = m.IdSector
                                } into x
                                select new
                                {
                                    Cod_Empresa = x.Key.Cod_Empresa,
                                    Cod_Prod = x.Key.Cod_Prod,
                                    Cod_Campo = x.Key.Cod_Campo,
                                    IdSector = x.Key.IdSector,
                                    Fecha = x.Max(m => m.Fecha)
                                })

                     join an in bd.ProdAnalisis_Residuo on new { m.Cod_Empresa, m.Cod_Prod, m.Cod_Campo, m.IdSector, m.Fecha } equals new { an.Cod_Empresa, an.Cod_Prod, an.Cod_Campo, an.IdSector, an.Fecha } into AnalisisR
                     from an in AnalisisR.DefaultIfEmpty()

                     join c in bd.ProdCamposCat on new { m.Cod_Empresa, m.Cod_Prod, m.Cod_Campo } equals new { c.Cod_Empresa, c.Cod_Prod, c.Cod_Campo } into MuestreoCam
                     from mcam in MuestreoCam.DefaultIfEmpty()

                     join p in bd.ProdProductoresCat on mcam.Cod_Prod equals p.Cod_Prod into MuestreoProd
                     from prod in MuestreoProd.DefaultIfEmpty()

                     join s in bd.ProdMuestreoSector on m.IdSector equals s.id into MuestreoSc
                     from ms in MuestreoSc.DefaultIfEmpty()

                     join r in bd.ProdMuestreo on an.Id_Muestreo equals r.Id into MuestreoAn
                     from man in MuestreoAn.DefaultIfEmpty()

                     join a in bd.ProdAgenteCat on mcam.IdAgen equals a.IdAgen into MuestreoAgentes
                     from ageP in MuestreoAgentes.DefaultIfEmpty()

                     join cf in bd.ProdCalidadMuestreo on man.Id equals cf.Id_Muestreo into MuestreoCa
                     from mc in MuestreoCa.DefaultIfEmpty()

                     join a in bd.ProdAgenteCat on mc.IdAgen equals a.IdAgen into MuestreoAgenC
                     from ageC in MuestreoAgenC.DefaultIfEmpty()

                     join a in bd.ProdAgenteCat on mcam.IdAgenC equals a.IdAgen into MuestreoAgenSC
                     from ageCS in MuestreoAgenSC.DefaultIfEmpty()

                     join a in bd.ProdAgenteCat on mcam.IdAgenI equals a.IdAgen into MuestreoAgenI
                     from ageI in MuestreoAgenI.DefaultIfEmpty()

                     join l in bd.CatLocalidades on mcam.CodLocalidad equals l.CodLocalidad into MuestreoLoc
                     from loc in MuestreoLoc.DefaultIfEmpty()

                     where an.Estatus == "L" //&& (mc.Estatus == "1" || mc.Estatus == "2") && man.Tarjeta == "S"
                     
                     group m by new
                     {
                         IdAnalisis_Residuo = an.Id,
                         IdMuestreo = man.Id,
                         Asesor = ageP.Nombre,
                         Cod_Prod = m.Cod_Prod,
                         Productor = prod.Nombre,
                         Cod_Campo = m.Cod_Campo,
                         Campo = mcam.Descripcion,
                         Sector = (short)ms.Sector,
                         Ha = mcam.Hectareas,
                         Compras_oportunidad = mcam.Compras_Oportunidad,
                         Fecha_solicitud = (DateTime)man.Fecha_solicitud,
                         Inicio_cosecha = (DateTime)man.Inicio_cosecha,
                         Ubicacion = loc.Descripcion,
                         Telefono = man.Telefono,
                         Liberacion = man.Liberacion,
                         Fecha_ejecucion = (DateTime)man.Fecha_ejecucion,
                         Analisis = an.Estatus,
                         Calidad_fruta = mc.Estatus,
                         IdAgenC = (short)mcam.IdAgenC,
                         AsesorC = ageC.Nombre,
                         AsesorCS = ageCS.Nombre,
                         AsesorI = ageI.Nombre,
                         Tarjeta = man.Tarjeta,
                         IdRegion = ageP.IdRegion,
                         Fecha_analisis = m.Fecha
                     } into x
                     select new ClassMuestreo()
                     {
                         IdAnalisis_Residuo = x.Key.IdAnalisis_Residuo,
                         IdMuestreo = x.Key.IdMuestreo,
                         Asesor = x.Key.Asesor,
                         Cod_Prod = x.Key.Cod_Prod,
                         Productor = x.Key.Productor,
                         Cod_Campo = x.Key.Cod_Campo,
                         Campo = x.Key.Campo,
                         Sector = (short)x.Key.Sector,
                         Ha = x.Key.Ha,
                         Compras_oportunidad = x.Key.Compras_oportunidad,
                         Fecha_solicitud = (DateTime)x.Key.Fecha_solicitud,
                         Inicio_cosecha = (DateTime)x.Key.Inicio_cosecha,
                         Ubicacion = x.Key.Ubicacion,
                         Telefono = x.Key.Telefono,
                         Liberacion = x.Key.Liberacion,
                         Fecha_ejecucion = (DateTime)x.Key.Fecha_ejecucion,
                         Analisis = x.Key.Analisis,
                         Calidad_fruta = x.Key.Calidad_fruta,
                         IdAgenC = (short)x.Key.IdAgenC,
                         AsesorC = x.Key.AsesorC,
                         AsesorCS = x.Key.AsesorCS,
                         AsesorI = x.Key.AsesorI,
                         Tarjeta = x.Key.Tarjeta,
                         IdRegion = x.Key.IdRegion,
                         Fecha_analisis = x.Key.Fecha_analisis
                     }).Distinct().OrderByDescending(m=>m.Fecha_analisis);
                }

                else
                {
                    item = (from m in (from m in bd.ProdAnalisis_Residuo
                                       group m by new
                                       {
                                           Cod_Empresa = m.Cod_Empresa,
                                           Cod_Prod = m.Cod_Prod,
                                           Cod_Campo = m.Cod_Campo,
                                           IdSector = m.IdSector
                                       } into x
                                       select new
                                       {
                                           Cod_Empresa = x.Key.Cod_Empresa,
                                           Cod_Prod = x.Key.Cod_Prod,
                                           Cod_Campo = x.Key.Cod_Campo,
                                           IdSector = x.Key.IdSector,
                                           Fecha = x.Max(m => m.Fecha)
                                       })

                            join an in bd.ProdAnalisis_Residuo on new { m.Cod_Empresa, m.Cod_Prod, m.Cod_Campo, m.IdSector, m.Fecha } equals new { an.Cod_Empresa, an.Cod_Prod, an.Cod_Campo, an.IdSector, an.Fecha } into AnalisisR
                            from an in AnalisisR.DefaultIfEmpty()

                            join c in bd.ProdCamposCat on new { m.Cod_Empresa, m.Cod_Prod, m.Cod_Campo } equals new { c.Cod_Empresa, c.Cod_Prod, c.Cod_Campo } into MuestreoCam
                            from mcam in MuestreoCam.DefaultIfEmpty()

                            join p in bd.ProdProductoresCat on mcam.Cod_Prod equals p.Cod_Prod into MuestreoProd
                            from prod in MuestreoProd.DefaultIfEmpty()

                            join s in bd.ProdMuestreoSector on m.IdSector equals s.id into MuestreoSc
                            from ms in MuestreoSc.DefaultIfEmpty()

                            join r in bd.ProdMuestreo on an.Id_Muestreo equals r.Id into MuestreoAn
                            from man in MuestreoAn.DefaultIfEmpty()

                            join a in bd.ProdAgenteCat on mcam.IdAgen equals a.IdAgen into MuestreoAgentes
                            from ageP in MuestreoAgentes.DefaultIfEmpty()

                            join cf in bd.ProdCalidadMuestreo on man.Id equals cf.Id_Muestreo into MuestreoCa
                            from mc in MuestreoCa.DefaultIfEmpty()

                            join a in bd.ProdAgenteCat on mc.IdAgen equals a.IdAgen into MuestreoAgenC
                            from ageC in MuestreoAgenC.DefaultIfEmpty()

                            join a in bd.ProdAgenteCat on mcam.IdAgenC equals a.IdAgen into MuestreoAgenSC
                            from ageCS in MuestreoAgenSC.DefaultIfEmpty()

                            join a in bd.ProdAgenteCat on mcam.IdAgenI equals a.IdAgen into MuestreoAgenI
                            from ageI in MuestreoAgenI.DefaultIfEmpty()

                            join l in bd.CatLocalidades on mcam.CodLocalidad equals l.CodLocalidad into MuestreoLoc
                            from loc in MuestreoLoc.DefaultIfEmpty()

                            where an.Estatus == "L" //&& (mc.Estatus == "1" || mc.Estatus == "2") && man.Tarjeta == "S"
                            && (mcam.IdAgen == agenteSesion || mcam.IdAgenI == agenteSesion || mcam.IdAgenC == agenteSesion)
                            group m by new
                            {
                                IdAnalisis_Residuo = an.Id,
                                IdMuestreo = man.Id,
                                Asesor = ageP.Nombre,
                                Cod_Prod = m.Cod_Prod,
                                Productor = prod.Nombre,
                                Cod_Campo = m.Cod_Campo,
                                Campo = mcam.Descripcion,
                                Sector = (short)ms.Sector,
                                Ha = mcam.Hectareas,
                                Compras_oportunidad = mcam.Compras_Oportunidad,
                                Fecha_solicitud = (DateTime)man.Fecha_solicitud,
                                Inicio_cosecha = (DateTime)man.Inicio_cosecha,
                                Ubicacion = loc.Descripcion,
                                Telefono = man.Telefono,
                                Liberacion = man.Liberacion,
                                Fecha_ejecucion = (DateTime)man.Fecha_ejecucion,
                                Analisis = an.Estatus,
                                Calidad_fruta = mc.Estatus,
                                IdAgenC = (short)mcam.IdAgenC,
                                AsesorC = ageC.Nombre,
                                AsesorCS = ageCS.Nombre,
                                AsesorI = ageI.Nombre,
                                Tarjeta = man.Tarjeta,
                                IdRegion = ageP.IdRegion,
                                Fecha_analisis = m.Fecha
                            } into x
                            select new ClassMuestreo()
                            {
                                IdAnalisis_Residuo = x.Key.IdAnalisis_Residuo,
                                IdMuestreo = x.Key.IdMuestreo,
                                Asesor = x.Key.Asesor,
                                Cod_Prod = x.Key.Cod_Prod,
                                Productor = x.Key.Productor,
                                Cod_Campo = x.Key.Cod_Campo,
                                Campo = x.Key.Campo,
                                Sector = (short)x.Key.Sector,
                                Ha = x.Key.Ha,
                                Compras_oportunidad = x.Key.Compras_oportunidad,
                                Fecha_solicitud = (DateTime)x.Key.Fecha_solicitud,
                                Inicio_cosecha = (DateTime)x.Key.Inicio_cosecha,
                                Ubicacion = x.Key.Ubicacion,
                                Telefono = x.Key.Telefono,
                                Liberacion = x.Key.Liberacion,
                                Fecha_ejecucion = (DateTime)x.Key.Fecha_ejecucion,
                                Analisis = x.Key.Analisis,
                                Calidad_fruta = x.Key.Calidad_fruta,
                                IdAgenC = (short)x.Key.IdAgenC,
                                AsesorC = x.Key.AsesorC,
                                AsesorCS = x.Key.AsesorCS,
                                AsesorI = x.Key.AsesorI,
                                Tarjeta = x.Key.Tarjeta,
                                IdRegion = x.Key.IdRegion,
                                Fecha_analisis = x.Key.Fecha_analisis
                            }).Distinct().OrderByDescending(m => m.Fecha_analisis);
                }
                return View(item.ToList());
            }
            else
            {
                return RedirectToAction("Index", "Login");
            }
        }
    }
}