using Sistema_Indicadores.Clases;
using Sistema_Indicadores.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.Mvc;

namespace Sistema_Indicadores.Controllers
{
    public class HomeController : Controller
    {
        SeasonSun1213Entities15 bd = new SeasonSun1213Entities15();    
        public ActionResult Index()
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

        public JsonResult Liberar_FLimite_List()
        {
            ClassMuestreo resultados = bd.Database.SqlQuery<ClassMuestreo>("select count(*) proximos_liberar " +
                "from(select *, DATEDIFF(day, LiberacionUSA, getdate()) as diasUSA, DATEDIFF(day, LiberacionEU, getdate()) as diasEU " +
                "from prodanalisis_residuo " +
                "where Estatus = 'F')A where A.diasUSA >= 0 OR A.diasEU >= 0").SingleOrDefault();
            return Json(resultados, JsonRequestBehavior.AllowGet);
        }
    }
}