using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Sistema_Indicadores.Controllers
{
    public class EncuestaController : Controller
    {
        // GET: Encuesta
        public ActionResult Index()
        {
            if (Session["Nombre"] != null)
            {
                ViewData["Nombre"] = Session["Nombre"].ToString(); 
                return View();
            }
            else
            {
                return RedirectToAction("Index", "Login");
            }            
        }

        // GET: Encuesta/Details/5
        public ActionResult Details(int id)
        {
            return View();
        }

        // GET: Encuesta/Create
        public ActionResult Create()
        {
            return View();
        }

        // POST: Encuesta/Create
        [HttpPost]
        public ActionResult Create(FormCollection collection)
        {
            try
            {
                // TODO: Add insert logic here

                return RedirectToAction("Index");
            }
            catch
            {
                return View();
            }
        }

        // GET: Encuesta/Edit/5
        public ActionResult Edit(int id)
        {
            return View();
        }

        // POST: Encuesta/Edit/5
        [HttpPost]
        public ActionResult Edit(int id, FormCollection collection)
        {
            try
            {
                // TODO: Add update logic here

                return RedirectToAction("Index");
            }
            catch
            {
                return View();
            }
        }

        // GET: Encuesta/Delete/5
        public ActionResult Delete(int id)
        {
            return View();
        }

        // POST: Encuesta/Delete/5
        [HttpPost]
        public ActionResult Delete(int id, FormCollection collection)
        {
            try
            {
                // TODO: Add delete logic here

                return RedirectToAction("Index");
            }
            catch
            {
                return View();
            }
        }
    }
}
