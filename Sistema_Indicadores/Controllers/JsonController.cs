using Sistema_Indicadores.Clases;
using Sistema_Indicadores.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Sistema_Indicadores.Controllers
{
    public class JsonController : Controller
    {
        SeasonSun1213Entities15 bd = new SeasonSun1213Entities15();
        SeasonSun1213Entities16 bdI = new SeasonSun1213Entities16();
        public JsonResult GetProductor(string Cod_Prod)
        {
            bd.Configuration.ProxyCreationEnabled = false;
            ProdProductoresCat nom_prod = bd.ProdProductoresCat.Where(x => x.Cod_Prod == Cod_Prod).FirstOrDefault();
            return Json(nom_prod, JsonRequestBehavior.AllowGet);
        }

        public JsonResult DescCampo(string Cod_Prod, short Cod_Campo)
        {
            bd.Configuration.ProxyCreationEnabled = false;
            ProdCamposCat descipcion = bd.ProdCamposCat.Where(x => x.Cod_Prod == Cod_Prod && x.Cod_Campo == Cod_Campo).FirstOrDefault();
            return Json(descipcion, JsonRequestBehavior.AllowGet);
        }
        //Lista de productoresc
        public JsonResult GetProductoressList(string Cod_Prod, int IdAgen)
        {
            IdAgen = Convert.ToInt32(Session["IdAgen"]);
            bd.Configuration.ProxyCreationEnabled = false;
            List<ProdCamposCat> ListProductores = bd.ProdCamposCat.Where(x => x.Cod_Prod == Cod_Prod && x.IdAgen == IdAgen).ToList();
            return Json(ListProductores, JsonRequestBehavior.AllowGet);
        }

        //Lista de campos
        public JsonResult GetCamposList(string Cod_Prod)
        {
            bd.Configuration.ProxyCreationEnabled = false;
            List<ProdCamposCat> ListCampos = bd.ProdCamposCat.Where(x => x.Cod_Prod == Cod_Prod).ToList();            
            return Json(ListCampos, JsonRequestBehavior.AllowGet);
        }

        //Lista de asesores
        public JsonResult GetAsesorPList(string Cod_Prod)
        {
            bd.Configuration.ProxyCreationEnabled = false;
            var ListAgentesP = (from c in bd.ProdCamposCat join a in bd.ProdAgenteCat on c.IdAgen equals a.IdAgen where c.Cod_Prod==Cod_Prod select new { IdAgen = a.IdAgen, Nombre = a.Nombre }).Distinct().ToList();
            return Json(ListAgentesP, JsonRequestBehavior.AllowGet);
        }

        //Lista de sectores
        //public JsonResult GetSectoresList(string Cod_Prod, short Cod_Campo)
        //{
        //    bd.Configuration.ProxyCreationEnabled = false;
        //    List<ProdProyeccion> ListSectores = bdI.ProdProyeccion.Where(x => x.Cod_Prod == Cod_Prod && x.Cod_Campo == Cod_Campo).ToList();
        //    return Json(ListSectores, JsonRequestBehavior.AllowGet);
        //}
        public JsonResult GetDataProyeccionList(int Id_Proyeccion)
        {
            //short idagen = (short)Session["IdAgen"];
            //bd.Configuration.ProxyCreationEnabled = false;
            //var temporada = bdI.CatSemanas.Where(x => DateTime.Now >= x.Inicio && DateTime.Now <= x.Fin).FirstOrDefault();             
            //var max_fecha = bdI.ProdProyeccion.Max(x => x.Fecha);
            //List<ProdProyeccion> ListSectores = bdI.ProdProyeccion.Where(x=> x.IdAgen == idagen && x.Cod_Prod == Cod_Prod && x.Cod_Campo == Cod_Campo && x.Temporada==temporada.Temporada && x.Fecha == max_fecha).ToList();
            List<ProdProyeccion> ListSectores = bdI.ProdProyeccion.Where(x => x.Id == Id_Proyeccion).ToList();
            return Json(ListSectores, JsonRequestBehavior.AllowGet);
        }

        public JsonResult GetDataSectorFechas(int Id_Proyeccion)
        {
            bd.Configuration.ProxyCreationEnabled = false;
            ClassCurva ListFechasxSector = bd.Database.SqlQuery<ClassCurva>("select isnull(Fecha_defoliacion,'') as Fecha_defoliacion, isnull(Fecha_corte1,'') as Fecha_corte1, isnull(Fecha_redefoliacion,'') as Fecha_redefoliacion, isnull(Fecha_corte2,'') as Fecha_corte2 " +
                "from ProdProyeccion where Id=" + Id_Proyeccion + "").FirstOrDefault();
            return Json(ListFechasxSector, JsonRequestBehavior.AllowGet);
        }

        //Ubicacion del campo
        public JsonResult GetUbicacion_campo(string Cod_Prod, short Cod_Campo)
        {
            bd.Configuration.ProxyCreationEnabled = false;
            ProdCamposCat ubicacion = bd.ProdCamposCat.Where(x => x.Cod_Prod == Cod_Prod && x.Cod_Campo == Cod_Campo).FirstOrDefault();
            return Json(ubicacion, JsonRequestBehavior.AllowGet);
        }

        public JsonResult GetLocalidad_campo(string Cod_Prod, int Cod_Campo)
        {
            ClassLocalidades localidad = bd.Database.SqlQuery<ClassLocalidades>("SELECT L.Descripcion From ProdCamposCat C left join CatLocalidades L on C.CodLocalidad=L.CodLocalidad where C.Cod_Prod='" + Cod_Prod + "' and C.Cod_Campo=" + Cod_Campo + "").First();
            return Json(localidad, JsonRequestBehavior.AllowGet);
        }

        public JsonResult GetLocalidad(string Cod_Localidad)
        {
            ClassLocalidades localidad = bd.Database.SqlQuery<ClassLocalidades>("SELECT Descripcion From CatLocalidades where CodLocalidad='" + Cod_Localidad + "'").First();
            return Json(localidad, JsonRequestBehavior.AllowGet);
        }

        public JsonResult GetTipoProducto(string Cod_Prod, int Cod_Campo)
        {
            ClassTipoProducto Tipo_Producto = bd.Database.SqlQuery<ClassTipoProducto>("select T.Descripcion as Tipo, P.Descripcion AS Producto from ProdCamposCat C Left join CatTiposProd T ON C.Tipo=T.Tipo left join CatProductos P on C.Tipo=P.Tipo and C.Producto=P.Producto where C.cod_prod='" + Cod_Prod + "' and C.Cod_Campo=" + Cod_Campo + "").First();
            return Json(Tipo_Producto, JsonRequestBehavior.AllowGet);
        }

        //Lista de variedades
        public JsonResult GetVariedadesList(string Cultivo)
        {
            int tipo = 0;
            if (Cultivo == "ZARZAMORA") { tipo = 1; }
            if (Cultivo == "FRAMBUESA") { tipo = 2; }
            if (Cultivo == "ARANDANO") { tipo = 3; }
            if (Cultivo == "FRESA") { tipo = 4; }
            bd.Configuration.ProxyCreationEnabled = false;
            List<CatProductos> ListCampos = bd.CatProductos.Where(x => x.Tipo == tipo).ToList();
            return Json(ListCampos, JsonRequestBehavior.AllowGet);
        }

        //nombre campo
        public JsonResult GetCampo(int Cod_Campo = 0)
        {
            bd.Configuration.ProxyCreationEnabled = false;
            ProdCamposCat nom_campo = bd.ProdCamposCat.Where(x => x.Cod_Campo == Cod_Campo).FirstOrDefault();
            return Json(nom_campo, JsonRequestBehavior.AllowGet);
        }

        public JsonResult GetInocuidadList(string Cultivo)
        {
            bd.Configuration.ProxyCreationEnabled = false;
            List<ProdAgenteCat> ListAgentes = bd.ProdAgenteCat.Where(x => x.Depto == "I").ToList();
            return Json(ListAgentes, JsonRequestBehavior.AllowGet);
        }

        //Buscar por ingeniero autocomplete con ajax  
        [HttpPost]
        public JsonResult GetIngenieroI(string Nombre)
        {
            var agenteC = (from c in bd.ProdAgenteCat
                           where c.Nombre.StartsWith(Nombre) && c.Depto == "C"
                           select new
                           {
                               label = c.Nombre,
                               val = c.IdAgen
                           }).ToList();

            return Json(agenteC);
        }

        public JsonResult IncidenciasList(int Id_Muestreo)
        {
            List<ClassMuestreo> resultados = bd.Database.SqlQuery<ClassMuestreo>("SELECT isnull(Incidencia,'') as Incidencia, isnull(Propuesta,'') as Propuesta from ProdCalidadMuestreo " +
                "where Id_Muestreo=" + Id_Muestreo + " order by Fecha desc").ToList();
            return Json(resultados, JsonRequestBehavior.AllowGet);
        }

        //Cajas entregadas  
        public JsonResult GetCajasEntregadas(string Cod_Prod = "", int Cod_Campo = 0)
        {
            List<ClassCurva> resultados = bd.Database.SqlQuery<ClassCurva>("WITH t AS(Select TOP 1 R.IdAgen, R.Cod_Prod, R.Cod_Campo, R.Semana, S.Inicio FROM UV_ProdRecepcion R left join CatSemanas S on R.Semana = S.Semana and R.Temporada = S.Temporada where R.CodEstatus <> 'C' and R.Temporada = (select Temporada from CatSemanas where getdate() between Inicio and Fin) and " +
                "R.IdAgen = "+(short)Session["IdAgen"]+" and R.Cod_Prod = '"+Cod_Prod+"' and R.Cod_Campo = "+Cod_Campo+" group by R.IdAgen, R.Cod_Prod, R.Cod_Campo, R.Semana, S.Inicio order by S.Inicio) SELECT t.Semana, isnull(round(sum(R.Convertidas), 0), 0) as EntregadoT FROM t left join UV_ProdRecepcion R on t.IdAgen = R.IdAgen and t.Cod_Prod = R.Cod_Prod and t.Cod_Campo = R.Cod_Campo where R.CodEstatus <> 'C' and R.Temporada = (select Temporada from CatSemanas where getdate() between Inicio and Fin) " +
                "and R.IdAgen = " + (short)Session["IdAgen"] + " and R.Cod_Prod = '" + Cod_Prod + "' and R.Cod_Campo = " + Cod_Campo + " group by t.Semana ").ToList();
            return Json(resultados, JsonRequestBehavior.AllowGet);
        }

        public JsonResult GetAsesorList(string Cod_Prod = "", int Cod_Campo = 0)
        {
            bd.Configuration.ProxyCreationEnabled = false;
            ProdCamposCat campo = bd.ProdCamposCat.Where(x => x.Cod_Prod == Cod_Prod && x.Cod_Campo == Cod_Campo).FirstOrDefault();
            ProdAgenteCat asesor = bd.ProdAgenteCat.Where(x => x.IdAgen == campo.IdAgen).FirstOrDefault();
            return Json(asesor, JsonRequestBehavior.AllowGet);
        }

        public JsonResult GetSeguimientoF(string Cod_Prod)
        {
            bd.Configuration.ProxyCreationEnabled = false;
            Seguimiento_financ cod_prod = bdI.Seguimiento_financ.Where(x => x.Cod_Prod == Cod_Prod).FirstOrDefault();
            return Json(cod_prod, JsonRequestBehavior.AllowGet);
        }

    }
}