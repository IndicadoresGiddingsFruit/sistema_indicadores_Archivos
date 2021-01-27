using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Web.Routing;

namespace Sistema_Indicadores
{
    public class RouteConfig
    {
        public static void RegisterRoutes(RouteCollection routes)
        {
            routes.IgnoreRoute("{resource}.axd/{*pathInfo}");

            routes.MapRoute(
            name: "Arandano",
            url: "EstimacionBerries/Arandano/",
            defaults: new { controller = "EstimacionBerries", action = "Arandano" }
            );

            routes.MapRoute(
            name: "Frambuesa",
            url: "EstimacionBerries/Frambuesa/",
            defaults: new { controller = "EstimacionBerries", action = "Frambuesa" }
            );

            routes.MapRoute(
            name: "Zarzamora",
            url: "EstimacionBerries/Zarzamora/",
            defaults: new { controller = "EstimacionBerries", action = "Zarzamora" }
            );

            routes.MapRoute(
            name: "Fresa",
            url: "Indicadores/Fresa/",
            defaults: new { controller = "EstimacionBerries", action = "Fresa" }
            );

            routes.MapRoute(
            name: "Nuevo_Analisis_Residuo",
            url: "Muestreo/Nuevo_Analisis_Residuo/",
            defaults: new { controller = "Muestreo", action = "Nuevo_Analisis_Residuo" }
            );

            routes.MapRoute(
            name: "Analisis_Residuos",
            url: "Muestreo/Analisis_Residuos/",
            defaults: new { controller = "Muestreo", action = "Analisis_Residuos" }
            );

            routes.MapRoute(
            name: "Resultados_Analisis",
            url: "Muestreo/Resultados_Analisis/",
            defaults: new { controller = "Muestreo", action = "Resultados_Analisis" }
            );

            routes.MapRoute(
            name: "EditSetSolicitud",
            url: "Muestreo/Comentarios/",
            defaults: new { controller = "Muestreo", action = "EditSetSolicitud" }
            );

            routes.MapRoute(
            name: "Muestreo",
            url: "Muestreo/Muestreo/",
            defaults: new { controller = "Muestreo", action = "Muestreo" }
            );

            routes.MapRoute(
            name: "SetSolicitud",
            url: "Muestreo/SetSolicitud/",
            defaults: new { controller = "Muestreo", action = "SetSolicitud" }
            );

            routes.MapRoute(
            name: "Evaluacion",
            url: "Muestreo/Evaluacion/",
            defaults: new { controller = "Muestreo", action = "Evaluacion" }
            );

            routes.MapRoute(
            name: "Liberar_Estatus",
            url: "Muestreo/Liberar_Estatus/",
            defaults: new { controller = "Muestreo", action = "Liberar_Estatus" }
            );

            routes.MapRoute(
            name: "Muestreos_Liberados",
            url: "Muestreo/Muestreos_Liberados/",
            defaults: new { controller = "Muestreo", action = "Muestreos_Liberados" }
            );

            routes.MapRoute(
            name: "Comentarios",
            url: "Indicadores/Comentarios/",
            defaults: new { controller = "Indicadores", action = "Comentarios" }
            );

            routes.MapRoute(
            name: "Visitas",
            url: "Indicadores/Visitas/",
            defaults: new { controller = "Indicadores", action = "Visitas" }
            );

           routes.MapRoute(
           name: "ReporteVisitas",
           url: "Indicadores/ReporteVisitas/",
           defaults: new { controller = "Indicadores", action = "ReporteVisitas" }
           );

            routes.MapRoute(
            name: "Produccion",
            url: "Indicadores/Produccion/",
            defaults: new { controller = "Indicadores", action = "Produccion" }
            );

            routes.MapRoute(
            name: "Curva",
            url: "Indicadores/Curva/",
            defaults: new { controller = "Indicadores", action = "Curva" }
            );

            routes.MapRoute(
            name: "CurvaNueva",
            url: "Indicadores/CurvaNueva/",
            defaults: new { controller = "Indicadores", action = "CurvaNueva" }
            );

            routes.MapRoute(
             name: "Default",
             url: "{controller}/{action}/{id}",
             defaults: new { controller = "Login", action = "Home", id = UrlParameter.Optional }
            );
        }
    }
}
