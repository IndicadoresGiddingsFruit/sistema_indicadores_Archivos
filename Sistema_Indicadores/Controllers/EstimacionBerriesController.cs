using OfficeOpenXml;
using Sistema_Indicadores.Clases;
using Sistema_Indicadores.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Web.Mvc;

namespace Sistema_Indicadores.Controllers
{
    public class EstimacionBerriesController : Controller
    {
        SeasonSun1213Entities5 bd = new SeasonSun1213Entities5();
        JsonResult dtaEjecucionTarea = default(JsonResult);

        //ZARZAMORA
        public ActionResult Zarzamora()
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

        public JsonResult RecepcionZarzamora()
        {
            List<R_Berries> recepcion_real = bd.Database.SqlQuery<R_Berries>("SELECT 'TOTAL' as 'SEMANA', round(isnull(V.[37],0),0) as _37,round(isnull(V.[38],0),0) as _38,round(isnull(V.[39],0),0) as _39,round(isnull(V.[40],0),0) as _40,round(isnull(V.[41],0),0) as _41,round(isnull(V.[42],0),0) as _42,round(isnull(V.[43],0),0) as _43,round(isnull(V.[44],0),0) as _44,round(isnull(V.[45],0),0) as _45,round(isnull(V.[46],0),0) as _46,round(isnull(V.[47],0),0) as _47,round(isnull(V.[48],0),0) as _48,round(isnull(V.[49],0),0) as _49,round(isnull(V.[50],0),0) as _50,round(isnull(V.[51],0),0) as _51,round(isnull(V.[52],0),0) as _52,round(isnull(V.[1],0),0) as _1,round(isnull(V.[2],0),0) as _2,round(isnull(V.[3],0),0) as _3,round(isnull(V.[4],0),0) as _4,round(isnull(V.[5],0),0) as _5,round(isnull(V.[6],0),0) as _6,round(isnull(V.[7],0),0) as _7,round(isnull(V.[8],0),0) as _8,round(isnull(V.[9],0),0) as _9,round(isnull(V.[10],0),0) as _10,round(isnull(V.[11],0),0) as _11,round(isnull(V.[12],0),0) as _12,round(isnull(V.[13],0),0) as _13,round(isnull(V.[14],0),0) as _14,round(isnull(V.[15],0),0) as _15,round(isnull(V.[16],0),0) as _16,round(isnull(V.[17],0),0) as _17,round(isnull(V.[18],0),0) as _18,round(isnull(V.[19],0),0) as _19,round(isnull(V.[20],0),0) as _20,round(isnull(V.[21],0),0) as _21,round(isnull(V.[22],0),0) as _22,round(isnull(V.[23],0),0) as _23,round(isnull(V.[24],0),0) as _24,round(isnull(V.[25],0),0) as _25 " +
                "FROM(Select * from(SELECT Semana, SUM(Convertidas) AS Convertidas FROM SEasonsun1213..UV_ProdRecepcion where CodEstatus <> 'C' and temporada = (select temporada from CatSemanas where GETDATE() between Inicio and Fin) and DescProducto = 'ZARZAMORA' GROUP BY Semana Union All " +
                "SELECT Semana, SUM(Convertidas) AS Convertidas FROM SeasonPlan..UV_ProdRecepcion where CodEstatus <> 'C' and temporada = (select temporada from CatSemanas where GETDATE() between Inicio and Fin) and DescProducto = 'ZARZAMORA' GROUP BY Semana" +
                ")V PIVOT(SUM(Convertidas) FOR Semana in ([37],[38],[39],[40],[41],[42],[43],[44],[45],[46],[47],[48],[49],[50],[51],[52],[1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],[17],[18],[19],[20],[21],[22],[23],[24],[25]))X)V").ToList();
            return Json(recepcion_real, JsonRequestBehavior.AllowGet);
        }

        public JsonResult EuropaZ()
        {
            List<E_Berries> estimacion = bd.Database.SqlQuery<E_Berries>("Select round(isnull(V.[37],0),0) as _37,round(isnull(V.[38],0),0) as _38,round(isnull(V.[39],0),0) as _39,round(isnull(V.[40],0),0) as _40,round(isnull(V.[41],0),0) as _41,round(isnull(V.[42],0),0) as _42,round(isnull(V.[43],0),0) as _43,round(isnull(V.[44],0),0) as _44,round(isnull(V.[45],0),0) as _45,round(isnull(V.[46],0),0) as _46,round(isnull(V.[47],0),0) as _47,round(isnull(V.[48],0),0) as _48,round(isnull(V.[49],0),0) as _49,round(isnull(V.[50],0),0) as _50,round(isnull(V.[51],0),0) as _51,round(isnull(V.[52],0),0) as _52,round(isnull(V.[1],0),0) as _1,round(isnull(V.[2],0),0) as _2,round(isnull(V.[3],0),0) as _3,round(isnull(V.[4],0),0) as _4,round(isnull(V.[5],0),0) as _5,round(isnull(V.[6],0),0) as _6,round(isnull(V.[7],0),0) as _7,round(isnull(V.[8],0),0) as _8,round(isnull(V.[9],0),0) as _9,round(isnull(V.[10],0),0) as _10,round(isnull(V.[11],0),0) as _11,round(isnull(V.[12],0),0) as _12,round(isnull(V.[13],0),0) as _13,round(isnull(V.[14],0),0) as _14,round(isnull(V.[15],0),0) as _15,round(isnull(V.[16],0),0) as _16,round(isnull(V.[17],0),0) as _17,round(isnull(V.[18],0),0) as _18,round(isnull(V.[19],0),0) as _19,round(isnull(V.[20],0),0) as _20,round(isnull(V.[21],0),0) as _21,round(isnull(V.[22],0),0) as _22,round(isnull(V.[23],0),0) as _23,round(isnull(V.[24],0),0) as _24,round(isnull(V.[25],0),0) as _25 from (Select * from(SELECT E.Concepto, S.Inicio, E.Semanas, E.Cantidad FROM Estimacion_Berries E left join CatSemanas S on E.Temporada=S.Temporada AND E.Semana=S.Semana where E.Temporada=(select temporada from CatSemanas where GETDATE() between Inicio and Fin) and " +
                "E.Concepto = 'EUROPA' and E.Cultivo = 1)V PIVOT(SUM(Cantidad) For Semanas in ([37],[38],[39],[40],[41],[42],[43],[44],[45],[46],[47],[48],[49],[50],[51],[52],[1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],[17],[18],[19],[20],[21],[22],[23],[24],[25]))X)V ORDER BY V.Inicio").ToList();
            return Json(estimacion, JsonRequestBehavior.AllowGet);
        }

        public JsonResult AlwaysFreshZ()
        {
            List<E_Berries> estimacion = bd.Database.SqlQuery<E_Berries>("Select round(isnull(V.[37],0),0) as _37,round(isnull(V.[38],0),0) as _38,round(isnull(V.[39],0),0) as _39,round(isnull(V.[40],0),0) as _40,round(isnull(V.[41],0),0) as _41,round(isnull(V.[42],0),0) as _42,round(isnull(V.[43],0),0) as _43,round(isnull(V.[44],0),0) as _44,round(isnull(V.[45],0),0) as _45,round(isnull(V.[46],0),0) as _46,round(isnull(V.[47],0),0) as _47,round(isnull(V.[48],0),0) as _48,round(isnull(V.[49],0),0) as _49,round(isnull(V.[50],0),0) as _50,round(isnull(V.[51],0),0) as _51,round(isnull(V.[52],0),0) as _52,round(isnull(V.[1],0),0) as _1,round(isnull(V.[2],0),0) as _2,round(isnull(V.[3],0),0) as _3,round(isnull(V.[4],0),0) as _4,round(isnull(V.[5],0),0) as _5,round(isnull(V.[6],0),0) as _6,round(isnull(V.[7],0),0) as _7,round(isnull(V.[8],0),0) as _8,round(isnull(V.[9],0),0) as _9,round(isnull(V.[10],0),0) as _10,round(isnull(V.[11],0),0) as _11,round(isnull(V.[12],0),0) as _12,round(isnull(V.[13],0),0) as _13,round(isnull(V.[14],0),0) as _14,round(isnull(V.[15],0),0) as _15,round(isnull(V.[16],0),0) as _16,round(isnull(V.[17],0),0) as _17,round(isnull(V.[18],0),0) as _18,round(isnull(V.[19],0),0) as _19,round(isnull(V.[20],0),0) as _20,round(isnull(V.[21],0),0) as _21,round(isnull(V.[22],0),0) as _22,round(isnull(V.[23],0),0) as _23,round(isnull(V.[24],0),0) as _24,round(isnull(V.[25],0),0) as _25 from(Select * from(SELECT S.Inicio, R.Semana, (R.Cantidad-C.Cantidad) Cantidad FROM(SELECT Semana, SUM(Convertidas) AS Cantidad, Temporada FROM SEasonsun1213..UV_ProdRecepcion where CodEstatus<> 'C' and temporada = (select temporada from CatSemanas where GETDATE() between Inicio and Fin) and " +
                "DescProducto = 'ZARZAMORA' GROUP BY Semana, Temporada Union All SELECT Semana, SUM(Convertidas) AS Cantidad, Temporada FROM SeasonPlan..UV_ProdRecepcion where CodEstatus <> 'C' and temporada = (select temporada from CatSemanas where GETDATE() between Inicio and Fin) and " +
                "DescProducto = 'ZARZAMORA' GROUP BY Semana,Temporada)R left join(Select Semanas, Cantidad, Temporada, Cultivo, Semana from Estimacion_Berries " +
                "where Concepto = 'EUROPA' and Cultivo = 1 and Temporada = (select temporada from CatSemanas where GETDATE() between Inicio and Fin))C on R.Temporada = C.Temporada and R.Semana = C.Semanas left join CatSemanas S on R.Temporada = S.Temporada AND C.Semana = S.Semana)V PIVOT(SUM(Cantidad) For Semana in ([37],[38],[39],[40],[41],[42],[43],[44],[45],[46],[47],[48],[49],[50],[51],[52],[1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],[17],[18],[19],[20],[21],[22],[23],[24],[25]))X)V ORDER BY V.Inicio").ToList();
            return Json(estimacion, JsonRequestBehavior.AllowGet);
        }

        public JsonResult CoyotesZ()
        {
            List<E_Berries> estimacion = bd.Database.SqlQuery<E_Berries>("Select V.Concepto, round(isnull(V.[37],0),0) as _37,round(isnull(V.[38],0),0) as _38,round(isnull(V.[39],0),0) as _39,round(isnull(V.[40],0),0) as _40,round(isnull(V.[41],0),0) as _41,round(isnull(V.[42],0),0) as _42,round(isnull(V.[43],0),0) as _43,round(isnull(V.[44],0),0) as _44,round(isnull(V.[45],0),0) as _45,round(isnull(V.[46],0),0) as _46,round(isnull(V.[47],0),0) as _47,round(isnull(V.[48],0),0) as _48,round(isnull(V.[49],0),0) as _49,round(isnull(V.[50],0),0) as _50,round(isnull(V.[51],0),0) as _51,round(isnull(V.[52],0),0) as _52,round(isnull(V.[1],0),0) as _1,round(isnull(V.[2],0),0) as _2,round(isnull(V.[3],0),0) as _3,round(isnull(V.[4],0),0) as _4,round(isnull(V.[5],0),0) as _5,round(isnull(V.[6],0),0) as _6,round(isnull(V.[7],0),0) as _7,round(isnull(V.[8],0),0) as _8,round(isnull(V.[9],0),0) as _9,round(isnull(V.[10],0),0) as _10,round(isnull(V.[11],0),0) as _11,round(isnull(V.[12],0),0) as _12,round(isnull(V.[13],0),0) as _13,round(isnull(V.[14],0),0) as _14,round(isnull(V.[15],0),0) as _15,round(isnull(V.[16],0),0) as _16,round(isnull(V.[17],0),0) as _17,round(isnull(V.[18],0),0) as _18,round(isnull(V.[19],0),0) as _19,round(isnull(V.[20],0),0) as _20,round(isnull(V.[21],0),0) as _21,round(isnull(V.[22],0),0) as _22,round(isnull(V.[23],0),0) as _23,round(isnull(V.[24],0),0) as _24,round(isnull(V.[25],0),0) as _25 " +
                "from(Select * from(SELECT E.Concepto, S.Inicio, E.Semanas, E.Cantidad " +
                "FROM Estimacion_Berries E left join CatSemanas S on E.Temporada = S.Temporada AND E.Semana = S.Semana " +
                "where E.Temporada = (select temporada from CatSemanas where GETDATE() between Inicio and Fin) and E.Concepto = 'COYOTES' and E.Cultivo = 1)V " +
                "PIVOT(SUM(Cantidad) For Semanas in ([37],[38],[39],[40],[41],[42],[43],[44],[45],[46],[47],[48],[49],[50],[51],[52],[1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],[17],[18],[19],[20],[21],[22],[23],[24],[25]))X)V ORDER BY V.Inicio").ToList();
            return Json(estimacion, JsonRequestBehavior.AllowGet);
        }

        public JsonResult AlwaysSCoyotesZ()
        {
            List<E_Berries> estimacion = bd.Database.SqlQuery<E_Berries>("Select round(isnull(V.[37],0),0) as _37,round(isnull(V.[38],0),0) as _38,round(isnull(V.[39],0),0) as _39,round(isnull(V.[40],0),0) as _40,round(isnull(V.[41],0),0) as _41,round(isnull(V.[42],0),0) as _42,round(isnull(V.[43],0),0) as _43,round(isnull(V.[44],0),0) as _44,round(isnull(V.[45],0),0) as _45,round(isnull(V.[46],0),0) as _46,round(isnull(V.[47],0),0) as _47,round(isnull(V.[48],0),0) as _48,round(isnull(V.[49],0),0) as _49,round(isnull(V.[50],0),0) as _50,round(isnull(V.[51],0),0) as _51,round(isnull(V.[52],0),0) as _52,round(isnull(V.[1],0),0) as _1,round(isnull(V.[2],0),0) as _2,round(isnull(V.[3],0),0) as _3,round(isnull(V.[4],0),0) as _4,round(isnull(V.[5],0),0) as _5,round(isnull(V.[6],0),0) as _6,round(isnull(V.[7],0),0) as _7,round(isnull(V.[8],0),0) as _8,round(isnull(V.[9],0),0) as _9,round(isnull(V.[10],0),0) as _10,round(isnull(V.[11],0),0) as _11,round(isnull(V.[12],0),0) as _12,round(isnull(V.[13],0),0) as _13,round(isnull(V.[14],0),0) as _14,round(isnull(V.[15],0),0) as _15,round(isnull(V.[16],0),0) as _16,round(isnull(V.[17],0),0) as _17,round(isnull(V.[18],0),0) as _18,round(isnull(V.[19],0),0) as _19,round(isnull(V.[20],0),0) as _20,round(isnull(V.[21],0),0) as _21,round(isnull(V.[22],0),0) as _22,round(isnull(V.[23],0),0) as _23,round(isnull(V.[24],0),0) as _24,round(isnull(V.[25],0),0) as _25 from(Select * from(SELECT S.Inicio, E.Semanas, (E.Cantidad-C.Cantidad) Cantidad FROM(Select Semanas, Cantidad, Temporada, Cultivo, Semana " +
                "from Estimacion_Berries where Concepto = 'ALWAYS FRESH')E left join(Select Semanas, Cantidad, Temporada, Cultivo " +
                "from Estimacion_Berries where Concepto = 'COYOTES')C on E.Temporada = C.Temporada and E.Semanas = C.Semanas AND E.Cultivo = C.Cultivo left join CatSemanas S on E.Temporada = S.Temporada AND E.Semana = S.Semana " +
                "where E.Temporada = (select temporada from CatSemanas where GETDATE() between Inicio and Fin) and E.Cultivo = 1)V PIVOT(SUM(Cantidad) For Semanas in ([37],[38],[39],[40],[41],[42],[43],[44],[45],[46],[47],[48],[49],[50],[51],[52],[1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],[17],[18],[19],[20],[21],[22],[23],[24],[25]))X)V ORDER BY V.Inicio").ToList();
            return Json(estimacion, JsonRequestBehavior.AllowGet);
        }

        public JsonResult EstimacionSZarzamora()
        {
            List<E_Berries> estimacion = bd.Database.SqlQuery<E_Berries>("Select V.Sem AS 'SEMANA', round(isnull(V.[37],0),0) as _37,round(isnull(V.[38],0),0) as _38,round(isnull(V.[39],0),0) as _39,round(isnull(V.[40],0),0) as _40,round(isnull(V.[41],0),0) as _41,round(isnull(V.[42],0),0) as _42,round(isnull(V.[43],0),0) as _43,round(isnull(V.[44],0),0) as _44,round(isnull(V.[45],0),0) as _45,round(isnull(V.[46],0),0) as _46,round(isnull(V.[47],0),0) as _47,round(isnull(V.[48],0),0) as _48,round(isnull(V.[49],0),0) as _49,round(isnull(V.[50],0),0) as _50,round(isnull(V.[51],0),0) as _51,round(isnull(V.[52],0),0) as _52,round(isnull(V.[1],0),0) as _1,round(isnull(V.[2],0),0) as _2,round(isnull(V.[3],0),0) as _3,round(isnull(V.[4],0),0) as _4,round(isnull(V.[5],0),0) as _5,round(isnull(V.[6],0),0) as _6,round(isnull(V.[7],0),0) as _7,round(isnull(V.[8],0),0) as _8,round(isnull(V.[9],0),0) as _9,round(isnull(V.[10],0),0) as _10,round(isnull(V.[11],0),0) as _11,round(isnull(V.[12],0),0) as _12,round(isnull(V.[13],0),0) as _13,round(isnull(V.[14],0),0) as _14,round(isnull(V.[15],0),0) as _15,round(isnull(V.[16],0),0) as _16,round(isnull(V.[17],0),0) as _17,round(isnull(V.[18],0),0) as _18,round(isnull(V.[19],0),0) as _19,round(isnull(V.[20],0),0) as _20,round(isnull(V.[21],0),0) as _21,round(isnull(V.[22],0),0) as _22,round(isnull(V.[23],0),0) as _23,round(isnull(V.[24],0),0) as _24,round(isnull(V.[25],0),0) as _25 " +
                "from(Select * from(SELECT S.Semana AS Sem, S.Inicio, E.Semanas, E.Cantidad FROM Estimacion_Berries E left join CatSemanas S on E.Temporada = S.Temporada AND E.Semana = S.Semana " +
                "where E.Temporada = (select temporada from CatSemanas where GETDATE() between Inicio and Fin) and E.Concepto = 'PROYECCION SEMANAL' and E.Cultivo = 1)V " +
                "PIVOT(SUM(Cantidad) For Semanas in ([37],[38],[39],[40],[41],[42],[43],[44],[45],[46],[47],[48],[49],[50],[51],[52],[1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],[17],[18],[19],[20],[21],[22],[23],[24],[25]))X)V ORDER BY V.Inicio").ToList();
            return Json(estimacion, JsonRequestBehavior.AllowGet);
        }

        public JsonResult EstimacionSZarzamoraP()
        {
            List<E_BerriesP> estimacion = bd.Database.SqlQuery<E_BerriesP>("select V.Sem AS 'SEMANA', (case when V._37='0%' then '' else V._37 end) as _37,(case when V._38='0%' then '' else V._38 end) as _38,(case when V._39='0%' then '' else V._39 end) as _39,(case when V._40='0%' then '' else V._40 end) as _40,(case when V._41='0%' then '' else V._41 end) as _41,(case when V._42='0%' then '' else V._42 end) as _42,(case when V._43='0%' then '' else V._43 end) as _43,(case when V._44='0%' then '' else V._44 end) as _44,(case when V._45='0%' then '' else V._45 end) as _45,(case when V._46='0%' then '' else V._46 end) as _46,(case when V._47='0%' then '' else V._47 end) as _47,(case when V._48='0%' then '' else V._48 end) as _48,(case when V._49='0%' then '' else V._49 end) as _49,(case when V._50='0%' then '' else V._50 end) as _50,(case when V._51='0%' then '' else V._51 end) as _51,(case when V._52='0%' then '' else V._52 end) as _52,(case when V._1='0%' then '' else V._1 end) as _1,(case when V._2='0%' then '' else V._2 end) as _2,(case when V._3='0%' then '' else V._3 end) as _3,(case when V._4='0%' then '' else V._4 end) as _4,(case when V._5='0%' then '' else V._5 end) as _5,(case when V._6='0%' then '' else V._6 end) as _6,(case when V._7='0%' then '' else V._7 end) as _7,(case when V._8='0%' then '' else V._8 end) as _8,(case when V._9='0%' then '' else V._9 end) as _9,(case when V._10='0%' then '' else V._10 end) as _10,(case when V._11='0%' then '' else V._11 end) as _11,(case when V._12='0%' then '' else V._12 end) as _12,(case when V._13='0%' then '' else V._13 end) as _13,(case when V._14='0%' then '' else V._14 end) as _14,(case when V._15='0%' then '' else V._15 end) as _15,(case when V._16='0%' then '' else V._16 end) as _16,(case when V._17='0%' then '' else V._17 end) as _17,(case when V._18='0%' then '' else V._18 end) as _18,(case when V._19='0%' then '' else V._19 end) as _19,(case when V._20='0%' then '' else V._20 end) as _20,(case when V._21='0%' then '' else V._21 end) as _21,(case when V._22='0%' then '' else V._22 end) as _22,(case when V._23='0%' then '' else V._23 end) as _23,(case when V._24='0%' then '' else V._24 end) as _24,(case when V._25='0%' then '' else V._25 end) as _25 from (select V.Inicio, V.Sem, cast(V._37 as varchar)+'%' as _37,cast(V._38 as varchar)+'%' as _38,cast(V._39 as varchar)+'%' as _39,cast(V._40 as varchar)+'%' as _40,cast(V._41 as varchar)+'%' as _41,cast(V._42 as varchar)+'%' as _42,cast(V._43 as varchar)+'%' as _43,cast(V._44 as varchar)+'%' as _44,cast(V._45 as varchar)+'%' as _45,cast(V._46 as varchar)+'%' as _46,cast(V._47 as varchar)+'%' as _47,cast(V._48 as varchar)+'%' as _48,cast(V._49 as varchar)+'%' as _49,cast(V._50 as varchar)+'%' as _50,cast(V._51 as varchar)+'%' as _51,cast(V._52 as varchar)+'%' as _52,cast(V._1 as varchar)+'%' as _1,cast(V._2 as varchar)+'%' _2,cast(V._3 as varchar)+'%' as _3,cast(V._4 as varchar)+'%' as _4,cast(V._5 as varchar)+'%' as _5,cast(V._6 as varchar)+'%' as _6,cast(V._7 as varchar)+'%' as _7,cast(V._8 as varchar)+'%' as _8,cast(V._9 as varchar)+'%' as _9,cast(V._10 as varchar)+'%' as _10,cast(V._11 as varchar)+'%' as _11,cast(V._12 as varchar)+'%' as _12,cast(V._13 as varchar)+'%' as _13,cast(V._14 as varchar)+'%' as _14,cast(V._15 as varchar)+'%' as _15,cast(V._16 as varchar)+'%' as _16,cast(V._17 as varchar)+'%' as _17,cast(V._18 as varchar)+'%' as _18,cast(V._19 as varchar)+'%' as _19,cast(V._20 as varchar)+'%'as _20,cast(V._21 as varchar)+'%' as _21,cast(V._22 as varchar)+'%' as _22,cast(V._23 as varchar)+'%' as _23,cast(V._24 as varchar)+'%' as _24,cast(V._25 as varchar)+'%' as _25 from (SELECT V.Inicio, V.Sem,(case when A._37=0 then 0 else case when V._37=0 then 0 else round((A._37/V._37)*100,0) end end) as _37,(case when A._38=0 then 0 else case when V._38=0 then 0 else round((A._38/V._38)*100,0) end end) as _38,(case when A._39=0 then 0 else case when V._39=0 then 0 else round((A._39/V._39)*100,0) end end) as _39,(case when A._40=0 then 0 else case when V._40=0 then 0 else round((A._40/V._40)*100,0) end end) as _40,(case when A._41=0 then 0 else case when V._41=0 then 0 else round((A._41/V._41)*100,0) end end) as _41,(case when A._42=0 then 0 else case when V._42=0 then 0 else round((A._42/V._42)*100,0) end end) as _42,(case when A._43=0 then 0 else case when V._43=0 then 0 else round((A._43/V._43)*100,0) end end) as _43,(case when A._44=0 then 0 else case when V._44=0 then 0 else round((A._44/V._44)*100,0) end end) as _44,(case when A._45=0 then 0 else case when V._45=0 then 0 else round((A._45/V._45)*100,0) end end) as _45,(case when A._46=0 then 0 else case when V._46=0 then 0 else round((A._46/V._46)*100,0) end end) as _46,(case when A._47=0 then 0 else case when V._47=0 then 0 else round((A._47/V._47)*100,0) end end) as _47,(case when A._48=0 then 0 else case when V._48=0 then 0 else round((A._48/V._48)*100,0) end end) as _48,(case when A._49=0 then 0 else case when V._49=0 then 0 else round((A._49/V._49)*100,0) end end) as _49,(case when A._50=0 then 0 else case when V._50=0 then 0 else round((A._50/V._50)*100,0) end end) as _50,(case when A._51=0 then 0 else case when V._51=0 then 0 else round((A._51/V._51)*100,0) end end) as _51,(case when A._52=0 then 0 else case when V._52=0 then 0 else round((A._52/V._52)*100,0) end end) as _52,(case when A._1=0 then 0 else case when V._1=0 then 0 else round((A._1/V._1)*100,0) end end) as _1,(case when A._2=0 then 0 else case when V._2=0 then 0 else round((A._2/V._2)*100,0) end end) as _2,(case when A._3=0 then 0 else case when V._3=0 then 0 else round((A._3/V._3)*100,0) end end) as _3,(case when A._4=0 then 0 else case when V._4=0 then 0 else round((A._4/V._4)*100,0) end end) as _4,(case when A._5=0 then 0 else case when V._5=0 then 0 else round((A._5/V._5)*100,0) end end) as _5,(case when A._6=0 then 0 else case when V._6=0 then 0 else round((A._6/V._6)*100,0) end end) as _6,(case when A._7=0 then 0 else case when V._7=0 then 0 else round((A._7/V._7)*100,0) end end) as _7,(case when A._8=0 then 0 else case when V._8=0 then 0 else round((A._8/V._8)*100,0) end end) as _8,(case when A._9=0 then 0 else case when V._9=0 then 0 else round((A._9/V._9)*100,0) end end) as _9,(case when A._10=0 then 0 else case when V._10=0 then 0 else round((A._10/V._10)*100,0) end end) as _10,(case when A._11=0 then 0 else case when V._11=0 then 0 else round((A._11/V._11)*100,0) end end) as _11,(case when A._12=0 then 0 else case when V._12=0 then 0 else round((A._12/V._12)*100,0) end end) as _12,(case when A._13=0 then 0 else case when V._13=0 then 0 else round((A._13/V._13)*100,0) end end) as _13,(case when A._14=0 then 0 else case when V._14=0 then 0 else round((A._14/V._14)*100,0) end end) as _14,(case when A._15=0 then 0 else case when V._15=0 then 0 else round((A._15/V._15)*100,0) end end) as _15,(case when A._16=0 then 0 else case when V._16=0 then 0 else round((A._16/V._16)*100,0) end end) as _16,(case when A._17=0 then 0 else case when V._17=0 then 0 else round((A._17/V._17)*100,0) end end) as _17,(case when A._18=0 then 0 else case when V._18=0 then 0 else round((A._18/V._18)*100,0) end end) as _18,(case when A._19=0 then 0 else case when V._19=0 then 0 else round((A._19/V._19)*100,0) end end) as _19,(case when A._20=0 then 0 else case when V._20=0 then 0 else round((A._20/V._20)*100,0) end end) as _20,(case when A._21=0 then 0 else case when V._21=0 then 0 else round((A._21/V._21)*100,0) end end) as _21,(case when A._22=0 then 0 else case when V._22=0 then 0 else round((A._22/V._22)*100,0) end end) as _22,(case when A._23=0 then 0 else case when V._23=0 then 0 else round((A._23/V._23)*100,0) end end) as _23,(case when A._24=0 then 0 else case when V._24=0 then 0 else round((A._24/V._24)*100,0) end end) as _24,(case when A._25=0 then 0 else case when V._25=0 then 0 else round((A._25/V._25)*100,0) end end) as _25 FROM(Select V.Cultivo, isnull(V.[37],0) as _37,isnull(V.[38],0) as _38,isnull(V.[39],0) as _39,isnull(V.[40],0) as _40,isnull(V.[41],0) as _41,isnull(V.[42],0) as _42,isnull(V.[43],0) as _43,isnull(V.[44],0) as _44,isnull(V.[45],0)as _45,isnull(V.[46],0) as _46,isnull(V.[47],0) as _47,isnull(V.[48],0) as _48,isnull(V.[49],0) as _49,isnull(V.[50],0) as _50,isnull(V.[51],0) as _51,isnull(V.[52],0) as _52,isnull(V.[1],0) as _1,isnull(V.[2],0) as _2,isnull(V.[3],0) as _3,isnull(V.[4],0) as _4,isnull(V.[5],0) as _5,isnull(V.[6],0) as _6,isnull(V.[7],0) as _7,isnull(V.[8],0) as _8,isnull(V.[9],0) as _9,isnull(V.[10],0) as _10,isnull(V.[11],0) as _11,isnull(V.[12],0) as _12,isnull(V.[13],0) as _13,isnull(V.[14],0) as _14,isnull(V.[15],0) as _15,isnull(V.[16],0) as _16,isnull(V.[17],0) as _17,isnull(V.[18],0) as _18,isnull(V.[19],0) as _19,isnull(V.[20],0) as _20,isnull(V.[21],0) as _21,isnull(V.[22],0) as _22,isnull(V.[23],0) as _23,isnull(V.[24],0) as _24,isnull(V.[25],0) as _25 from (Select * from(SELECT E.Concepto, S.Inicio, E.Semanas, E.Cantidad, E.Cultivo FROM Estimacion_Berries E left join CatSemanas S on E.Temporada=S.Temporada AND E.Semana=S.Semana where E.Temporada=(select temporada from CatSemanas where GETDATE() between Inicio and Fin) and " +
                "E.Concepto = 'ALWAYS FRESH' and E.Cultivo = 1)V PIVOT(SUM(Cantidad) For Semanas in ([37],[38],[39],[40],[41],[42],[43],[44],[45],[46],[47],[48],[49],[50],[51],[52],[1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],[17],[18],[19],[20],[21],[22],[23],[24],[25]))X)V)A left join(Select V.Inicio, V.Cultivo, V.Sem, isnull(V.[37], 0) as _37, isnull(V.[38], 0) as _38, isnull(V.[39], 0) as _39, isnull(V.[40], 0) as _40, isnull(V.[41], 0) as _41, isnull(V.[42], 0) as _42, isnull(V.[43], 0) as _43, isnull(V.[44], 0) as _44, isnull(V.[45], 0) as _45, isnull(V.[46], 0) as _46, isnull(V.[47], 0) as _47, isnull(V.[48], 0) as _48, isnull(V.[49], 0) as _49, isnull(V.[50], 0) as _50, isnull(V.[51], 0) as _51, isnull(V.[52], 0) as _52, isnull(V.[1], 0) as _1, isnull(V.[2], 0) as _2, isnull(V.[3], 0) as _3, isnull(V.[4], 0) as _4, isnull(V.[5], 0) as _5, isnull(V.[6], 0) as _6, isnull(V.[7], 0) as _7, isnull(V.[8], 0) as _8, isnull(V.[9], 0) as _9, isnull(V.[10], 0) as _10, isnull(V.[11], 0) as _11, isnull(V.[12], 0) as _12, isnull(V.[13], 0) as _13, isnull(V.[14], 0) as _14, isnull(V.[15], 0) as _15, isnull(V.[16], 0) as _16, isnull(V.[17], 0) as _17, isnull(V.[18], 0) as _18, isnull(V.[19], 0) as _19, isnull(V.[20], 0) as _20, isnull(V.[21], 0) as _21, isnull(V.[22], 0) as _22, isnull(V.[23], 0) as _23, isnull(V.[24], 0) as _24, isnull(V.[25], 0) as _25 " +
                "from(select * from(SELECT S.Semana AS Sem, S.Inicio, E.Semanas, E.Cantidad, E.Cultivo FROM Estimacion_Berries E left join CatSemanas S on E.Temporada = S.Temporada AND E.Semana = S.Semana where E.Temporada = (select temporada from CatSemanas where GETDATE() between Inicio and Fin) and " +
                "E.Concepto = 'PROYECCION SEMANAL' and E.Cultivo = 1)V PIVOT(SUM(Cantidad) For Semanas in ([37],[38],[39],[40],[41],[42],[43],[44],[45],[46],[47],[48],[49],[50],[51],[52],[1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],[17],[18],[19],[20],[21],[22],[23],[24],[25]))X)V )V on A.Cultivo = V.Cultivo)V)V ORDER BY V.Inicio").ToList();
            return Json(estimacion, JsonRequestBehavior.AllowGet);
        }

        [HttpPost]
        public ActionResult Zarzamora(int[] array, string concepto)
        {
            JsonResult dtaEjecucionTarea = default(JsonResult);
                if (concepto == "PROYECCION SEMANAL")
                {
                    if (Update_Zarzamora(array))
                    {
                        dtaEjecucionTarea = Json(new
                        {
                            rstProceso = "true",
                            MessageGestion = "Cambios guardados con éxito"
                        });
                    }
                    else
                    {
                        dtaEjecucionTarea = Json(new
                        {
                            rstProceso = "false",
                            MessageGestion = "Error, algo salió mal, intente de nuevo"
                        });
                    }
                }
                else
                {
                    if (Update_ZarzamoraAFC(array, concepto))
                    {
                        dtaEjecucionTarea = Json(new
                        {
                            rstProceso = "true",
                            MessageGestion = "Cambios guardados con éxito"
                        });
                    }
                    else
                    {
                        dtaEjecucionTarea = Json(new
                        {
                            rstProceso = "false",
                            MessageGestion = "Error, algo salió mal, intente de nuevo"
                        });
                    }
                }
                return dtaEjecucionTarea;            
        }

        public bool Update_Zarzamora(int[] array)
        {
            if (array != null)
            {
                var sem = array[0];
                var _37 = array[1];
                var _38 = array[2];
                var _39 = array[3];
                var _40 = array[4];
                var _41 = array[5];
                var _42 = array[6];
                var _43 = array[7];
                var _44 = array[8];
                var _45 = array[9];
                var _46 = array[10];
                var _47 = array[11];
                var _48 = array[12];
                var _49 = array[13];
                var _50 = array[14];
                var _51 = array[15];
                var _52 = array[16];
                var _1 = array[17];
                var _2 = array[18];
                var _3 = array[19];
                var _4 = array[20];
                var _5 = array[21];
                var _6 = array[22];
                var _7 = array[23];
                var _8 = array[24];
                var _9 = array[25];
                var _10 = array[26];
                var _11 = array[27];
                var _12 = array[28];
                var _13 = array[29];
                var _14 = array[30];
                var _15 = array[31];
                var _16 = array[32];
                var _17 = array[33];
                var _18 = array[34];
                var _19 = array[35];
                var _20 = array[36];
                var _21 = array[37];
                var _22 = array[38];
                var _23 = array[39];
                var _24 = array[40];
                var _25 = array[41];

                var item_37 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 37 && x.Cultivo == 1 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_37.Cantidad = _37;
                item_37.Fecha = DateTime.Now;

                var item_38 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 38 && x.Cultivo == 1 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_38.Cantidad = _38;
                item_38.Fecha = DateTime.Now;

                var item_39 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 39 && x.Cultivo == 1 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_39.Cantidad = _39;
                item_39.Fecha = DateTime.Now;

                var item_40 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 40 && x.Cultivo == 1 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_40.Cantidad = _40;
                item_40.Fecha = DateTime.Now;

                var item_41 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 41 && x.Cultivo == 1 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_41.Cantidad = _41;
                item_41.Fecha = DateTime.Now;

                var item_42 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 42 && x.Cultivo == 1 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_42.Cantidad = _42;
                item_42.Fecha = DateTime.Now;

                var item_43 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 43 && x.Cultivo == 1 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_43.Cantidad = _43;
                item_43.Fecha = DateTime.Now;

                var item_44 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 44 && x.Cultivo == 1 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_44.Cantidad = _44;
                item_44.Fecha = DateTime.Now;

                var item_45 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 45 && x.Cultivo == 1 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_45.Cantidad = _45;
                item_45.Fecha = DateTime.Now;

                var item_46 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 46 && x.Cultivo == 1 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_46.Cantidad = _46;
                item_46.Fecha = DateTime.Now;

                var item_47 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 47 && x.Cultivo == 1 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_47.Cantidad = _47;
                item_47.Fecha = DateTime.Now;

                var item_48 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 48 && x.Cultivo == 1 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_48.Cantidad = _48;
                item_48.Fecha = DateTime.Now;

                var item_49 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 49 && x.Cultivo == 1 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_49.Cantidad = _49;
                item_49.Fecha = DateTime.Now;

                var item_50 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 50 && x.Cultivo == 1 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_50.Cantidad = _50;
                item_50.Fecha = DateTime.Now;

                var item_51 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 51 && x.Cultivo == 1 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_51.Cantidad = _51;
                item_51.Fecha = DateTime.Now;

                var item_52 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 52 && x.Cultivo == 1 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_52.Cantidad = _52;
                item_52.Fecha = DateTime.Now;

                var item_1 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 1 && x.Cultivo == 1 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_1.Cantidad = _1;
                item_1.Fecha = DateTime.Now;

                var item_2 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 2 && x.Cultivo == 1 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_2.Cantidad = _2;
                item_2.Fecha = DateTime.Now;

                var item_3 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 3 && x.Cultivo == 1 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_3.Cantidad = _3;
                item_3.Fecha = DateTime.Now;

                var item_4 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 4 && x.Cultivo == 1 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_4.Cantidad = _4;
                item_4.Fecha = DateTime.Now;

                var item_5 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 5 && x.Cultivo == 1 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_5.Cantidad = _5;
                item_5.Fecha = DateTime.Now;

                var item_6 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 6 && x.Cultivo == 1 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_6.Cantidad = _6;
                item_6.Fecha = DateTime.Now;

                var item_7 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 7 && x.Cultivo == 1 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_7.Cantidad = _7;
                item_7.Fecha = DateTime.Now;

                var item_8 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 8 && x.Cultivo == 1 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_8.Cantidad = _8;
                item_8.Fecha = DateTime.Now;

                var item_9 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 9 && x.Cultivo == 1 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_9.Cantidad = _9;
                item_9.Fecha = DateTime.Now;

                var item_10 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 10 && x.Cultivo == 1 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_10.Cantidad = _10;
                item_10.Fecha = DateTime.Now;

                var item_11 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 11 && x.Cultivo == 1 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_11.Cantidad = _11;
                item_11.Fecha = DateTime.Now;

                var item_12 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 12 && x.Cultivo == 1 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_12.Cantidad = _12;
                item_12.Fecha = DateTime.Now;

                var item_13 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 13 && x.Cultivo == 1 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_13.Cantidad = _13;
                item_13.Fecha = DateTime.Now;

                var item_14 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 14 && x.Cultivo == 1 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_14.Cantidad = _14;
                item_14.Fecha = DateTime.Now;

                var item_15 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 15 && x.Cultivo == 1 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_15.Cantidad = _15;
                item_15.Fecha = DateTime.Now;

                var item_16 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 16 && x.Cultivo == 1 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_16.Cantidad = _16;
                item_16.Fecha = DateTime.Now;

                var item_17 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 17 && x.Cultivo == 1 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_17.Cantidad = _17;
                item_17.Fecha = DateTime.Now;

                var item_18 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 18 && x.Cultivo == 1 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_18.Cantidad = _18;
                item_18.Fecha = DateTime.Now;

                var item_19 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 19 && x.Cultivo == 1 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_19.Cantidad = _19;
                item_19.Fecha = DateTime.Now;

                var item_20 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 20 && x.Cultivo == 1 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_20.Cantidad = _20;
                item_20.Fecha = DateTime.Now;

                var item_21 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 21 && x.Cultivo == 1 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_21.Cantidad = _21;
                item_21.Fecha = DateTime.Now;

                var item_22 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 22 && x.Cultivo == 1 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_22.Cantidad = _22;
                item_22.Fecha = DateTime.Now;

                var item_23 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 23 && x.Cultivo == 1 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_23.Cantidad = _23;
                item_23.Fecha = DateTime.Now;

                var item_24 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 24 && x.Cultivo == 1 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_24.Cantidad = _24;
                item_24.Fecha = DateTime.Now;

                var item_25 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 25 && x.Cultivo == 1 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_25.Cantidad = _25;
                item_25.Fecha = DateTime.Now;

                bd.SaveChanges();
                return true;
            }
            else
            {
                return false;
            }
        }

        public bool Update_ZarzamoraAFC(int[] array, string concepto)
        {
            if (array != null)
            {
                var _37 = array[0];
                var _38 = array[1];
                var _39 = array[2];
                var _40 = array[3];
                var _41 = array[4];
                var _42 = array[5];
                var _43 = array[6];
                var _44 = array[7];
                var _45 = array[8];
                var _46 = array[9];
                var _47 = array[10];
                var _48 = array[11];
                var _49 = array[12];
                var _50 = array[13];
                var _51 = array[14];
                var _52 = array[15];
                var _1 = array[16];
                var _2 = array[17];
                var _3 = array[18];
                var _4 = array[19];
                var _5 = array[20];
                var _6 = array[21];
                var _7 = array[22];
                var _8 = array[23];
                var _9 = array[24];
                var _10 = array[25];
                var _11 = array[26];
                var _12 = array[27];
                var _13 = array[28];
                var _14 = array[29];
                var _15 = array[30];
                var _16 = array[31];
                var _17 = array[32];
                var _18 = array[33];
                var _19 = array[34];
                var _20 = array[35];
                var _21 = array[36];
                var _22 = array[37];
                var _23 = array[38];
                var _24 = array[39];
                var _25 = array[40];

                var item_37 = bd.Estimacion_Berries.Where(x => x.Semanas == 37 && x.Cultivo == 1 && x.Concepto == concepto).FirstOrDefault();
                item_37.Cantidad = _37;
                item_37.Fecha = DateTime.Now;

                var item_38 = bd.Estimacion_Berries.Where(x => x.Semanas == 38 && x.Cultivo == 1 && x.Concepto == concepto).FirstOrDefault();
                item_38.Cantidad = _38;
                item_38.Fecha = DateTime.Now;

                var item_39 = bd.Estimacion_Berries.Where(x => x.Semanas == 39 && x.Cultivo == 1 && x.Concepto == concepto).FirstOrDefault();
                item_39.Cantidad = _39;
                item_39.Fecha = DateTime.Now;

                var item_40 = bd.Estimacion_Berries.Where(x => x.Semanas == 40 && x.Cultivo == 1 && x.Concepto == concepto).FirstOrDefault();
                item_40.Cantidad = _40;
                item_40.Fecha = DateTime.Now;

                var item_41 = bd.Estimacion_Berries.Where(x => x.Semanas == 41 && x.Cultivo == 1 && x.Concepto == concepto).FirstOrDefault();
                item_41.Cantidad = _41;
                item_41.Fecha = DateTime.Now;

                var item_42 = bd.Estimacion_Berries.Where(x => x.Semanas == 42 && x.Cultivo == 1 && x.Concepto == concepto).FirstOrDefault();
                item_42.Cantidad = _42;
                item_42.Fecha = DateTime.Now;

                var item_43 = bd.Estimacion_Berries.Where(x => x.Semanas == 43 && x.Cultivo == 1 && x.Concepto == concepto).FirstOrDefault();
                item_43.Cantidad = _43;
                item_43.Fecha = DateTime.Now;

                var item_44 = bd.Estimacion_Berries.Where(x => x.Semanas == 44 && x.Cultivo == 1 && x.Concepto == concepto).FirstOrDefault();
                item_44.Cantidad = _44;
                item_44.Fecha = DateTime.Now;

                var item_45 = bd.Estimacion_Berries.Where(x => x.Semanas == 45 && x.Cultivo == 1 && x.Concepto == concepto).FirstOrDefault();
                item_45.Cantidad = _45;
                item_45.Fecha = DateTime.Now;

                var item_46 = bd.Estimacion_Berries.Where(x => x.Semanas == 46 && x.Cultivo == 1 && x.Concepto == concepto).FirstOrDefault();
                item_46.Cantidad = _46;
                item_46.Fecha = DateTime.Now;

                var item_47 = bd.Estimacion_Berries.Where(x => x.Semanas == 47 && x.Cultivo == 1 && x.Concepto == concepto).FirstOrDefault();
                item_47.Cantidad = _47;
                item_47.Fecha = DateTime.Now;

                var item_48 = bd.Estimacion_Berries.Where(x => x.Semanas == 48 && x.Cultivo == 1 && x.Concepto == concepto).FirstOrDefault();
                item_48.Cantidad = _48;
                item_48.Fecha = DateTime.Now;

                var item_49 = bd.Estimacion_Berries.Where(x => x.Semanas == 49 && x.Cultivo == 1 && x.Concepto == concepto).FirstOrDefault();
                item_49.Cantidad = _49;
                item_49.Fecha = DateTime.Now;

                var item_50 = bd.Estimacion_Berries.Where(x => x.Semanas == 50 && x.Cultivo == 1 && x.Concepto == concepto).FirstOrDefault();
                item_50.Cantidad = _50;
                item_50.Fecha = DateTime.Now;

                var item_51 = bd.Estimacion_Berries.Where(x => x.Semanas == 51 && x.Cultivo == 1 && x.Concepto == concepto).FirstOrDefault();
                item_51.Cantidad = _51;
                item_51.Fecha = DateTime.Now;

                var item_52 = bd.Estimacion_Berries.Where(x => x.Semanas == 52 && x.Cultivo == 1 && x.Concepto == concepto).FirstOrDefault();
                item_52.Cantidad = _52;
                item_52.Fecha = DateTime.Now;

                var item_1 = bd.Estimacion_Berries.Where(x => x.Semanas == 1 && x.Cultivo == 1 && x.Concepto == concepto).FirstOrDefault();
                item_1.Cantidad = _1;
                item_1.Fecha = DateTime.Now;

                var item_2 = bd.Estimacion_Berries.Where(x => x.Semanas == 2 && x.Cultivo == 1 && x.Concepto == concepto).FirstOrDefault();
                item_2.Cantidad = _2;
                item_2.Fecha = DateTime.Now;

                var item_3 = bd.Estimacion_Berries.Where(x => x.Semanas == 3 && x.Cultivo == 1 && x.Concepto == concepto).FirstOrDefault();
                item_3.Cantidad = _3;
                item_3.Fecha = DateTime.Now;

                var item_4 = bd.Estimacion_Berries.Where(x => x.Semanas == 4 && x.Cultivo == 1 && x.Concepto == concepto).FirstOrDefault();
                item_4.Cantidad = _4;
                item_4.Fecha = DateTime.Now;

                var item_5 = bd.Estimacion_Berries.Where(x => x.Semanas == 5 && x.Cultivo == 1 && x.Concepto == concepto).FirstOrDefault();
                item_5.Cantidad = _5;
                item_5.Fecha = DateTime.Now;

                var item_6 = bd.Estimacion_Berries.Where(x => x.Semanas == 6 && x.Cultivo == 1 && x.Concepto == concepto).FirstOrDefault();
                item_6.Cantidad = _6;
                item_6.Fecha = DateTime.Now;

                var item_7 = bd.Estimacion_Berries.Where(x => x.Semanas == 7 && x.Cultivo == 1 && x.Concepto == concepto).FirstOrDefault();
                item_7.Cantidad = _7;
                item_7.Fecha = DateTime.Now;

                var item_8 = bd.Estimacion_Berries.Where(x => x.Semanas == 8 && x.Cultivo == 1 && x.Concepto == concepto).FirstOrDefault();
                item_8.Cantidad = _8;
                item_8.Fecha = DateTime.Now;

                var item_9 = bd.Estimacion_Berries.Where(x => x.Semanas == 9 && x.Cultivo == 1 && x.Concepto == concepto).FirstOrDefault();
                item_9.Cantidad = _9;
                item_9.Fecha = DateTime.Now;

                var item_10 = bd.Estimacion_Berries.Where(x => x.Semanas == 10 && x.Cultivo == 1 && x.Concepto == concepto).FirstOrDefault();
                item_10.Cantidad = _10;
                item_10.Fecha = DateTime.Now;

                var item_11 = bd.Estimacion_Berries.Where(x => x.Semanas == 11 && x.Cultivo == 1 && x.Concepto == concepto).FirstOrDefault();
                item_11.Cantidad = _11;
                item_11.Fecha = DateTime.Now;

                var item_12 = bd.Estimacion_Berries.Where(x => x.Semanas == 12 && x.Cultivo == 1 && x.Concepto == concepto).FirstOrDefault();
                item_12.Cantidad = _12;
                item_12.Fecha = DateTime.Now;

                var item_13 = bd.Estimacion_Berries.Where(x => x.Semanas == 13 && x.Cultivo == 1 && x.Concepto == concepto).FirstOrDefault();
                item_13.Cantidad = _13;
                item_13.Fecha = DateTime.Now;

                var item_14 = bd.Estimacion_Berries.Where(x => x.Semanas == 14 && x.Cultivo == 1 && x.Concepto == concepto).FirstOrDefault();
                item_14.Cantidad = _14;
                item_14.Fecha = DateTime.Now;

                var item_15 = bd.Estimacion_Berries.Where(x => x.Semanas == 15 && x.Cultivo == 1 && x.Concepto == concepto).FirstOrDefault();
                item_15.Cantidad = _15;
                item_15.Fecha = DateTime.Now;

                var item_16 = bd.Estimacion_Berries.Where(x => x.Semanas == 16 && x.Cultivo == 1 && x.Concepto == concepto).FirstOrDefault();
                item_16.Cantidad = _16;
                item_16.Fecha = DateTime.Now;

                var item_17 = bd.Estimacion_Berries.Where(x => x.Semanas == 17 && x.Cultivo == 1 && x.Concepto == concepto).FirstOrDefault();
                item_17.Cantidad = _17;
                item_17.Fecha = DateTime.Now;

                var item_18 = bd.Estimacion_Berries.Where(x => x.Semanas == 18 && x.Cultivo == 1 && x.Concepto == concepto).FirstOrDefault();
                item_18.Cantidad = _18;
                item_18.Fecha = DateTime.Now;

                var item_19 = bd.Estimacion_Berries.Where(x => x.Semanas == 19 && x.Cultivo == 1 && x.Concepto == concepto).FirstOrDefault();
                item_19.Cantidad = _19;
                item_19.Fecha = DateTime.Now;

                var item_20 = bd.Estimacion_Berries.Where(x => x.Semanas == 20 && x.Cultivo == 1 && x.Concepto == concepto).FirstOrDefault();
                item_20.Cantidad = _20;
                item_20.Fecha = DateTime.Now;

                var item_21 = bd.Estimacion_Berries.Where(x => x.Semanas == 21 && x.Cultivo == 1 && x.Concepto == concepto).FirstOrDefault();
                item_21.Cantidad = _21;
                item_21.Fecha = DateTime.Now;

                var item_22 = bd.Estimacion_Berries.Where(x => x.Semanas == 22 && x.Cultivo == 1 && x.Concepto == concepto).FirstOrDefault();
                item_22.Cantidad = _22;
                item_22.Fecha = DateTime.Now;

                var item_23 = bd.Estimacion_Berries.Where(x => x.Semanas == 23 && x.Cultivo == 1 && x.Concepto == concepto).FirstOrDefault();
                item_23.Cantidad = _23;
                item_23.Fecha = DateTime.Now;

                var item_24 = bd.Estimacion_Berries.Where(x => x.Semanas == 24 && x.Cultivo == 1 && x.Concepto == concepto).FirstOrDefault();
                item_24.Cantidad = _24;
                item_24.Fecha = DateTime.Now;

                var item_25 = bd.Estimacion_Berries.Where(x => x.Semanas == 25 && x.Cultivo == 1 && x.Concepto == concepto).FirstOrDefault();
                item_25.Cantidad = _25;
                item_25.Fecha = DateTime.Now;

                bd.SaveChanges();
                return true;
            }
            else
            {
                return false;
            }
        }

        //FRAMBUESA
        public ActionResult Frambuesa()
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

        public JsonResult RecepcionFrambuesa()
        {
            List<R_Berries> recepcion_real = bd.Database.SqlQuery<R_Berries>("SELECT 'TOTAL' as 'SEMANA',  round(isnull(V.[37],0),0) as _37,round(isnull(V.[38],0),0) as _38,round(isnull(V.[39],0),0) as _39,round(isnull(V.[40],0),0) as _40,round(isnull(V.[41],0),0) as _41,round(isnull(V.[42],0),0) as _42,round(isnull(V.[43],0),0) as _43,round(isnull(V.[44],0),0) as _44,round(isnull(V.[45],0),0) as _45,round(isnull(V.[46],0),0) as _46,round(isnull(V.[47],0),0) as _47,round(isnull(V.[48],0),0) as _48,round(isnull(V.[49],0),0) as _49,round(isnull(V.[50],0),0) as _50,round(isnull(V.[51],0),0) as _51,round(isnull(V.[52],0),0) as _52,round(isnull(V.[1],0),0) as _1,round(isnull(V.[2],0),0) as _2,round(isnull(V.[3],0),0) as _3,round(isnull(V.[4],0),0) as _4,round(isnull(V.[5],0),0) as _5,round(isnull(V.[6],0),0) as _6,round(isnull(V.[7],0),0) as _7,round(isnull(V.[8],0),0) as _8,round(isnull(V.[9],0),0) as _9,round(isnull(V.[10],0),0) as _10,round(isnull(V.[11],0),0) as _11,round(isnull(V.[12],0),0) as _12,round(isnull(V.[13],0),0) as _13,round(isnull(V.[14],0),0) as _14,round(isnull(V.[15],0),0) as _15,round(isnull(V.[16],0),0) as _16,round(isnull(V.[17],0),0) as _17,round(isnull(V.[18],0),0) as _18,round(isnull(V.[19],0),0) as _19,round(isnull(V.[20],0),0) as _20,round(isnull(V.[21],0),0) as _21,round(isnull(V.[22],0),0) as _22,round(isnull(V.[23],0),0) as _23,round(isnull(V.[24],0),0) as _24,round(isnull(V.[25],0),0) as _25 " +
                "FROM(Select * from(SELECT Semana, SUM(Convertidas) AS Convertidas FROM SEasonsun1213..UV_ProdRecepcion where CodEstatus <> 'C' and temporada = (select temporada from CatSemanas where GETDATE() between Inicio and Fin) and DescProducto = 'FRAMBUESA' GROUP BY Semana Union All " +
                "SELECT Semana, SUM(Convertidas) AS Convertidas FROM SeasonPlan..UV_ProdRecepcion where CodEstatus <> 'C' and temporada = (select temporada from CatSemanas where GETDATE() between Inicio and Fin) and DescProducto = 'FRAMBUESA' GROUP BY Semana" +
                ")V PIVOT(SUM(Convertidas) FOR Semana in ([37],[38],[39],[40],[41],[42],[43],[44],[45],[46],[47],[48],[49],[50],[51],[52],[1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],[17],[18],[19],[20],[21],[22],[23],[24],[25]))X)V ").ToList();
            return Json(recepcion_real, JsonRequestBehavior.AllowGet);
        }

        public JsonResult EstimacionSFrambuesa()
        {
            List<E_Berries> estimacion = bd.Database.SqlQuery<E_Berries>("Select V.Sem AS 'SEMANA', round(isnull(V.[37],0),0) as _37,round(isnull(V.[38],0),0) as _38,round(isnull(V.[39],0),0) as _39,round(isnull(V.[40],0),0) as _40,round(isnull(V.[41],0),0) as _41,round(isnull(V.[42],0),0) as _42,round(isnull(V.[43],0),0) as _43,round(isnull(V.[44],0),0) as _44,round(isnull(V.[45],0),0) as _45,round(isnull(V.[46],0),0) as _46,round(isnull(V.[47],0),0) as _47,round(isnull(V.[48],0),0) as _48,round(isnull(V.[49],0),0) as _49,round(isnull(V.[50],0),0) as _50,round(isnull(V.[51],0),0) as _51,round(isnull(V.[52],0),0) as _52,round(isnull(V.[1],0),0) as _1,round(isnull(V.[2],0),0) as _2,round(isnull(V.[3],0),0) as _3,round(isnull(V.[4],0),0) as _4,round(isnull(V.[5],0),0) as _5,round(isnull(V.[6],0),0) as _6,round(isnull(V.[7],0),0) as _7,round(isnull(V.[8],0),0) as _8,round(isnull(V.[9],0),0) as _9,round(isnull(V.[10],0),0) as _10,round(isnull(V.[11],0),0) as _11,round(isnull(V.[12],0),0) as _12,round(isnull(V.[13],0),0) as _13,round(isnull(V.[14],0),0) as _14,round(isnull(V.[15],0),0) as _15,round(isnull(V.[16],0),0) as _16,round(isnull(V.[17],0),0) as _17,round(isnull(V.[18],0),0) as _18,round(isnull(V.[19],0),0) as _19,round(isnull(V.[20],0),0) as _20,round(isnull(V.[21],0),0) as _21,round(isnull(V.[22],0),0) as _22,round(isnull(V.[23],0),0) as _23,round(isnull(V.[24],0),0) as _24,round(isnull(V.[25],0),0) as _25 " +
                "from(Select * from(SELECT S.Semana AS Sem, S.Inicio, E.Semanas, E.Cantidad FROM Estimacion_Berries E left join CatSemanas S on E.Temporada = S.Temporada AND E.Semana = S.Semana " +
                "where E.Temporada = (select temporada from CatSemanas where GETDATE() between Inicio and Fin) and E.Concepto = 'PROYECCION SEMANAL' and E.Cultivo = 2)V " +
                "PIVOT(SUM(Cantidad) For Semanas in ([37],[38],[39],[40],[41],[42],[43],[44],[45],[46],[47],[48],[49],[50],[51],[52],[1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],[17],[18],[19],[20],[21],[22],[23],[24],[25]))X)V ORDER BY V.Inicio").ToList();
            return Json(estimacion, JsonRequestBehavior.AllowGet);
        }

        public JsonResult EstimacionSFrambuesaP()
        {
            List<E_BerriesP> estimacion = bd.Database.SqlQuery<E_BerriesP>("select V.Sem AS 'SEMANA',(case when V._37='0%' then '' else V._37 end) as _37,(case when V._38='0%' then '' else V._38 end) as _38,(case when V._39='0%' then '' else V._39 end) as _39,(case when V._40='0%' then '' else V._40 end) as _40,(case when V._41='0%' then '' else V._41 end) as _41,(case when V._42='0%' then '' else V._42 end) as _42,(case when V._43='0%' then '' else V._43 end) as _43,(case when V._44='0%' then '' else V._44 end) as _44,(case when V._45='0%' then '' else V._45 end) as _45,(case when V._46='0%' then '' else V._46 end) as _46,(case when V._47='0%' then '' else V._47 end) as _47,(case when V._48='0%' then '' else V._48 end) as _48,(case when V._49='0%' then '' else V._49 end) as _49,(case when V._50='0%' then '' else V._50 end) as _50,(case when V._51='0%' then '' else V._51 end) as _51,(case when V._52='0%' then '' else V._52 end) as _52,(case when V._1='0%' then '' else V._1 end) as _1,(case when V._2='0%' then '' else V._2 end) as _2,(case when V._3='0%' then '' else V._3 end) as _3,(case when V._4='0%' then '' else V._4 end) as _4,(case when V._5='0%' then '' else V._5 end) as _5,(case when V._6='0%' then '' else V._6 end) as _6,(case when V._7='0%' then '' else V._7 end) as _7,(case when V._8='0%' then '' else V._8 end) as _8,(case when V._9='0%' then '' else V._9 end) as _9,(case when V._10='0%' then '' else V._10 end) as _10,(case when V._11='0%' then '' else V._11 end) as _11,(case when V._12='0%' then '' else V._12 end) as _12,(case when V._13='0%' then '' else V._13 end) as _13,(case when V._14='0%' then '' else V._14 end) as _14,(case when V._15='0%' then '' else V._15 end) as _15,(case when V._16='0%' then '' else V._16 end) as _16,(case when V._17='0%' then '' else V._17 end) as _17,(case when V._18='0%' then '' else V._18 end) as _18,(case when V._19='0%' then '' else V._19 end) as _19,(case when V._20='0%' then '' else V._20 end) as _20,(case when V._21='0%' then '' else V._21 end) as _21,(case when V._22='0%' then '' else V._22 end) as _22,(case when V._23='0%' then '' else V._23 end) as _23,(case when V._24='0%' then '' else V._24 end) as _24,(case when V._25='0%' then '' else V._25 end) as _25 from(select V.Inicio, V.Sem, cast(V._35 as varchar)+'%' as _35,cast(V._36 as varchar)+'%' as _36,cast(V._37 as varchar)+'%' as _37,cast(V._38 as varchar)+'%' as _38,cast(V._39 as varchar)+'%' as _39,cast(V._40 as varchar)+'%' as _40,cast(V._41 as varchar)+'%' as _41,cast(V._42 as varchar)+'%' as _42,cast(V._43 as varchar)+'%' as _43,cast(V._44 as varchar)+'%' as _44,cast(V._45 as varchar)+'%' as _45,cast(V._46 as varchar)+'%' as _46,cast(V._47 as varchar)+'%' as _47,cast(V._48 as varchar)+'%' as _48,cast(V._49 as varchar)+'%' as _49,cast(V._50 as varchar)+'%' as _50,cast(V._51 as varchar)+'%' as _51,cast(V._52 as varchar)+'%' as _52,cast(V._1 as varchar)+'%' as _1,cast(V._2 as varchar)+'%' _2,cast(V._3 as varchar)+'%' as _3,cast(V._4 as varchar)+'%' as _4,cast(V._5 as varchar)+'%' as _5,cast(V._6 as varchar)+'%' as _6,cast(V._7 as varchar)+'%' as _7,cast(V._8 as varchar)+'%' as _8,cast(V._9 as varchar)+'%' as _9,cast(V._10 as varchar)+'%' as _10,cast(V._11 as varchar)+'%' as _11,cast(V._12 as varchar)+'%' as _12,cast(V._13 as varchar)+'%' as _13,cast(V._14 as varchar)+'%' as _14,cast(V._15 as varchar)+'%' as _15,cast(V._16 as varchar)+'%' as _16,cast(V._17 as varchar)+'%' as _17,cast(V._18 as varchar)+'%' as _18,cast(V._19 as varchar)+'%' as _19,cast(V._20 as varchar)+'%'as _20,cast(V._21 as varchar)+'%' as _21,cast(V._22 as varchar)+'%' as _22,cast(V._23 as varchar)+'%' as _23,cast(V._24 as varchar)+'%' as _24,cast(V._25 as varchar)+'%' as _25 from (SELECT V.Inicio, V.Sem,(case when A._35=0 then '' else case when V._35=0 then '' else round((A._35/V._35)*100,0) end end) as _35,(case when A._36=0 then 0 else case when V._36=0 then 0 else round((A._36/V._36)*100,0) end end) as _36,(case when A._37=0 then 0 else case when V._37=0 then 0 else round((A._37/V._37)*100,0) end end) as _37,(case when A._38=0 then 0 else case when V._38=0 then 0 else round((A._38/V._38)*100,0) end end) as _38,(case when A._39=0 then 0 else case when V._39=0 then 0 else round((A._39/V._39)*100,0) end end) as _39,(case when A._40=0 then 0 else case when V._40=0 then 0 else round((A._40/V._40)*100,0) end end) as _40,(case when A._41=0 then 0 else case when V._41=0 then 0 else round((A._41/V._41)*100,0) end end) as _41,(case when A._42=0 then 0 else case when V._42=0 then 0 else round((A._42/V._42)*100,0) end end) as _42,(case when A._43=0 then 0 else case when V._43=0 then 0 else round((A._43/V._43)*100,0) end end) as _43,(case when A._44=0 then 0 else case when V._44=0 then 0 else round((A._44/V._44)*100,0) end end) as _44,(case when A._45=0 then 0 else case when V._45=0 then 0 else round((A._45/V._45)*100,0) end end) as _45,(case when A._46=0 then 0 else case when V._46=0 then 0 else round((A._46/V._46)*100,0) end end) as _46,(case when A._47=0 then 0 else case when V._47=0 then 0 else round((A._47/V._47)*100,0) end end) as _47,(case when A._48=0 then 0 else case when V._48=0 then 0 else round((A._48/V._48)*100,0) end end) as _48,(case when A._49=0 then 0 else case when V._49=0 then 0 else round((A._49/V._49)*100,0) end end) as _49,(case when A._50=0 then 0 else case when V._50=0 then 0 else round((A._50/V._50)*100,0) end end) as _50,(case when A._51=0 then 0 else case when V._51=0 then 0 else round((A._51/V._51)*100,0) end end) as _51,(case when A._52=0 then 0 else case when V._52=0 then 0 else round((A._52/V._52)*100,0) end end) as _52,(case when A._1=0 then 0 else case when V._1=0 then 0 else round((A._1/V._1)*100,0) end end) as _1,(case when A._2=0 then 0 else case when V._2=0 then 0 else round((A._2/V._2)*100,0) end end) as _2,(case when A._3=0 then 0 else case when V._3=0 then 0 else round((A._3/V._3)*100,0) end end) as _3,(case when A._4=0 then 0 else case when V._4=0 then 0 else round((A._4/V._4)*100,0) end end) as _4,(case when A._5=0 then 0 else case when V._5=0 then 0 else round((A._5/V._5)*100,0) end end) as _5,(case when A._6=0 then 0 else case when V._6=0 then 0 else round((A._6/V._6)*100,0) end end) as _6,(case when A._7=0 then 0 else case when V._7=0 then 0 else round((A._7/V._7)*100,0) end end) as _7,(case when A._8=0 then 0 else case when V._8=0 then 0 else round((A._8/V._8)*100,0) end end) as _8,(case when A._9=0 then 0 else case when V._9=0 then 0 else round((A._9/V._9)*100,0) end end) as _9,(case when A._10=0 then 0 else case when V._10=0 then 0 else round((A._10/V._10)*100,0) end end) as _10,(case when A._11=0 then 0 else case when V._11=0 then 0 else round((A._11/V._11)*100,0) end end) as _11,(case when A._12=0 then 0 else case when V._12=0 then 0 else round((A._12/V._12)*100,0) end end) as _12,(case when A._13=0 then 0 else case when V._13=0 then 0 else round((A._13/V._13)*100,0) end end) as _13,(case when A._14=0 then 0 else case when V._14=0 then 0 else round((A._14/V._14)*100,0) end end) as _14,(case when A._15=0 then 0 else case when V._15=0 then 0 else round((A._15/V._15)*100,0) end end) as _15,(case when A._16=0 then 0 else case when V._16=0 then 0 else round((A._16/V._16)*100,0) end end) as _16,(case when A._17=0 then 0 else case when V._17=0 then 0 else round((A._17/V._17)*100,0) end end) as _17,(case when A._18=0 then 0 else case when V._18=0 then 0 else round((A._18/V._18)*100,0) end end) as _18,(case when A._19=0 then 0 else case when V._19=0 then 0 else round((A._19/V._19)*100,0) end end) as _19,(case when A._20=0 then 0 else case when V._20=0 then 0 else round((A._20/V._20)*100,0) end end) as _20,(case when A._21=0 then 0 else case when V._21=0 then 0 else round((A._21/V._21)*100,0) end end) as _21,(case when A._22=0 then 0 else case when V._22=0 then 0 else round((A._22/V._22)*100,0) end end) as _22,(case when A._23=0 then 0 else case when V._23=0 then 0 else round((A._23/V._23)*100,0) end end) as _23,(case when A._24=0 then 0 else case when V._24=0 then 0 else round((A._24/V._24)*100,0) end end) as _24,(case when A._25=0 then 0 else case when V._25=0 then 0 else round((A._25/V._25)*100,0) end end) as _25 FROM(SELECT V.Temporada, isnull(V.[35],0) as _35,isnull(V.[36],0) as _36,isnull(V.[37],0) as _37,isnull(V.[38],0) as _38,isnull(V.[39],0) as _39,isnull(V.[40],0) as _40,isnull(V.[41],0) as _41,isnull(V.[42],0) as _42,isnull(V.[43],0) as _43,isnull(V.[44],0) as _44,isnull(V.[45],0) as _45,isnull(V.[46],0) as _46,isnull(V.[47],0) as _47,isnull(V.[48],0) as _48,isnull(V.[49],0) as _49,isnull(V.[50],0) as _50,isnull(V.[51],0) as _51,isnull(V.[52],0) as _52,isnull(V.[1],0) as _1,isnull(V.[2],0) as _2,isnull(V.[3],0) as _3,isnull(V.[4],0) as _4,isnull(V.[5],0) as _5,isnull(V.[6],0) as _6,isnull(V.[7],0) as _7,isnull(V.[8],0) as _8,isnull(V.[9],0) as _9,isnull(V.[10],0) as _10,isnull(V.[11],0) as _11,isnull(V.[12],0) as _12,isnull(V.[13],0) as _13,isnull(V.[14],0) as _14,isnull(V.[15],0) as _15,isnull(V.[16],0) as _16,isnull(V.[17],0) as _17,isnull(V.[18],0) as _18,isnull(V.[19],0) as _19,isnull(V.[20],0) as _20,isnull(V.[21],0) as _21,isnull(V.[22],0) as _22,isnull(V.[23],0) as _23,isnull(V.[24],0) as _24,isnull(V.[25],0) as _25 FROM(Select * from(SELECT Temporada, Semana, SUM(Convertidas) AS Convertidas FROM SEasonsun1213..UV_ProdRecepcion where CodEstatus<> 'C' and temporada = (select temporada from CatSemanas where GETDATE() between Inicio and Fin) and " +
           "DescProducto = 'FRAMBUESA' GROUP BY Temporada, Semana Union All SELECT Temporada, Semana, SUM(Convertidas) AS Convertidas FROM SeasonPlan..UV_ProdRecepcion where CodEstatus <> 'C' and temporada = (select temporada from CatSemanas where GETDATE() between Inicio and Fin) and " +
           "DescProducto = 'FRAMBUESA' GROUP BY Temporada, Semana)V PIVOT(SUM(Convertidas) FOR Semana in ([35],[36],[37],[38],[39],[40],[41],[42],[43],[44],[45],[46],[47],[48],[49],[50],[51],[52],[1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],[17],[18],[19],[20],[21],[22],[23],[24],[25]))X)V)A left join(Select V.Inicio, V.Temporada, V.Sem, isnull(V.[35], 0) as _35, isnull(V.[36], 0) as _36, isnull(V.[37], 0) as _37, isnull(V.[38], 0) as _38, isnull(V.[39], 0) as _39, isnull(V.[40], 0) as _40, isnull(V.[41], 0) as _41, isnull(V.[42], 0) as _42, isnull(V.[43], 0) as _43, isnull(V.[44], 0) as _44, isnull(V.[45], 0) as _45, isnull(V.[46], 0) as _46, isnull(V.[47], 0) as _47, isnull(V.[48], 0) as _48, isnull(V.[49], 0) as _49, isnull(V.[50], 0) as _50, isnull(V.[51], 0) as _51, isnull(V.[52], 0) as _52, isnull(V.[1], 0) as _1, isnull(V.[2], 0) as _2, isnull(V.[3], 0) as _3, isnull(V.[4], 0) as _4, isnull(V.[5], 0) as _5, isnull(V.[6], 0) as _6, isnull(V.[7], 0) as _7, isnull(V.[8], 0) as _8, isnull(V.[9], 0) as _9, isnull(V.[10], 0) as _10, isnull(V.[11], 0) as _11, isnull(V.[12], 0) as _12, isnull(V.[13], 0) as _13, isnull(V.[14], 0) as _14, isnull(V.[15], 0) as _15, isnull(V.[16], 0) as _16, isnull(V.[17], 0) as _17, isnull(V.[18], 0) as _18, isnull(V.[19], 0) as _19, isnull(V.[20], 0) as _20, isnull(V.[21], 0) as _21, isnull(V.[22], 0) as _22, isnull(V.[23], 0) as _23, isnull(V.[24], 0) as _24, isnull(V.[25], 0) as _25 from(select * from(SELECT S.Semana AS Sem, S.Inicio, E.Semanas, E.Cantidad, E.Temporada FROM Estimacion_Berries E left join CatSemanas S on E.Temporada = S.Temporada AND E.Semana = S.Semana where E.Temporada = (select temporada from CatSemanas where GETDATE() between Inicio and Fin) and " +
           "E.Concepto = 'PROYECCION SEMANAL' and E.Cultivo = 2)V PIVOT(SUM(Cantidad) For Semanas in ([35],[36],[37],[38],[39],[40],[41],[42],[43],[44],[45],[46],[47],[48],[49],[50],[51],[52],[1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],[17],[18],[19],[20],[21],[22],[23],[24],[25]))X)V)V on A.Temporada = V.Temporada)V)V ORDER BY V.Inicio").ToList();
            return Json(estimacion, JsonRequestBehavior.AllowGet);
        }

        [HttpPost]
        public ActionResult Frambuesa(int[] array)
        {
            try
            {
                if (Update_Frambuesa(array))
                {
                    dtaEjecucionTarea = Json(new
                    {
                        rstProceso = "true",
                        MessageGestion = "Cambios guardados con éxito"
                    });
                }
                else
                {
                    dtaEjecucionTarea = Json(new
                    {
                        rstProceso = "false",
                        MessageGestion = "Error, algo salió mal, intente de nuevo"
                    });
                    TempData["sms"] = "Error, algo salió mal, intente de nuevo";
                    ViewBag.error = TempData["sms"].ToString();
                }
            }
            catch (Exception e)
            {
                e.ToString();
            }

            if (array[0] == 25)
            {
                TempData["sms"] = "Cambios guardados con éxito";
                ViewBag.sms = TempData["sms"].ToString();
            }

            return dtaEjecucionTarea;
        }

        public bool Update_Frambuesa(int[] array)
        {
            if (array != null)
            {
                var sem = array[0];
                var _37 = array[1];
                var _38 = array[2];
                var _39 = array[3];
                var _40 = array[4];
                var _41 = array[5];
                var _42 = array[6];
                var _43 = array[7];
                var _44 = array[8];
                var _45 = array[9];
                var _46 = array[10];
                var _47 = array[11];
                var _48 = array[12];
                var _49 = array[13];
                var _50 = array[14];
                var _51 = array[15];
                var _52 = array[16];
                var _1 = array[17];
                var _2 = array[18];
                var _3 = array[19];
                var _4 = array[20];
                var _5 = array[21];
                var _6 = array[22];
                var _7 = array[23];
                var _8 = array[24];
                var _9 = array[25];
                var _10 = array[26];
                var _11 = array[27];
                var _12 = array[28];
                var _13 = array[29];
                var _14 = array[30];
                var _15 = array[31];
                var _16 = array[32];
                var _17 = array[33];
                var _18 = array[34];
                var _19 = array[35];
                var _20 = array[36];
                var _21 = array[37];
                var _22 = array[38];
                var _23 = array[39];
                var _24 = array[40];
                var _25 = array[41];

                var item_37 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 37 && x.Cultivo == 2 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_37.Cantidad = _37;
                item_37.Fecha = DateTime.Now;

                var item_38 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 38 && x.Cultivo == 2 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_38.Cantidad = _38;
                item_38.Fecha = DateTime.Now;

                var item_39 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 39 && x.Cultivo == 2 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_39.Cantidad = _39;
                item_39.Fecha = DateTime.Now;

                var item_40 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 40 && x.Cultivo == 2 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_40.Cantidad = _40;
                item_40.Fecha = DateTime.Now;

                var item_41 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 41 && x.Cultivo == 2 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_41.Cantidad = _41;
                item_41.Fecha = DateTime.Now;

                var item_42 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 42 && x.Cultivo == 2 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_42.Cantidad = _42;
                item_42.Fecha = DateTime.Now;

                var item_43 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 43 && x.Cultivo == 2 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_43.Cantidad = _43;
                item_43.Fecha = DateTime.Now;

                var item_44 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 44 && x.Cultivo == 2 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_44.Cantidad = _44;
                item_44.Fecha = DateTime.Now;

                var item_45 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 45 && x.Cultivo == 2 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_45.Cantidad = _45;
                item_45.Fecha = DateTime.Now;

                var item_46 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 46 && x.Cultivo == 2 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_46.Cantidad = _46;
                item_46.Fecha = DateTime.Now;

                var item_47 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 47 && x.Cultivo == 2 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_47.Cantidad = _47;
                item_47.Fecha = DateTime.Now;

                var item_48 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 48 && x.Cultivo == 2 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_48.Cantidad = _48;
                item_48.Fecha = DateTime.Now;

                var item_49 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 49 && x.Cultivo == 2 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_49.Cantidad = _49;
                item_49.Fecha = DateTime.Now;

                var item_50 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 50 && x.Cultivo == 2 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_50.Cantidad = _50;
                item_50.Fecha = DateTime.Now;

                var item_51 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 51 && x.Cultivo == 2 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_51.Cantidad = _51;
                item_51.Fecha = DateTime.Now;

                var item_52 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 52 && x.Cultivo == 2 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_52.Cantidad = _52;
                item_52.Fecha = DateTime.Now;

                var item_1 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 1 && x.Cultivo == 2 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_1.Cantidad = _1;
                item_1.Fecha = DateTime.Now;

                var item_2 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 2 && x.Cultivo == 2 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_2.Cantidad = _2;
                item_2.Fecha = DateTime.Now;

                var item_3 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 3 && x.Cultivo == 2 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_3.Cantidad = _3;
                item_3.Fecha = DateTime.Now;

                var item_4 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 4 && x.Cultivo == 2 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_4.Cantidad = _4;
                item_4.Fecha = DateTime.Now;

                var item_5 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 5 && x.Cultivo == 2 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_5.Cantidad = _5;
                item_5.Fecha = DateTime.Now;

                var item_6 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 6 && x.Cultivo == 2 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_6.Cantidad = _6;
                item_6.Fecha = DateTime.Now;

                var item_7 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 7 && x.Cultivo == 2 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_7.Cantidad = _7;
                item_7.Fecha = DateTime.Now;

                var item_8 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 8 && x.Cultivo == 2 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_8.Cantidad = _8;
                item_8.Fecha = DateTime.Now;

                var item_9 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 9 && x.Cultivo == 2 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_9.Cantidad = _9;
                item_9.Fecha = DateTime.Now;

                var item_10 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 10 && x.Cultivo == 2 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_10.Cantidad = _10;
                item_10.Fecha = DateTime.Now;

                var item_11 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 11 && x.Cultivo == 2 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_11.Cantidad = _11;
                item_11.Fecha = DateTime.Now;

                var item_12 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 12 && x.Cultivo == 2 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_12.Cantidad = _12;
                item_12.Fecha = DateTime.Now;

                var item_13 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 13 && x.Cultivo == 2 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_13.Cantidad = _13;
                item_13.Fecha = DateTime.Now;

                var item_14 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 14 && x.Cultivo == 2 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_14.Cantidad = _14;
                item_14.Fecha = DateTime.Now;

                var item_15 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 15 && x.Cultivo == 2 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_15.Cantidad = _15;
                item_15.Fecha = DateTime.Now;

                var item_16 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 16 && x.Cultivo == 2 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_16.Cantidad = _16;
                item_16.Fecha = DateTime.Now;

                var item_17 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 17 && x.Cultivo == 2 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_17.Cantidad = _17;
                item_17.Fecha = DateTime.Now;

                var item_18 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 18 && x.Cultivo == 2 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_18.Cantidad = _18;
                item_18.Fecha = DateTime.Now;

                var item_19 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 19 && x.Cultivo == 2 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_19.Cantidad = _19;
                item_19.Fecha = DateTime.Now;

                var item_20 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 20 && x.Cultivo == 2 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_20.Cantidad = _20;
                item_20.Fecha = DateTime.Now;

                var item_21 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 21 && x.Cultivo == 2 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_21.Cantidad = _21;
                item_21.Fecha = DateTime.Now;

                var item_22 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 22 && x.Cultivo == 2 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_22.Cantidad = _22;
                item_22.Fecha = DateTime.Now;

                var item_23 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 23 && x.Cultivo == 2 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_23.Cantidad = _23;
                item_23.Fecha = DateTime.Now;

                var item_24 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 24 && x.Cultivo == 2 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_24.Cantidad = _24;
                item_24.Fecha = DateTime.Now;

                var item_25 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 25 && x.Cultivo == 2 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_25.Cantidad = _25;
                item_25.Fecha = DateTime.Now;

                bd.SaveChanges();
                return true;
            }
            else
            {
                return false;
            }
        }

        //ARANDANO
        public ActionResult Arandano()
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

        public JsonResult RecepcionArandano()
        {
            List<R_Berries> recepcion_real = bd.Database.SqlQuery<R_Berries>("SELECT 'TOTAL' as 'SEMANA', round(isnull(V.[35],0),0) as _35,round(isnull(V.[36],0),0) as _36,round(isnull(V.[37],0),0) as _37,round(isnull(V.[38],0),0) as _38,round(isnull(V.[39],0),0) as _39,round(isnull(V.[40],0),0) as _40,round(isnull(V.[41],0),0) as _41,round(isnull(V.[42],0),0) as _42,round(isnull(V.[43],0),0) as _43,round(isnull(V.[44],0),0) as _44,round(isnull(V.[45],0),0) as _45,round(isnull(V.[46],0),0) as _46,round(isnull(V.[47],0),0) as _47,round(isnull(V.[48],0),0) as _48,round(isnull(V.[49],0),0) as _49,round(isnull(V.[50],0),0) as _50,round(isnull(V.[51],0),0) as _51,round(isnull(V.[52],0),0) as _52,round(isnull(V.[1],0),0) as _1,round(isnull(V.[2],0),0) as _2,round(isnull(V.[3],0),0) as _3,round(isnull(V.[4],0),0) as _4,round(isnull(V.[5],0),0) as _5,round(isnull(V.[6],0),0) as _6,round(isnull(V.[7],0),0) as _7,round(isnull(V.[8],0),0) as _8,round(isnull(V.[9],0),0) as _9,round(isnull(V.[10],0),0) as _10,round(isnull(V.[11],0),0) as _11,round(isnull(V.[12],0),0) as _12,round(isnull(V.[13],0),0) as _13,round(isnull(V.[14],0),0) as _14,round(isnull(V.[15],0),0) as _15,round(isnull(V.[16],0),0) as _16,round(isnull(V.[17],0),0) as _17,round(isnull(V.[18],0),0) as _18,round(isnull(V.[19],0),0) as _19,round(isnull(V.[20],0),0) as _20,round(isnull(V.[21],0),0) as _21,round(isnull(V.[22],0),0) as _22,round(isnull(V.[23],0),0) as _23,round(isnull(V.[24],0),0) as _24,round(isnull(V.[25],0),0) as _25 " +
                "FROM(Select * from(SELECT Semana, SUM(Convertidas) AS Convertidas FROM SEasonsun1213..UV_ProdRecepcion where CodEstatus <> 'C' and temporada = (select temporada from CatSemanas where GETDATE() between Inicio and Fin) and DescProducto = 'ARANDANO' GROUP BY Semana " +
                "Union All SELECT Semana, SUM(Convertidas) AS Convertidas FROM SeasonPlan..UV_ProdRecepcion where CodEstatus <> 'C' and temporada = (select temporada from CatSemanas where GETDATE() between Inicio and Fin) and DescProducto = 'ARANDANO' GROUP BY Semana)V PIVOT(SUM(Convertidas) FOR Semana in ([35],[36],[37],[38],[39],[40],[41],[42],[43],[44],[45],[46],[47],[48],[49],[50],[51],[52],[1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],[17],[18],[19],[20],[21],[22],[23],[24],[25]))X)V").ToList();
            return Json(recepcion_real, JsonRequestBehavior.AllowGet);
        }

        public JsonResult EuropaA()
        {
            List<E_Berries> estimacion = bd.Database.SqlQuery<E_Berries>("Select round(isnull(V.[35],0),0) as _35,round(isnull(V.[36],0),0) as _36, round(isnull(V.[37],0),0) as _37,round(isnull(V.[38],0),0) as _38,round(isnull(V.[39],0),0) as _39,round(isnull(V.[40],0),0) as _40,round(isnull(V.[41],0),0) as _41,round(isnull(V.[42],0),0) as _42,round(isnull(V.[43],0),0) as _43,round(isnull(V.[44],0),0) as _44,round(isnull(V.[45],0),0) as _45,round(isnull(V.[46],0),0) as _46,round(isnull(V.[47],0),0) as _47,round(isnull(V.[48],0),0) as _48,round(isnull(V.[49],0),0) as _49,round(isnull(V.[50],0),0) as _50,round(isnull(V.[51],0),0) as _51,round(isnull(V.[52],0),0) as _52,round(isnull(V.[1],0),0) as _1,round(isnull(V.[2],0),0) as _2,round(isnull(V.[3],0),0) as _3,round(isnull(V.[4],0),0) as _4,round(isnull(V.[5],0),0) as _5,round(isnull(V.[6],0),0) as _6,round(isnull(V.[7],0),0) as _7,round(isnull(V.[8],0),0) as _8,round(isnull(V.[9],0),0) as _9,round(isnull(V.[10],0),0) as _10,round(isnull(V.[11],0),0) as _11,round(isnull(V.[12],0),0) as _12,round(isnull(V.[13],0),0) as _13,round(isnull(V.[14],0),0) as _14,round(isnull(V.[15],0),0) as _15,round(isnull(V.[16],0),0) as _16,round(isnull(V.[17],0),0) as _17,round(isnull(V.[18],0),0) as _18,round(isnull(V.[19],0),0) as _19,round(isnull(V.[20],0),0) as _20,round(isnull(V.[21],0),0) as _21,round(isnull(V.[22],0),0) as _22,round(isnull(V.[23],0),0) as _23,round(isnull(V.[24],0),0) as _24,round(isnull(V.[25],0),0) as _25 " +
                "from(Select * from(SELECT E.Concepto, S.Inicio, E.Semanas, E.Cantidad FROM Estimacion_Berries E left join CatSemanas S on E.Temporada = S.Temporada AND E.Semana = S.Semana where E.Temporada = (select temporada from CatSemanas where GETDATE() between Inicio and Fin) and " +
                "E.Concepto = 'EUROPA' and E.Cultivo = 3)V PIVOT(SUM(Cantidad) For Semanas in ([35],[36],[37],[38],[39],[40],[41],[42],[43],[44],[45],[46],[47],[48],[49],[50],[51],[52],[1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],[17],[18],[19],[20],[21],[22],[23],[24],[25]))X)V ORDER BY V.Inicio").ToList();
            return Json(estimacion, JsonRequestBehavior.AllowGet);
        }

        public JsonResult EstimacionSArandano()
        {
            List<E_Berries> estimacion = bd.Database.SqlQuery<E_Berries>("Select V.Sem AS 'SEMANA', round(isnull(V.[35],0),0) as _35, round(isnull(V.[36],0),0) as _36, round(isnull(V.[37],0),0) as _37,round(isnull(V.[38],0),0) as _38,round(isnull(V.[39],0),0) as _39,round(isnull(V.[40],0),0) as _40,round(isnull(V.[41],0),0) as _41,round(isnull(V.[42],0),0) as _42,round(isnull(V.[43],0),0) as _43,round(isnull(V.[44],0),0) as _44,round(isnull(V.[45],0),0) as _45,round(isnull(V.[46],0),0) as _46,round(isnull(V.[47],0),0) as _47,round(isnull(V.[48],0),0) as _48,round(isnull(V.[49],0),0) as _49,round(isnull(V.[50],0),0) as _50,round(isnull(V.[51],0),0) as _51,round(isnull(V.[52],0),0) as _52,round(isnull(V.[1],0),0) as _1,round(isnull(V.[2],0),0) as _2,round(isnull(V.[3],0),0) as _3,round(isnull(V.[4],0),0) as _4,round(isnull(V.[5],0),0) as _5,round(isnull(V.[6],0),0) as _6,round(isnull(V.[7],0),0) as _7,round(isnull(V.[8],0),0) as _8,round(isnull(V.[9],0),0) as _9,round(isnull(V.[10],0),0) as _10,round(isnull(V.[11],0),0) as _11,round(isnull(V.[12],0),0) as _12,round(isnull(V.[13],0),0) as _13,round(isnull(V.[14],0),0) as _14,round(isnull(V.[15],0),0) as _15,round(isnull(V.[16],0),0) as _16,round(isnull(V.[17],0),0) as _17,round(isnull(V.[18],0),0) as _18,round(isnull(V.[19],0),0) as _19,round(isnull(V.[20],0),0) as _20,round(isnull(V.[21],0),0) as _21,round(isnull(V.[22],0),0) as _22,round(isnull(V.[23],0),0) as _23,round(isnull(V.[24],0),0) as _24,round(isnull(V.[25],0),0) as _25 " +
                "from(Select * from(SELECT S.Semana AS Sem, S.Inicio, E.Semanas, E.Cantidad FROM Estimacion_Berries E left join CatSemanas S on E.Temporada = S.Temporada AND E.Semana = S.Semana " +
                "where E.Temporada = (select temporada from CatSemanas where GETDATE() between Inicio and Fin) and E.Concepto = 'PROYECCION SEMANAL' and E.Cultivo = 3)V " +
                "PIVOT(SUM(Cantidad) For Semanas in ([35],[36],[37],[38],[39],[40],[41],[42],[43],[44],[45],[46],[47],[48],[49],[50],[51],[52],[1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],[17],[18],[19],[20],[21],[22],[23],[24],[25]))X)V ORDER BY V.Inicio").ToList();
            return Json(estimacion, JsonRequestBehavior.AllowGet);
        }

        public JsonResult EstimacionSArandanoP()
        {
            List<E_BerriesP> estimacion = bd.Database.SqlQuery<E_BerriesP>("select V.Sem AS 'SEMANA', (case when V._35='0%' then '' else V._35 end) as _35,(case when V._36='0%' then '' else V._36 end) as _36,(case when V._37='0%' then '' else V._37 end) as _37,(case when V._38='0%' then '' else V._38 end) as _38,(case when V._39='0%' then '' else V._39 end) as _39,(case when V._40='0%' then '' else V._40 end) as _40,(case when V._41='0%' then '' else V._41 end) as _41,(case when V._42='0%' then '' else V._42 end) as _42,(case when V._43='0%' then '' else V._43 end) as _43,(case when V._44='0%' then '' else V._44 end) as _44,(case when V._45='0%' then '' else V._45 end) as _45,(case when V._46='0%' then '' else V._46 end) as _46,(case when V._47='0%' then '' else V._47 end) as _47,(case when V._48='0%' then '' else V._48 end) as _48,(case when V._49='0%' then '' else V._49 end) as _49,(case when V._50='0%' then '' else V._50 end) as _50,(case when V._51='0%' then '' else V._51 end) as _51,(case when V._52='0%' then '' else V._52 end) as _52,(case when V._1='0%' then '' else V._1 end) as _1,(case when V._2='0%' then '' else V._2 end) as _2,(case when V._3='0%' then '' else V._3 end) as _3,(case when V._4='0%' then '' else V._4 end) as _4,(case when V._5='0%' then '' else V._5 end) as _5,(case when V._6='0%' then '' else V._6 end) as _6,(case when V._7='0%' then '' else V._7 end) as _7,(case when V._8='0%' then '' else V._8 end) as _8,(case when V._9='0%' then '' else V._9 end) as _9,(case when V._10='0%' then '' else V._10 end) as _10,(case when V._11='0%' then '' else V._11 end) as _11,(case when V._12='0%' then '' else V._12 end) as _12,(case when V._13='0%' then '' else V._13 end) as _13,(case when V._14='0%' then '' else V._14 end) as _14,(case when V._15='0%' then '' else V._15 end) as _15,(case when V._16='0%' then '' else V._16 end) as _16,(case when V._17='0%' then '' else V._17 end) as _17,(case when V._18='0%' then '' else V._18 end) as _18,(case when V._19='0%' then '' else V._19 end) as _19,(case when V._20='0%' then '' else V._20 end) as _20,(case when V._21='0%' then '' else V._21 end) as _21,(case when V._22='0%' then '' else V._22 end) as _22,(case when V._23='0%' then '' else V._23 end) as _23,(case when V._24='0%' then '' else V._24 end) as _24,(case when V._25='0%' then '' else V._25 end) as _25 from(select V.Inicio, V.Sem, cast(V._35 as varchar)+'%' as _35,cast(V._36 as varchar)+'%' as _36,cast(V._37 as varchar)+'%' as _37,cast(V._38 as varchar)+'%' as _38,cast(V._39 as varchar)+'%' as _39,cast(V._40 as varchar)+'%' as _40,cast(V._41 as varchar)+'%' as _41,cast(V._42 as varchar)+'%' as _42,cast(V._43 as varchar)+'%' as _43,cast(V._44 as varchar)+'%' as _44,cast(V._45 as varchar)+'%' as _45,cast(V._46 as varchar)+'%' as _46,cast(V._47 as varchar)+'%' as _47,cast(V._48 as varchar)+'%' as _48,cast(V._49 as varchar)+'%' as _49,cast(V._50 as varchar)+'%' as _50,cast(V._51 as varchar)+'%' as _51,cast(V._52 as varchar)+'%' as _52,cast(V._1 as varchar)+'%' as _1,cast(V._2 as varchar)+'%' _2,cast(V._3 as varchar)+'%' as _3,cast(V._4 as varchar)+'%' as _4,cast(V._5 as varchar)+'%' as _5,cast(V._6 as varchar)+'%' as _6,cast(V._7 as varchar)+'%' as _7,cast(V._8 as varchar)+'%' as _8,cast(V._9 as varchar)+'%' as _9,cast(V._10 as varchar)+'%' as _10,cast(V._11 as varchar)+'%' as _11,cast(V._12 as varchar)+'%' as _12,cast(V._13 as varchar)+'%' as _13,cast(V._14 as varchar)+'%' as _14,cast(V._15 as varchar)+'%' as _15,cast(V._16 as varchar)+'%' as _16,cast(V._17 as varchar)+'%' as _17,cast(V._18 as varchar)+'%' as _18,cast(V._19 as varchar)+'%' as _19,cast(V._20 as varchar)+'%'as _20,cast(V._21 as varchar)+'%' as _21,cast(V._22 as varchar)+'%' as _22,cast(V._23 as varchar)+'%' as _23,cast(V._24 as varchar)+'%' as _24,cast(V._25 as varchar)+'%' as _25 from (SELECT V.Inicio, V.Sem,(case when A._35=0 then '' else case when V._35=0 then '' else round((A._35/V._35)*100,0) end end) as _35,(case when A._36=0 then 0 else case when V._36=0 then 0 else round((A._36/V._36)*100,0) end end) as _36,(case when A._37=0 then 0 else case when V._37=0 then 0 else round((A._37/V._37)*100,0) end end) as _37,(case when A._38=0 then 0 else case when V._38=0 then 0 else round((A._38/V._38)*100,0) end end) as _38,(case when A._39=0 then 0 else case when V._39=0 then 0 else round((A._39/V._39)*100,0) end end) as _39,(case when A._40=0 then 0 else case when V._40=0 then 0 else round((A._40/V._40)*100,0) end end) as _40,(case when A._41=0 then 0 else case when V._41=0 then 0 else round((A._41/V._41)*100,0) end end) as _41,(case when A._42=0 then 0 else case when V._42=0 then 0 else round((A._42/V._42)*100,0) end end) as _42,(case when A._43=0 then 0 else case when V._43=0 then 0 else round((A._43/V._43)*100,0) end end) as _43,(case when A._44=0 then 0 else case when V._44=0 then 0 else round((A._44/V._44)*100,0) end end) as _44,(case when A._45=0 then 0 else case when V._45=0 then 0 else round((A._45/V._45)*100,0) end end) as _45,(case when A._46=0 then 0 else case when V._46=0 then 0 else round((A._46/V._46)*100,0) end end) as _46,(case when A._47=0 then 0 else case when V._47=0 then 0 else round((A._47/V._47)*100,0) end end) as _47,(case when A._48=0 then 0 else case when V._48=0 then 0 else round((A._48/V._48)*100,0) end end) as _48,(case when A._49=0 then 0 else case when V._49=0 then 0 else round((A._49/V._49)*100,0) end end) as _49,(case when A._50=0 then 0 else case when V._50=0 then 0 else round((A._50/V._50)*100,0) end end) as _50,(case when A._51=0 then 0 else case when V._51=0 then 0 else round((A._51/V._51)*100,0) end end) as _51,(case when A._52=0 then 0 else case when V._52=0 then 0 else round((A._52/V._52)*100,0) end end) as _52,(case when A._1=0 then 0 else case when V._1=0 then 0 else round((A._1/V._1)*100,0) end end) as _1,(case when A._2=0 then 0 else case when V._2=0 then 0 else round((A._2/V._2)*100,0) end end) as _2,(case when A._3=0 then 0 else case when V._3=0 then 0 else round((A._3/V._3)*100,0) end end) as _3,(case when A._4=0 then 0 else case when V._4=0 then 0 else round((A._4/V._4)*100,0) end end) as _4,(case when A._5=0 then 0 else case when V._5=0 then 0 else round((A._5/V._5)*100,0) end end) as _5,(case when A._6=0 then 0 else case when V._6=0 then 0 else round((A._6/V._6)*100,0) end end) as _6,(case when A._7=0 then 0 else case when V._7=0 then 0 else round((A._7/V._7)*100,0) end end) as _7,(case when A._8=0 then 0 else case when V._8=0 then 0 else round((A._8/V._8)*100,0) end end) as _8,(case when A._9=0 then 0 else case when V._9=0 then 0 else round((A._9/V._9)*100,0) end end) as _9,(case when A._10=0 then 0 else case when V._10=0 then 0 else round((A._10/V._10)*100,0) end end) as _10,(case when A._11=0 then 0 else case when V._11=0 then 0 else round((A._11/V._11)*100,0) end end) as _11,(case when A._12=0 then 0 else case when V._12=0 then 0 else round((A._12/V._12)*100,0) end end) as _12,(case when A._13=0 then 0 else case when V._13=0 then 0 else round((A._13/V._13)*100,0) end end) as _13,(case when A._14=0 then 0 else case when V._14=0 then 0 else round((A._14/V._14)*100,0) end end) as _14,(case when A._15=0 then 0 else case when V._15=0 then 0 else round((A._15/V._15)*100,0) end end) as _15,(case when A._16=0 then 0 else case when V._16=0 then 0 else round((A._16/V._16)*100,0) end end) as _16,(case when A._17=0 then 0 else case when V._17=0 then 0 else round((A._17/V._17)*100,0) end end) as _17,(case when A._18=0 then 0 else case when V._18=0 then 0 else round((A._18/V._18)*100,0) end end) as _18,(case when A._19=0 then 0 else case when V._19=0 then 0 else round((A._19/V._19)*100,0) end end) as _19,(case when A._20=0 then 0 else case when V._20=0 then 0 else round((A._20/V._20)*100,0) end end) as _20,(case when A._21=0 then 0 else case when V._21=0 then 0 else round((A._21/V._21)*100,0) end end) as _21,(case when A._22=0 then 0 else case when V._22=0 then 0 else round((A._22/V._22)*100,0) end end) as _22,(case when A._23=0 then 0 else case when V._23=0 then 0 else round((A._23/V._23)*100,0) end end) as _23,(case when A._24=0 then 0 else case when V._24=0 then 0 else round((A._24/V._24)*100,0) end end) as _24,(case when A._25=0 then 0 else case when V._25=0 then 0 else round((A._25/V._25)*100,0) end end) as _25 FROM(SELECT V.Temporada, isnull(V.[35],0) as _35,isnull(V.[36],0) as _36,isnull(V.[37],0) as _37,isnull(V.[38],0) as _38,isnull(V.[39],0) as _39,isnull(V.[40],0) as _40,isnull(V.[41],0) as _41,isnull(V.[42],0) as _42,isnull(V.[43],0) as _43,isnull(V.[44],0) as _44,isnull(V.[45],0) as _45,isnull(V.[46],0) as _46,isnull(V.[47],0) as _47,isnull(V.[48],0) as _48,isnull(V.[49],0) as _49,isnull(V.[50],0) as _50,isnull(V.[51],0) as _51,isnull(V.[52],0) as _52,isnull(V.[1],0) as _1,isnull(V.[2],0) as _2,isnull(V.[3],0) as _3,isnull(V.[4],0) as _4,isnull(V.[5],0) as _5,isnull(V.[6],0) as _6,isnull(V.[7],0) as _7,isnull(V.[8],0) as _8,isnull(V.[9],0) as _9,isnull(V.[10],0) as _10,isnull(V.[11],0) as _11,isnull(V.[12],0) as _12,isnull(V.[13],0) as _13,isnull(V.[14],0) as _14,isnull(V.[15],0) as _15,isnull(V.[16],0) as _16,isnull(V.[17],0) as _17,isnull(V.[18],0) as _18,isnull(V.[19],0) as _19,isnull(V.[20],0) as _20,isnull(V.[21],0) as _21,isnull(V.[22],0) as _22,isnull(V.[23],0) as _23,isnull(V.[24],0) as _24,isnull(V.[25],0) as _25 FROM(Select * from(SELECT Temporada, Semana, SUM(Convertidas) AS Convertidas FROM SEasonsun1213..UV_ProdRecepcion where CodEstatus<> 'C' and temporada = (select temporada from CatSemanas where GETDATE() between Inicio and Fin) and " +
               "DescProducto = 'ARANDANO' GROUP BY Temporada, Semana Union All SELECT Temporada, Semana, SUM(Convertidas) AS Convertidas FROM SeasonPlan..UV_ProdRecepcion where CodEstatus <> 'C' and temporada = (select temporada from CatSemanas where GETDATE() between Inicio and Fin) and " +
               "DescProducto = 'ARANDANO' GROUP BY Temporada, Semana)V PIVOT(SUM(Convertidas) FOR Semana in ([35],[36],[37],[38],[39],[40],[41],[42],[43],[44],[45],[46],[47],[48],[49],[50],[51],[52],[1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],[17],[18],[19],[20],[21],[22],[23],[24],[25]))X)V)A left join(Select V.Inicio, V.Temporada, V.Sem, isnull(V.[35], 0) as _35, isnull(V.[36], 0) as _36, isnull(V.[37], 0) as _37, isnull(V.[38], 0) as _38, isnull(V.[39], 0) as _39, isnull(V.[40], 0) as _40, isnull(V.[41], 0) as _41, isnull(V.[42], 0) as _42, isnull(V.[43], 0) as _43, isnull(V.[44], 0) as _44, isnull(V.[45], 0) as _45, isnull(V.[46], 0) as _46, isnull(V.[47], 0) as _47, isnull(V.[48], 0) as _48, isnull(V.[49], 0) as _49, isnull(V.[50], 0) as _50, isnull(V.[51], 0) as _51, isnull(V.[52], 0) as _52, isnull(V.[1], 0) as _1, isnull(V.[2], 0) as _2, isnull(V.[3], 0) as _3, isnull(V.[4], 0) as _4, isnull(V.[5], 0) as _5, isnull(V.[6], 0) as _6, isnull(V.[7], 0) as _7, isnull(V.[8], 0) as _8, isnull(V.[9], 0) as _9, isnull(V.[10], 0) as _10, isnull(V.[11], 0) as _11, isnull(V.[12], 0) as _12, isnull(V.[13], 0) as _13, isnull(V.[14], 0) as _14, isnull(V.[15], 0) as _15, isnull(V.[16], 0) as _16, isnull(V.[17], 0) as _17, isnull(V.[18], 0) as _18, isnull(V.[19], 0) as _19, isnull(V.[20], 0) as _20, isnull(V.[21], 0) as _21, isnull(V.[22], 0) as _22, isnull(V.[23], 0) as _23, isnull(V.[24], 0) as _24, isnull(V.[25], 0) as _25 from(select * from(SELECT S.Semana AS Sem, S.Inicio, E.Semanas, E.Cantidad, E.Temporada FROM Estimacion_Berries E left join CatSemanas S on E.Temporada = S.Temporada AND E.Semana = S.Semana where E.Temporada = (select temporada from CatSemanas where GETDATE() between Inicio and Fin) and " +
               "E.Concepto = 'PROYECCION SEMANAL' and E.Cultivo = 3)V PIVOT(SUM(Cantidad) For Semanas in ([35],[36],[37],[38],[39],[40],[41],[42],[43],[44],[45],[46],[47],[48],[49],[50],[51],[52],[1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],[17],[18],[19],[20],[21],[22],[23],[24],[25]))X)V)V on A.Temporada = V.Temporada)V)V ORDER BY V.Inicio").ToList();
            return Json(estimacion, JsonRequestBehavior.AllowGet);
        }

        [HttpPost]
        public ActionResult Arandano(int[] array, string concepto)
        {
            JsonResult dtaEjecucionTarea = default(JsonResult);
            try
            {
                if (concepto == "PROYECCION SEMANAL")
                {
                    if (Update_Arandano(array, concepto))
                    {
                        dtaEjecucionTarea = Json(new
                        {
                            rstProceso = "true",
                            MessageGestion = "Cambios guardados con éxito"
                        });
                    }
                    else
                    {
                        dtaEjecucionTarea = Json(new
                        {
                            rstProceso = "false",
                            MessageGestion = "Error, algo salió mal, intente de nuevo"
                        });
                    }
                }
                else
                {
                    if (Update_ArandanoE(array, concepto))
                    {
                        dtaEjecucionTarea = Json(new
                        {
                            rstProceso = "true",
                            MessageGestion = "Cambios guardados con éxito"
                        });
                    }
                    else
                    {
                        dtaEjecucionTarea = Json(new
                        {
                            rstProceso = "false",
                            MessageGestion = "Error, algo salió mal, intente de nuevo"
                        });
                    }
                }
                return dtaEjecucionTarea;
            }
            catch (Exception e)
            {
                e.ToString();
            }
            return dtaEjecucionTarea;
        }

        public bool Update_Arandano(int[] array, string concepto)
        {
            if (array != null)
            {
                var sem = array[0];
                var _35 = array[1];
                var _36 = array[2];
                var _37 = array[3];
                var _38 = array[4];
                var _39 = array[5];
                var _40 = array[6];
                var _41 = array[7];
                var _42 = array[8];
                var _43 = array[9];
                var _44 = array[10];
                var _45 = array[11];
                var _46 = array[12];
                var _47 = array[13];
                var _48 = array[14];
                var _49 = array[15];
                var _50 = array[16];
                var _51 = array[17];
                var _52 = array[18];
                var _1 = array[19];
                var _2 = array[20];
                var _3 = array[21];
                var _4 = array[22];
                var _5 = array[23];
                var _6 = array[24];
                var _7 = array[25];
                var _8 = array[26];
                var _9 = array[27];
                var _10 = array[28];
                var _11 = array[29];
                var _12 = array[30];
                var _13 = array[31];
                var _14 = array[32];
                var _15 = array[33];
                var _16 = array[34];
                var _17 = array[35];
                var _18 = array[36];
                var _19 = array[37];
                var _20 = array[38];
                var _21 = array[39];
                var _22 = array[40];
                var _23 = array[41];
                var _24 = array[42];
                var _25 = array[43];

                var item_35 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 35 && x.Cultivo == 3 && x.Concepto == concepto).FirstOrDefault();
                item_35.Cantidad = _35;
                item_35.Fecha = DateTime.Now;

                var item_36 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 36 && x.Cultivo == 3 && x.Concepto == concepto).FirstOrDefault();
                item_36.Cantidad = _36;
                item_36.Fecha = DateTime.Now;

                var item_37 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 37 && x.Cultivo == 3 && x.Concepto == concepto).FirstOrDefault();
                item_37.Cantidad = _37;
                item_37.Fecha = DateTime.Now;

                var item_38 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 38 && x.Cultivo == 3 && x.Concepto == concepto).FirstOrDefault();
                item_38.Cantidad = _38;
                item_38.Fecha = DateTime.Now;

                var item_39 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 39 && x.Cultivo == 3 && x.Concepto == concepto).FirstOrDefault();
                item_39.Cantidad = _39;
                item_39.Fecha = DateTime.Now;

                var item_40 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 40 && x.Cultivo == 3 && x.Concepto == concepto).FirstOrDefault();
                item_40.Cantidad = _40;
                item_40.Fecha = DateTime.Now;

                var item_41 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 41 && x.Cultivo == 3 && x.Concepto == concepto).FirstOrDefault();
                item_41.Cantidad = _41;
                item_41.Fecha = DateTime.Now;

                var item_42 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 42 && x.Cultivo == 3 && x.Concepto == concepto).FirstOrDefault();
                item_42.Cantidad = _42;
                item_42.Fecha = DateTime.Now;

                var item_43 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 43 && x.Cultivo == 3 && x.Concepto == concepto).FirstOrDefault();
                item_43.Cantidad = _43;
                item_43.Fecha = DateTime.Now;

                var item_44 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 44 && x.Cultivo == 3 && x.Concepto == concepto).FirstOrDefault();
                item_44.Cantidad = _44;
                item_44.Fecha = DateTime.Now;

                var item_45 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 45 && x.Cultivo == 3 && x.Concepto == concepto).FirstOrDefault();
                item_45.Cantidad = _45;
                item_45.Fecha = DateTime.Now;

                var item_46 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 46 && x.Cultivo == 3 && x.Concepto == concepto).FirstOrDefault();
                item_46.Cantidad = _46;
                item_46.Fecha = DateTime.Now;

                var item_47 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 47 && x.Cultivo == 3 && x.Concepto == concepto).FirstOrDefault();
                item_47.Cantidad = _47;
                item_47.Fecha = DateTime.Now;

                var item_48 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 48 && x.Cultivo == 3 && x.Concepto == concepto).FirstOrDefault();
                item_48.Cantidad = _48;
                item_48.Fecha = DateTime.Now;

                var item_49 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 49 && x.Cultivo == 3 && x.Concepto == concepto).FirstOrDefault();
                item_49.Cantidad = _49;
                item_49.Fecha = DateTime.Now;

                var item_50 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 50 && x.Cultivo == 3 && x.Concepto == concepto).FirstOrDefault();
                item_50.Cantidad = _50;
                item_50.Fecha = DateTime.Now;

                var item_51 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 51 && x.Cultivo == 3 && x.Concepto == concepto).FirstOrDefault();
                item_51.Cantidad = _51;
                item_51.Fecha = DateTime.Now;

                var item_52 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 52 && x.Cultivo == 3 && x.Concepto == concepto).FirstOrDefault();
                item_52.Cantidad = _52;
                item_52.Fecha = DateTime.Now;

                var item_1 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 1 && x.Cultivo == 3 && x.Concepto == concepto).FirstOrDefault();
                item_1.Cantidad = _1;
                item_1.Fecha = DateTime.Now;

                var item_2 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 2 && x.Cultivo == 3 && x.Concepto == concepto).FirstOrDefault();
                item_2.Cantidad = _2;
                item_2.Fecha = DateTime.Now;

                var item_3 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 3 && x.Cultivo == 3 && x.Concepto == concepto).FirstOrDefault();
                item_3.Cantidad = _3;
                item_3.Fecha = DateTime.Now;

                var item_4 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 4 && x.Cultivo == 3 && x.Concepto == concepto).FirstOrDefault();
                item_4.Cantidad = _4;
                item_4.Fecha = DateTime.Now;

                var item_5 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 5 && x.Cultivo == 3 && x.Concepto == concepto).FirstOrDefault();
                item_5.Cantidad = _5;
                item_5.Fecha = DateTime.Now;

                var item_6 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 6 && x.Cultivo == 3 && x.Concepto == concepto).FirstOrDefault();
                item_6.Cantidad = _6;
                item_6.Fecha = DateTime.Now;

                var item_7 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 7 && x.Cultivo == 3 && x.Concepto == concepto).FirstOrDefault();
                item_7.Cantidad = _7;
                item_7.Fecha = DateTime.Now;

                var item_8 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 8 && x.Cultivo == 3 && x.Concepto == concepto).FirstOrDefault();
                item_8.Cantidad = _8;
                item_8.Fecha = DateTime.Now;

                var item_9 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 9 && x.Cultivo == 3 && x.Concepto == concepto).FirstOrDefault();
                item_9.Cantidad = _9;
                item_9.Fecha = DateTime.Now;

                var item_10 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 10 && x.Cultivo == 3 && x.Concepto == concepto).FirstOrDefault();
                item_10.Cantidad = _10;
                item_10.Fecha = DateTime.Now;

                var item_11 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 11 && x.Cultivo == 3 && x.Concepto == concepto).FirstOrDefault();
                item_11.Cantidad = _11;
                item_11.Fecha = DateTime.Now;

                var item_12 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 12 && x.Cultivo == 3 && x.Concepto == concepto).FirstOrDefault();
                item_12.Cantidad = _12;
                item_12.Fecha = DateTime.Now;

                var item_13 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 13 && x.Cultivo == 3 && x.Concepto == concepto).FirstOrDefault();
                item_13.Cantidad = _13;
                item_13.Fecha = DateTime.Now;

                var item_14 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 14 && x.Cultivo == 3 && x.Concepto == concepto).FirstOrDefault();
                item_14.Cantidad = _14;
                item_14.Fecha = DateTime.Now;

                var item_15 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 15 && x.Cultivo == 3 && x.Concepto == concepto).FirstOrDefault();
                item_15.Cantidad = _15;
                item_15.Fecha = DateTime.Now;

                var item_16 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 16 && x.Cultivo == 3 && x.Concepto == concepto).FirstOrDefault();
                item_16.Cantidad = _16;
                item_16.Fecha = DateTime.Now;

                var item_17 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 17 && x.Cultivo == 3 && x.Concepto == concepto).FirstOrDefault();
                item_17.Cantidad = _17;
                item_17.Fecha = DateTime.Now;

                var item_18 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 18 && x.Cultivo == 3 && x.Concepto == concepto).FirstOrDefault();
                item_18.Cantidad = _18;
                item_18.Fecha = DateTime.Now;

                var item_19 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 19 && x.Cultivo == 3 && x.Concepto == concepto).FirstOrDefault();
                item_19.Cantidad = _19;
                item_19.Fecha = DateTime.Now;

                var item_20 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 20 && x.Cultivo == 3 && x.Concepto == concepto).FirstOrDefault();
                item_20.Cantidad = _20;
                item_20.Fecha = DateTime.Now;

                var item_21 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 21 && x.Cultivo == 3 && x.Concepto == concepto).FirstOrDefault();
                item_21.Cantidad = _21;
                item_21.Fecha = DateTime.Now;

                var item_22 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 22 && x.Cultivo == 3 && x.Concepto == concepto).FirstOrDefault();
                item_22.Cantidad = _22;
                item_22.Fecha = DateTime.Now;

                var item_23 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 23 && x.Cultivo == 3 && x.Concepto == concepto).FirstOrDefault();
                item_23.Cantidad = _23;
                item_23.Fecha = DateTime.Now;

                var item_24 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 24 && x.Cultivo == 3 && x.Concepto == concepto).FirstOrDefault();
                item_24.Cantidad = _24;
                item_24.Fecha = DateTime.Now;

                var item_25 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 25 && x.Cultivo == 3 && x.Concepto == concepto).FirstOrDefault();
                item_25.Cantidad = _25;
                item_25.Fecha = DateTime.Now;

                bd.SaveChanges();
                return true;
            }
            else
            {
                return false;
            }
        }

        public bool Update_ArandanoE(int[] array, string concepto)
        {
            if (array != null)
            {
                var _35 = array[0];
                var _36 = array[1];
                var _37 = array[2];
                var _38 = array[3];
                var _39 = array[4];
                var _40 = array[5];
                var _41 = array[6];
                var _42 = array[7];
                var _43 = array[8];
                var _44 = array[9];
                var _45 = array[10];
                var _46 = array[11];
                var _47 = array[12];
                var _48 = array[13];
                var _49 = array[14];
                var _50 = array[15];
                var _51 = array[16];
                var _52 = array[17];
                var _1 = array[18];
                var _2 = array[19];
                var _3 = array[20];
                var _4 = array[21];
                var _5 = array[22];
                var _6 = array[23];
                var _7 = array[24];
                var _8 = array[25];
                var _9 = array[26];
                var _10 = array[27];
                var _11 = array[28];
                var _12 = array[29];
                var _13 = array[30];
                var _14 = array[31];
                var _15 = array[32];
                var _16 = array[33];
                var _17 = array[34];
                var _18 = array[35];
                var _19 = array[36];
                var _20 = array[37];
                var _21 = array[38];
                var _22 = array[39];
                var _23 = array[40];
                var _24 = array[41];
                var _25 = array[42];

                var item_35 = bd.Estimacion_Berries.Where(x => x.Semanas == 35 && x.Cultivo == 3 && x.Concepto == concepto).FirstOrDefault();
                item_35.Cantidad = _35;
                item_35.Fecha = DateTime.Now;

                var item_36 = bd.Estimacion_Berries.Where(x => x.Semanas == 36 && x.Cultivo == 3 && x.Concepto == concepto).FirstOrDefault();
                item_36.Cantidad = _36;
                item_36.Fecha = DateTime.Now;

                var item_37 = bd.Estimacion_Berries.Where(x => x.Semanas == 37 && x.Cultivo == 3 && x.Concepto == concepto).FirstOrDefault();
                item_37.Cantidad = _37;
                item_37.Fecha = DateTime.Now;

                var item_38 = bd.Estimacion_Berries.Where(x => x.Semanas == 38 && x.Cultivo == 3 && x.Concepto == concepto).FirstOrDefault();
                item_38.Cantidad = _38;
                item_38.Fecha = DateTime.Now;

                var item_39 = bd.Estimacion_Berries.Where(x => x.Semanas == 39 && x.Cultivo == 3 && x.Concepto == concepto).FirstOrDefault();
                item_39.Cantidad = _39;
                item_39.Fecha = DateTime.Now;

                var item_40 = bd.Estimacion_Berries.Where(x => x.Semanas == 40 && x.Cultivo == 3 && x.Concepto == concepto).FirstOrDefault();
                item_40.Cantidad = _40;
                item_40.Fecha = DateTime.Now;

                var item_41 = bd.Estimacion_Berries.Where(x => x.Semanas == 41 && x.Cultivo == 3 && x.Concepto == concepto).FirstOrDefault();
                item_41.Cantidad = _41;
                item_41.Fecha = DateTime.Now;

                var item_42 = bd.Estimacion_Berries.Where(x => x.Semanas == 42 && x.Cultivo == 3 && x.Concepto == concepto).FirstOrDefault();
                item_42.Cantidad = _42;
                item_42.Fecha = DateTime.Now;

                var item_43 = bd.Estimacion_Berries.Where(x => x.Semanas == 43 && x.Cultivo == 3 && x.Concepto == concepto).FirstOrDefault();
                item_43.Cantidad = _43;
                item_43.Fecha = DateTime.Now;

                var item_44 = bd.Estimacion_Berries.Where(x => x.Semanas == 44 && x.Cultivo == 3 && x.Concepto == concepto).FirstOrDefault();
                item_44.Cantidad = _44;
                item_44.Fecha = DateTime.Now;

                var item_45 = bd.Estimacion_Berries.Where(x => x.Semanas == 45 && x.Cultivo == 3 && x.Concepto == concepto).FirstOrDefault();
                item_45.Cantidad = _45;
                item_45.Fecha = DateTime.Now;

                var item_46 = bd.Estimacion_Berries.Where(x => x.Semanas == 46 && x.Cultivo == 3 && x.Concepto == concepto).FirstOrDefault();
                item_46.Cantidad = _46;
                item_46.Fecha = DateTime.Now;

                var item_47 = bd.Estimacion_Berries.Where(x => x.Semanas == 47 && x.Cultivo == 3 && x.Concepto == concepto).FirstOrDefault();
                item_47.Cantidad = _47;
                item_47.Fecha = DateTime.Now;

                var item_48 = bd.Estimacion_Berries.Where(x => x.Semanas == 48 && x.Cultivo == 3 && x.Concepto == concepto).FirstOrDefault();
                item_48.Cantidad = _48;
                item_48.Fecha = DateTime.Now;

                var item_49 = bd.Estimacion_Berries.Where(x => x.Semanas == 49 && x.Cultivo == 3 && x.Concepto == concepto).FirstOrDefault();
                item_49.Cantidad = _49;
                item_49.Fecha = DateTime.Now;

                var item_50 = bd.Estimacion_Berries.Where(x => x.Semanas == 50 && x.Cultivo == 3 && x.Concepto == concepto).FirstOrDefault();
                item_50.Cantidad = _50;
                item_50.Fecha = DateTime.Now;

                var item_51 = bd.Estimacion_Berries.Where(x => x.Semanas == 51 && x.Cultivo == 3 && x.Concepto == concepto).FirstOrDefault();
                item_51.Cantidad = _51;
                item_51.Fecha = DateTime.Now;

                var item_52 = bd.Estimacion_Berries.Where(x => x.Semanas == 52 && x.Cultivo == 3 && x.Concepto == concepto).FirstOrDefault();
                item_52.Cantidad = _52;
                item_52.Fecha = DateTime.Now;

                var item_1 = bd.Estimacion_Berries.Where(x => x.Semanas == 1 && x.Cultivo == 3 && x.Concepto == concepto).FirstOrDefault();
                item_1.Cantidad = _1;
                item_1.Fecha = DateTime.Now;

                var item_2 = bd.Estimacion_Berries.Where(x => x.Semanas == 2 && x.Cultivo == 3 && x.Concepto == concepto).FirstOrDefault();
                item_2.Cantidad = _2;
                item_2.Fecha = DateTime.Now;

                var item_3 = bd.Estimacion_Berries.Where(x => x.Semanas == 3 && x.Cultivo == 3 && x.Concepto == concepto).FirstOrDefault();
                item_3.Cantidad = _3;
                item_3.Fecha = DateTime.Now;

                var item_4 = bd.Estimacion_Berries.Where(x => x.Semanas == 4 && x.Cultivo == 3 && x.Concepto == concepto).FirstOrDefault();
                item_4.Cantidad = _4;
                item_4.Fecha = DateTime.Now;

                var item_5 = bd.Estimacion_Berries.Where(x => x.Semanas == 5 && x.Cultivo == 3 && x.Concepto == concepto).FirstOrDefault();
                item_5.Cantidad = _5;
                item_5.Fecha = DateTime.Now;

                var item_6 = bd.Estimacion_Berries.Where(x => x.Semanas == 6 && x.Cultivo == 3 && x.Concepto == concepto).FirstOrDefault();
                item_6.Cantidad = _6;
                item_6.Fecha = DateTime.Now;

                var item_7 = bd.Estimacion_Berries.Where(x => x.Semanas == 7 && x.Cultivo == 3 && x.Concepto == concepto).FirstOrDefault();
                item_7.Cantidad = _7;
                item_7.Fecha = DateTime.Now;

                var item_8 = bd.Estimacion_Berries.Where(x => x.Semanas == 8 && x.Cultivo == 3 && x.Concepto == concepto).FirstOrDefault();
                item_8.Cantidad = _8;
                item_8.Fecha = DateTime.Now;

                var item_9 = bd.Estimacion_Berries.Where(x => x.Semanas == 9 && x.Cultivo == 3 && x.Concepto == concepto).FirstOrDefault();
                item_9.Cantidad = _9;
                item_9.Fecha = DateTime.Now;

                var item_10 = bd.Estimacion_Berries.Where(x => x.Semanas == 10 && x.Cultivo == 3 && x.Concepto == concepto).FirstOrDefault();
                item_10.Cantidad = _10;
                item_10.Fecha = DateTime.Now;

                var item_11 = bd.Estimacion_Berries.Where(x => x.Semanas == 11 && x.Cultivo == 3 && x.Concepto == concepto).FirstOrDefault();
                item_11.Cantidad = _11;
                item_11.Fecha = DateTime.Now;

                var item_12 = bd.Estimacion_Berries.Where(x => x.Semanas == 12 && x.Cultivo == 3 && x.Concepto == concepto).FirstOrDefault();
                item_12.Cantidad = _12;
                item_12.Fecha = DateTime.Now;

                var item_13 = bd.Estimacion_Berries.Where(x => x.Semanas == 13 && x.Cultivo == 3 && x.Concepto == concepto).FirstOrDefault();
                item_13.Cantidad = _13;
                item_13.Fecha = DateTime.Now;

                var item_14 = bd.Estimacion_Berries.Where(x => x.Semanas == 14 && x.Cultivo == 3 && x.Concepto == concepto).FirstOrDefault();
                item_14.Cantidad = _14;
                item_14.Fecha = DateTime.Now;

                var item_15 = bd.Estimacion_Berries.Where(x => x.Semanas == 15 && x.Cultivo == 3 && x.Concepto == concepto).FirstOrDefault();
                item_15.Cantidad = _15;
                item_15.Fecha = DateTime.Now;

                var item_16 = bd.Estimacion_Berries.Where(x => x.Semanas == 16 && x.Cultivo == 3 && x.Concepto == concepto).FirstOrDefault();
                item_16.Cantidad = _16;
                item_16.Fecha = DateTime.Now;

                var item_17 = bd.Estimacion_Berries.Where(x => x.Semanas == 17 && x.Cultivo == 3 && x.Concepto == concepto).FirstOrDefault();
                item_17.Cantidad = _17;
                item_17.Fecha = DateTime.Now;

                var item_18 = bd.Estimacion_Berries.Where(x => x.Semanas == 18 && x.Cultivo == 3 && x.Concepto == concepto).FirstOrDefault();
                item_18.Cantidad = _18;
                item_18.Fecha = DateTime.Now;

                var item_19 = bd.Estimacion_Berries.Where(x => x.Semanas == 19 && x.Cultivo == 3 && x.Concepto == concepto).FirstOrDefault();
                item_19.Cantidad = _19;
                item_19.Fecha = DateTime.Now;

                var item_20 = bd.Estimacion_Berries.Where(x => x.Semanas == 20 && x.Cultivo == 3 && x.Concepto == concepto).FirstOrDefault();
                item_20.Cantidad = _20;
                item_20.Fecha = DateTime.Now;

                var item_21 = bd.Estimacion_Berries.Where(x => x.Semanas == 21 && x.Cultivo == 3 && x.Concepto == concepto).FirstOrDefault();
                item_21.Cantidad = _21;
                item_21.Fecha = DateTime.Now;

                var item_22 = bd.Estimacion_Berries.Where(x => x.Semanas == 22 && x.Cultivo == 3 && x.Concepto == concepto).FirstOrDefault();
                item_22.Cantidad = _22;
                item_22.Fecha = DateTime.Now;

                var item_23 = bd.Estimacion_Berries.Where(x => x.Semanas == 23 && x.Cultivo == 3 && x.Concepto == concepto).FirstOrDefault();
                item_23.Cantidad = _23;
                item_23.Fecha = DateTime.Now;

                var item_24 = bd.Estimacion_Berries.Where(x => x.Semanas == 24 && x.Cultivo == 3 && x.Concepto == concepto).FirstOrDefault();
                item_24.Cantidad = _24;
                item_24.Fecha = DateTime.Now;

                var item_25 = bd.Estimacion_Berries.Where(x => x.Semanas == 25 && x.Cultivo == 3 && x.Concepto == concepto).FirstOrDefault();
                item_25.Cantidad = _25;
                item_25.Fecha = DateTime.Now;

                bd.SaveChanges();
                return true;
            }
            else
            {
                return false;
            }
        }

        //FRESA
        public ActionResult Fresa()
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

        public JsonResult RecepcionFresa()
        {
            List<R_Berries> recepcion_real = bd.Database.SqlQuery<R_Berries>("SELECT 'TOTAL' as 'SEMANA',  round(isnull(V.[37],0),0) as _37,round(isnull(V.[38],0),0) as _38,round(isnull(V.[39],0),0) as _39,round(isnull(V.[40],0),0) as _40,round(isnull(V.[41],0),0) as _41,round(isnull(V.[42],0),0) as _42,round(isnull(V.[43],0),0) as _43,round(isnull(V.[44],0),0) as _44,round(isnull(V.[45],0),0) as _45,round(isnull(V.[46],0),0) as _46,round(isnull(V.[47],0),0) as _47,round(isnull(V.[48],0),0) as _48,round(isnull(V.[49],0),0) as _49,round(isnull(V.[50],0),0) as _50,round(isnull(V.[51],0),0) as _51,round(isnull(V.[52],0),0) as _52,round(isnull(V.[1],0),0) as _1,round(isnull(V.[2],0),0) as _2,round(isnull(V.[3],0),0) as _3,round(isnull(V.[4],0),0) as _4,round(isnull(V.[5],0),0) as _5,round(isnull(V.[6],0),0) as _6,round(isnull(V.[7],0),0) as _7,round(isnull(V.[8],0),0) as _8,round(isnull(V.[9],0),0) as _9,round(isnull(V.[10],0),0) as _10,round(isnull(V.[11],0),0) as _11,round(isnull(V.[12],0),0) as _12,round(isnull(V.[13],0),0) as _13,round(isnull(V.[14],0),0) as _14,round(isnull(V.[15],0),0) as _15,round(isnull(V.[16],0),0) as _16,round(isnull(V.[17],0),0) as _17,round(isnull(V.[18],0),0) as _18,round(isnull(V.[19],0),0) as _19,round(isnull(V.[20],0),0) as _20,round(isnull(V.[21],0),0) as _21,round(isnull(V.[22],0),0) as _22,round(isnull(V.[23],0),0) as _23,round(isnull(V.[24],0),0) as _24,round(isnull(V.[25],0),0) as _25 " +
                "FROM(Select * from(SELECT Semana, SUM(Convertidas) AS Convertidas FROM SEasonsun1213..UV_ProdRecepcion where CodEstatus <> 'C' and temporada = (select temporada from CatSemanas where GETDATE() between Inicio and Fin) and DescProducto = 'FRESA' GROUP BY Semana Union All " +
                "SELECT Semana, SUM(Convertidas) AS Convertidas FROM SeasonPlan..UV_ProdRecepcion where CodEstatus <> 'C' and temporada = (select temporada from CatSemanas where GETDATE() between Inicio and Fin) and DescProducto = 'FRESA' GROUP BY Semana" +
                ")V PIVOT(SUM(Convertidas) FOR Semana in ([37],[38],[39],[40],[41],[42],[43],[44],[45],[46],[47],[48],[49],[50],[51],[52],[1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],[17],[18],[19],[20],[21],[22],[23],[24],[25]))X)V ").ToList();
            return Json(recepcion_real, JsonRequestBehavior.AllowGet);
        }

        public JsonResult EstimacionSFresa()
        {
            List<E_Berries> estimacion = bd.Database.SqlQuery<E_Berries>("Select V.Sem AS 'SEMANA', round(isnull(V.[37],0),0) as _37,round(isnull(V.[38],0),0) as _38,round(isnull(V.[39],0),0) as _39,round(isnull(V.[40],0),0) as _40,round(isnull(V.[41],0),0) as _41,round(isnull(V.[42],0),0) as _42,round(isnull(V.[43],0),0) as _43,round(isnull(V.[44],0),0) as _44,round(isnull(V.[45],0),0) as _45,round(isnull(V.[46],0),0) as _46,round(isnull(V.[47],0),0) as _47,round(isnull(V.[48],0),0) as _48,round(isnull(V.[49],0),0) as _49,round(isnull(V.[50],0),0) as _50,round(isnull(V.[51],0),0) as _51,round(isnull(V.[52],0),0) as _52,round(isnull(V.[1],0),0) as _1,round(isnull(V.[2],0),0) as _2,round(isnull(V.[3],0),0) as _3,round(isnull(V.[4],0),0) as _4,round(isnull(V.[5],0),0) as _5,round(isnull(V.[6],0),0) as _6,round(isnull(V.[7],0),0) as _7,round(isnull(V.[8],0),0) as _8,round(isnull(V.[9],0),0) as _9,round(isnull(V.[10],0),0) as _10,round(isnull(V.[11],0),0) as _11,round(isnull(V.[12],0),0) as _12,round(isnull(V.[13],0),0) as _13,round(isnull(V.[14],0),0) as _14,round(isnull(V.[15],0),0) as _15,round(isnull(V.[16],0),0) as _16,round(isnull(V.[17],0),0) as _17,round(isnull(V.[18],0),0) as _18,round(isnull(V.[19],0),0) as _19,round(isnull(V.[20],0),0) as _20,round(isnull(V.[21],0),0) as _21,round(isnull(V.[22],0),0) as _22,round(isnull(V.[23],0),0) as _23,round(isnull(V.[24],0),0) as _24,round(isnull(V.[25],0),0) as _25 " +
                "from(Select * from(SELECT S.Semana AS Sem, S.Inicio, E.Semanas, E.Cantidad FROM Estimacion_Berries E left join CatSemanas S on E.Temporada = S.Temporada AND E.Semana = S.Semana " +
                "where E.Temporada = (select temporada from CatSemanas where GETDATE() between Inicio and Fin) and E.Concepto = 'PROYECCION SEMANAL' and E.Cultivo = 4)V " +
                "PIVOT(SUM(Cantidad) For Semanas in ([37],[38],[39],[40],[41],[42],[43],[44],[45],[46],[47],[48],[49],[50],[51],[52],[1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],[17],[18],[19],[20],[21],[22],[23],[24],[25]))X)V ORDER BY V.Inicio").ToList();
            return Json(estimacion, JsonRequestBehavior.AllowGet);
        }

        public JsonResult EstimacionSFresaP()
        {
            List<E_BerriesP> estimacion = bd.Database.SqlQuery<E_BerriesP>("select V.Sem AS 'SEMANA', (case when V._37='0%' then '' else V._37 end) as _37,(case when V._38='0%' then '' else V._38 end) as _38,(case when V._39='0%' then '' else V._39 end) as _39,(case when V._40='0%' then '' else V._40 end) as _40,(case when V._41='0%' then '' else V._41 end) as _41,(case when V._42='0%' then '' else V._42 end) as _42,(case when V._43='0%' then '' else V._43 end) as _43,(case when V._44='0%' then '' else V._44 end) as _44,(case when V._45='0%' then '' else V._45 end) as _45,(case when V._46='0%' then '' else V._46 end) as _46,(case when V._47='0%' then '' else V._47 end) as _47,(case when V._48='0%' then '' else V._48 end) as _48,(case when V._49='0%' then '' else V._49 end) as _49,(case when V._50='0%' then '' else V._50 end) as _50,(case when V._51='0%' then '' else V._51 end) as _51,(case when V._52='0%' then '' else V._52 end) as _52,(case when V._1='0%' then '' else V._1 end) as _1,(case when V._2='0%' then '' else V._2 end) as _2,(case when V._3='0%' then '' else V._3 end) as _3,(case when V._4='0%' then '' else V._4 end) as _4,(case when V._5='0%' then '' else V._5 end) as _5,(case when V._6='0%' then '' else V._6 end) as _6,(case when V._7='0%' then '' else V._7 end) as _7,(case when V._8='0%' then '' else V._8 end) as _8,(case when V._9='0%' then '' else V._9 end) as _9,(case when V._10='0%' then '' else V._10 end) as _10,(case when V._11='0%' then '' else V._11 end) as _11,(case when V._12='0%' then '' else V._12 end) as _12,(case when V._13='0%' then '' else V._13 end) as _13,(case when V._14='0%' then '' else V._14 end) as _14,(case when V._15='0%' then '' else V._15 end) as _15,(case when V._16='0%' then '' else V._16 end) as _16,(case when V._17='0%' then '' else V._17 end) as _17,(case when V._18='0%' then '' else V._18 end) as _18,(case when V._19='0%' then '' else V._19 end) as _19,(case when V._20='0%' then '' else V._20 end) as _20,(case when V._21='0%' then '' else V._21 end) as _21,(case when V._22='0%' then '' else V._22 end) as _22,(case when V._23='0%' then '' else V._23 end) as _23,(case when V._24='0%' then '' else V._24 end) as _24,(case when V._25='0%' then '' else V._25 end) as _25 from(select V.Inicio, V.Sem, cast(V._35 as varchar)+'%' as _35,cast(V._36 as varchar)+'%' as _36,cast(V._37 as varchar)+'%' as _37,cast(V._38 as varchar)+'%' as _38,cast(V._39 as varchar)+'%' as _39,cast(V._40 as varchar)+'%' as _40,cast(V._41 as varchar)+'%' as _41,cast(V._42 as varchar)+'%' as _42,cast(V._43 as varchar)+'%' as _43,cast(V._44 as varchar)+'%' as _44,cast(V._45 as varchar)+'%' as _45,cast(V._46 as varchar)+'%' as _46,cast(V._47 as varchar)+'%' as _47,cast(V._48 as varchar)+'%' as _48,cast(V._49 as varchar)+'%' as _49,cast(V._50 as varchar)+'%' as _50,cast(V._51 as varchar)+'%' as _51,cast(V._52 as varchar)+'%' as _52,cast(V._1 as varchar)+'%' as _1,cast(V._2 as varchar)+'%' _2,cast(V._3 as varchar)+'%' as _3,cast(V._4 as varchar)+'%' as _4,cast(V._5 as varchar)+'%' as _5,cast(V._6 as varchar)+'%' as _6,cast(V._7 as varchar)+'%' as _7,cast(V._8 as varchar)+'%' as _8,cast(V._9 as varchar)+'%' as _9,cast(V._10 as varchar)+'%' as _10,cast(V._11 as varchar)+'%' as _11,cast(V._12 as varchar)+'%' as _12,cast(V._13 as varchar)+'%' as _13,cast(V._14 as varchar)+'%' as _14,cast(V._15 as varchar)+'%' as _15,cast(V._16 as varchar)+'%' as _16,cast(V._17 as varchar)+'%' as _17,cast(V._18 as varchar)+'%' as _18,cast(V._19 as varchar)+'%' as _19,cast(V._20 as varchar)+'%'as _20,cast(V._21 as varchar)+'%' as _21,cast(V._22 as varchar)+'%' as _22,cast(V._23 as varchar)+'%' as _23,cast(V._24 as varchar)+'%' as _24,cast(V._25 as varchar)+'%' as _25 from (SELECT V.Inicio, V.Sem,(case when A._35=0 then '' else case when V._35=0 then '' else round((A._35/V._35)*100,0) end end) as _35,(case when A._36=0 then 0 else case when V._36=0 then 0 else round((A._36/V._36)*100,0) end end) as _36,(case when A._37=0 then 0 else case when V._37=0 then 0 else round((A._37/V._37)*100,0) end end) as _37,(case when A._38=0 then 0 else case when V._38=0 then 0 else round((A._38/V._38)*100,0) end end) as _38,(case when A._39=0 then 0 else case when V._39=0 then 0 else round((A._39/V._39)*100,0) end end) as _39,(case when A._40=0 then 0 else case when V._40=0 then 0 else round((A._40/V._40)*100,0) end end) as _40,(case when A._41=0 then 0 else case when V._41=0 then 0 else round((A._41/V._41)*100,0) end end) as _41,(case when A._42=0 then 0 else case when V._42=0 then 0 else round((A._42/V._42)*100,0) end end) as _42,(case when A._43=0 then 0 else case when V._43=0 then 0 else round((A._43/V._43)*100,0) end end) as _43,(case when A._44=0 then 0 else case when V._44=0 then 0 else round((A._44/V._44)*100,0) end end) as _44,(case when A._45=0 then 0 else case when V._45=0 then 0 else round((A._45/V._45)*100,0) end end) as _45,(case when A._46=0 then 0 else case when V._46=0 then 0 else round((A._46/V._46)*100,0) end end) as _46,(case when A._47=0 then 0 else case when V._47=0 then 0 else round((A._47/V._47)*100,0) end end) as _47,(case when A._48=0 then 0 else case when V._48=0 then 0 else round((A._48/V._48)*100,0) end end) as _48,(case when A._49=0 then 0 else case when V._49=0 then 0 else round((A._49/V._49)*100,0) end end) as _49,(case when A._50=0 then 0 else case when V._50=0 then 0 else round((A._50/V._50)*100,0) end end) as _50,(case when A._51=0 then 0 else case when V._51=0 then 0 else round((A._51/V._51)*100,0) end end) as _51,(case when A._52=0 then 0 else case when V._52=0 then 0 else round((A._52/V._52)*100,0) end end) as _52,(case when A._1=0 then 0 else case when V._1=0 then 0 else round((A._1/V._1)*100,0) end end) as _1,(case when A._2=0 then 0 else case when V._2=0 then 0 else round((A._2/V._2)*100,0) end end) as _2,(case when A._3=0 then 0 else case when V._3=0 then 0 else round((A._3/V._3)*100,0) end end) as _3,(case when A._4=0 then 0 else case when V._4=0 then 0 else round((A._4/V._4)*100,0) end end) as _4,(case when A._5=0 then 0 else case when V._5=0 then 0 else round((A._5/V._5)*100,0) end end) as _5,(case when A._6=0 then 0 else case when V._6=0 then 0 else round((A._6/V._6)*100,0) end end) as _6,(case when A._7=0 then 0 else case when V._7=0 then 0 else round((A._7/V._7)*100,0) end end) as _7,(case when A._8=0 then 0 else case when V._8=0 then 0 else round((A._8/V._8)*100,0) end end) as _8,(case when A._9=0 then 0 else case when V._9=0 then 0 else round((A._9/V._9)*100,0) end end) as _9,(case when A._10=0 then 0 else case when V._10=0 then 0 else round((A._10/V._10)*100,0) end end) as _10,(case when A._11=0 then 0 else case when V._11=0 then 0 else round((A._11/V._11)*100,0) end end) as _11,(case when A._12=0 then 0 else case when V._12=0 then 0 else round((A._12/V._12)*100,0) end end) as _12,(case when A._13=0 then 0 else case when V._13=0 then 0 else round((A._13/V._13)*100,0) end end) as _13,(case when A._14=0 then 0 else case when V._14=0 then 0 else round((A._14/V._14)*100,0) end end) as _14,(case when A._15=0 then 0 else case when V._15=0 then 0 else round((A._15/V._15)*100,0) end end) as _15,(case when A._16=0 then 0 else case when V._16=0 then 0 else round((A._16/V._16)*100,0) end end) as _16,(case when A._17=0 then 0 else case when V._17=0 then 0 else round((A._17/V._17)*100,0) end end) as _17,(case when A._18=0 then 0 else case when V._18=0 then 0 else round((A._18/V._18)*100,0) end end) as _18,(case when A._19=0 then 0 else case when V._19=0 then 0 else round((A._19/V._19)*100,0) end end) as _19,(case when A._20=0 then 0 else case when V._20=0 then 0 else round((A._20/V._20)*100,0) end end) as _20,(case when A._21=0 then 0 else case when V._21=0 then 0 else round((A._21/V._21)*100,0) end end) as _21,(case when A._22=0 then 0 else case when V._22=0 then 0 else round((A._22/V._22)*100,0) end end) as _22,(case when A._23=0 then 0 else case when V._23=0 then 0 else round((A._23/V._23)*100,0) end end) as _23,(case when A._24=0 then 0 else case when V._24=0 then 0 else round((A._24/V._24)*100,0) end end) as _24,(case when A._25=0 then 0 else case when V._25=0 then 0 else round((A._25/V._25)*100,0) end end) as _25 FROM(SELECT V.Temporada, isnull(V.[35],0) as _35,isnull(V.[36],0) as _36,isnull(V.[37],0) as _37,isnull(V.[38],0) as _38,isnull(V.[39],0) as _39,isnull(V.[40],0) as _40,isnull(V.[41],0) as _41,isnull(V.[42],0) as _42,isnull(V.[43],0) as _43,isnull(V.[44],0) as _44,isnull(V.[45],0) as _45,isnull(V.[46],0) as _46,isnull(V.[47],0) as _47,isnull(V.[48],0) as _48,isnull(V.[49],0) as _49,isnull(V.[50],0) as _50,isnull(V.[51],0) as _51,isnull(V.[52],0) as _52,isnull(V.[1],0) as _1,isnull(V.[2],0) as _2,isnull(V.[3],0) as _3,isnull(V.[4],0) as _4,isnull(V.[5],0) as _5,isnull(V.[6],0) as _6,isnull(V.[7],0) as _7,isnull(V.[8],0) as _8,isnull(V.[9],0) as _9,isnull(V.[10],0) as _10,isnull(V.[11],0) as _11,isnull(V.[12],0) as _12,isnull(V.[13],0) as _13,isnull(V.[14],0) as _14,isnull(V.[15],0) as _15,isnull(V.[16],0) as _16,isnull(V.[17],0) as _17,isnull(V.[18],0) as _18,isnull(V.[19],0) as _19,isnull(V.[20],0) as _20,isnull(V.[21],0) as _21,isnull(V.[22],0) as _22,isnull(V.[23],0) as _23,isnull(V.[24],0) as _24,isnull(V.[25],0) as _25 FROM(Select * from(SELECT Temporada, Semana, SUM(Convertidas) AS Convertidas FROM SEasonsun1213..UV_ProdRecepcion where CodEstatus<> 'C' and temporada = (select temporada from CatSemanas where GETDATE() between Inicio and Fin) and " +
                "DescProducto = 'FRESA' GROUP BY Temporada, Semana Union All SELECT Temporada, Semana, SUM(Convertidas) AS Convertidas FROM SeasonPlan..UV_ProdRecepcion where CodEstatus <> 'C' and temporada = (select temporada from CatSemanas where GETDATE() between Inicio and Fin) and " +
                "DescProducto = 'FRESA' GROUP BY Temporada, Semana)V PIVOT(SUM(Convertidas) FOR Semana in ([35],[36],[37],[38],[39],[40],[41],[42],[43],[44],[45],[46],[47],[48],[49],[50],[51],[52],[1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],[17],[18],[19],[20],[21],[22],[23],[24],[25]))X)V)A left join(Select V.Inicio, V.Temporada, V.Sem, isnull(V.[35], 0) as _35, isnull(V.[36], 0) as _36, isnull(V.[37], 0) as _37, isnull(V.[38], 0) as _38, isnull(V.[39], 0) as _39, isnull(V.[40], 0) as _40, isnull(V.[41], 0) as _41, isnull(V.[42], 0) as _42, isnull(V.[43], 0) as _43, isnull(V.[44], 0) as _44, isnull(V.[45], 0) as _45, isnull(V.[46], 0) as _46, isnull(V.[47], 0) as _47, isnull(V.[48], 0) as _48, isnull(V.[49], 0) as _49, isnull(V.[50], 0) as _50, isnull(V.[51], 0) as _51, isnull(V.[52], 0) as _52, isnull(V.[1], 0) as _1, isnull(V.[2], 0) as _2, isnull(V.[3], 0) as _3, isnull(V.[4], 0) as _4, isnull(V.[5], 0) as _5, isnull(V.[6], 0) as _6, isnull(V.[7], 0) as _7, isnull(V.[8], 0) as _8, isnull(V.[9], 0) as _9, isnull(V.[10], 0) as _10, isnull(V.[11], 0) as _11, isnull(V.[12], 0) as _12, isnull(V.[13], 0) as _13, isnull(V.[14], 0) as _14, isnull(V.[15], 0) as _15, isnull(V.[16], 0) as _16, isnull(V.[17], 0) as _17, isnull(V.[18], 0) as _18, isnull(V.[19], 0) as _19, isnull(V.[20], 0) as _20, isnull(V.[21], 0) as _21, isnull(V.[22], 0) as _22, isnull(V.[23], 0) as _23, isnull(V.[24], 0) as _24, isnull(V.[25], 0) as _25 from(select * from(SELECT S.Semana AS Sem, S.Inicio, E.Semanas, E.Cantidad, E.Temporada FROM Estimacion_Berries E left join CatSemanas S on E.Temporada = S.Temporada AND E.Semana = S.Semana where E.Temporada = (select temporada from CatSemanas where GETDATE() between Inicio and Fin) and " +
                "E.Concepto = 'PROYECCION SEMANAL' and E.Cultivo = 4)V PIVOT(SUM(Cantidad) For Semanas in ([35],[36],[37],[38],[39],[40],[41],[42],[43],[44],[45],[46],[47],[48],[49],[50],[51],[52],[1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],[17],[18],[19],[20],[21],[22],[23],[24],[25]))X)V)V on A.Temporada = V.Temporada)V)V ORDER BY V.Inicio").ToList();
            return Json(estimacion, JsonRequestBehavior.AllowGet);
        }

        [HttpPost]
        public ActionResult Fresa(int[] array)
        {
            JsonResult dtaEjecucionTarea = default(JsonResult);
            try
            {
                if (Update_Fresa(array))
                {
                    dtaEjecucionTarea = Json(new
                    {
                        rstProceso = "true",
                        MessageGestion = "Cambios guardados con éxito"
                    });
                }
                else
                {
                    dtaEjecucionTarea = Json(new
                    {
                        rstProceso = "false",
                        MessageGestion = "Error, algo salió mal, intente de nuevo"
                    });
                }
            }
            catch (Exception e)
            {
                e.ToString();
            }
            return dtaEjecucionTarea;
        }

        public bool Update_Fresa(int[] array)
        {
            if (array != null)
            {
                var sem = array[0];
                var _37 = array[1];
                var _38 = array[2];
                var _39 = array[3];
                var _40 = array[4];
                var _41 = array[5];
                var _42 = array[6];
                var _43 = array[7];
                var _44 = array[8];
                var _45 = array[9];
                var _46 = array[10];
                var _47 = array[11];
                var _48 = array[12];
                var _49 = array[13];
                var _50 = array[14];
                var _51 = array[15];
                var _52 = array[16];
                var _1 = array[17];
                var _2 = array[18];
                var _3 = array[19];
                var _4 = array[20];
                var _5 = array[21];
                var _6 = array[22];
                var _7 = array[23];
                var _8 = array[24];
                var _9 = array[25];
                var _10 = array[26];
                var _11 = array[27];
                var _12 = array[28];
                var _13 = array[29];
                var _14 = array[30];
                var _15 = array[31];
                var _16 = array[32];
                var _17 = array[33];
                var _18 = array[34];
                var _19 = array[35];
                var _20 = array[36];
                var _21 = array[37];
                var _22 = array[38];
                var _23 = array[39];
                var _24 = array[40];
                var _25 = array[41];

                var item_37 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 37 && x.Cultivo == 4 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_37.Cantidad = _37;
                item_37.Fecha = DateTime.Now;

                var item_38 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 38 && x.Cultivo == 4 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_38.Cantidad = _38;
                item_38.Fecha = DateTime.Now;

                var item_39 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 39 && x.Cultivo == 4 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_39.Cantidad = _39;
                item_39.Fecha = DateTime.Now;

                var item_40 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 40 && x.Cultivo == 4 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_40.Cantidad = _40;
                item_40.Fecha = DateTime.Now;

                var item_41 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 41 && x.Cultivo == 4 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_41.Cantidad = _41;
                item_41.Fecha = DateTime.Now;

                var item_42 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 42 && x.Cultivo == 4 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_42.Cantidad = _42;
                item_42.Fecha = DateTime.Now;

                var item_43 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 43 && x.Cultivo == 4 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_43.Cantidad = _43;
                item_43.Fecha = DateTime.Now;

                var item_44 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 44 && x.Cultivo == 4 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_44.Cantidad = _44;
                item_44.Fecha = DateTime.Now;

                var item_45 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 45 && x.Cultivo == 4 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_45.Cantidad = _45;
                item_45.Fecha = DateTime.Now;

                var item_46 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 46 && x.Cultivo == 4 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_46.Cantidad = _46;
                item_46.Fecha = DateTime.Now;

                var item_47 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 47 && x.Cultivo == 4 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_47.Cantidad = _47;
                item_47.Fecha = DateTime.Now;

                var item_48 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 48 && x.Cultivo == 4 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_48.Cantidad = _48;
                item_48.Fecha = DateTime.Now;

                var item_49 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 49 && x.Cultivo == 4 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_49.Cantidad = _49;
                item_49.Fecha = DateTime.Now;

                var item_50 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 50 && x.Cultivo == 4 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_50.Cantidad = _50;
                item_50.Fecha = DateTime.Now;

                var item_51 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 51 && x.Cultivo == 4 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_51.Cantidad = _51;
                item_51.Fecha = DateTime.Now;

                var item_52 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 52 && x.Cultivo == 4 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_52.Cantidad = _52;
                item_52.Fecha = DateTime.Now;

                var item_1 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 1 && x.Cultivo == 4 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_1.Cantidad = _1;
                item_1.Fecha = DateTime.Now;

                var item_2 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 2 && x.Cultivo == 4 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_2.Cantidad = _2;
                item_2.Fecha = DateTime.Now;

                var item_3 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 3 && x.Cultivo == 4 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_3.Cantidad = _3;
                item_3.Fecha = DateTime.Now;

                var item_4 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 4 && x.Cultivo == 4 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_4.Cantidad = _4;
                item_4.Fecha = DateTime.Now;

                var item_5 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 5 && x.Cultivo == 4 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_5.Cantidad = _5;
                item_5.Fecha = DateTime.Now;

                var item_6 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 6 && x.Cultivo == 4 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_6.Cantidad = _6;
                item_6.Fecha = DateTime.Now;

                var item_7 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 7 && x.Cultivo == 4 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_7.Cantidad = _7;
                item_7.Fecha = DateTime.Now;

                var item_8 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 8 && x.Cultivo == 4 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_8.Cantidad = _8;
                item_8.Fecha = DateTime.Now;

                var item_9 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 9 && x.Cultivo == 4 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_9.Cantidad = _9;
                item_9.Fecha = DateTime.Now;

                var item_10 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 10 && x.Cultivo == 4 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_10.Cantidad = _10;
                item_10.Fecha = DateTime.Now;

                var item_11 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 11 && x.Cultivo == 4 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_11.Cantidad = _11;
                item_11.Fecha = DateTime.Now;

                var item_12 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 12 && x.Cultivo == 4 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_12.Cantidad = _12;
                item_12.Fecha = DateTime.Now;

                var item_13 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 13 && x.Cultivo == 4 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_13.Cantidad = _13;
                item_13.Fecha = DateTime.Now;

                var item_14 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 14 && x.Cultivo == 4 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_14.Cantidad = _14;
                item_14.Fecha = DateTime.Now;

                var item_15 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 15 && x.Cultivo == 4 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_15.Cantidad = _15;
                item_15.Fecha = DateTime.Now;

                var item_16 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 16 && x.Cultivo == 4 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_16.Cantidad = _16;
                item_16.Fecha = DateTime.Now;

                var item_17 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 17 && x.Cultivo == 4 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_17.Cantidad = _17;
                item_17.Fecha = DateTime.Now;

                var item_18 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 18 && x.Cultivo == 4 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_18.Cantidad = _18;
                item_18.Fecha = DateTime.Now;

                var item_19 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 19 && x.Cultivo == 4 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_19.Cantidad = _19;
                item_19.Fecha = DateTime.Now;

                var item_20 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 20 && x.Cultivo == 4 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_20.Cantidad = _20;
                item_20.Fecha = DateTime.Now;

                var item_21 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 21 && x.Cultivo == 4 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_21.Cantidad = _21;
                item_21.Fecha = DateTime.Now;

                var item_22 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 22 && x.Cultivo == 4 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_22.Cantidad = _22;
                item_22.Fecha = DateTime.Now;

                var item_23 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 23 && x.Cultivo == 4 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_23.Cantidad = _23;
                item_23.Fecha = DateTime.Now;

                var item_24 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 24 && x.Cultivo == 4 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_24.Cantidad = _24;
                item_24.Fecha = DateTime.Now;

                var item_25 = bd.Estimacion_Berries.Where(x => x.Semana == sem && x.Semanas == 25 && x.Cultivo == 4 && x.Concepto == "PROYECCION SEMANAL").FirstOrDefault();
                item_25.Cantidad = _25;
                item_25.Fecha = DateTime.Now;

                bd.SaveChanges();
                return true;
            }
            else
            {
                return false;
            }
        }

        //EXCEL
        public void VARIACION_ESTIMACION_BERRIES1920()
        {
            ViewData["Nombre"] = Session["Nombre"].ToString();

            ExcelPackage excel = new ExcelPackage();

            //ZARZAMORA
            ExcelWorksheet ws = excel.Workbook.Worksheets.Add("ZARZAMORA");
            ws.Cells["A1"].Value = "VARIACION SEMANAL DE LA ESTIMACION DE ZARZAMORA";
            ws.Cells["A2"].Value = "SEMANA";
            ws.Cells["B2"].Value = "37";
            ws.Cells["C2"].Value = "38";
            ws.Cells["D2"].Value = "39";
            ws.Cells["E2"].Value = "40";
            ws.Cells["F2"].Value = "41";
            ws.Cells["G2"].Value = "42";
            ws.Cells["H2"].Value = "43";
            ws.Cells["I2"].Value = "44";
            ws.Cells["J2"].Value = "45";
            ws.Cells["K2"].Value = "46";
            ws.Cells["L2"].Value = "47";
            ws.Cells["M2"].Value = "48";
            ws.Cells["N2"].Value = "49";
            ws.Cells["O2"].Value = "50";
            ws.Cells["P2"].Value = "51";
            ws.Cells["Q2"].Value = "52";
            ws.Cells["R2"].Value = "1";
            ws.Cells["S2"].Value = "2";
            ws.Cells["T2"].Value = "3";
            ws.Cells["U2"].Value = "4";
            ws.Cells["V2"].Value = "5";
            ws.Cells["W2"].Value = "6";
            ws.Cells["X2"].Value = "7";
            ws.Cells["Y2"].Value = "8";
            ws.Cells["Z2"].Value = "9";
            ws.Cells["AA2"].Value = "10";
            ws.Cells["AB2"].Value = "11";
            ws.Cells["AC2"].Value = "12";
            ws.Cells["AD2"].Value = "13";
            ws.Cells["AE2"].Value = "14";
            ws.Cells["AF2"].Value = "15";
            ws.Cells["AG2"].Value = "16";
            ws.Cells["AH2"].Value = "17";
            ws.Cells["AI2"].Value = "18";
            ws.Cells["AJ2"].Value = "19";
            ws.Cells["AK2"].Value = "20";
            ws.Cells["AL2"].Value = "21";
            ws.Cells["AM2"].Value = "22";
            ws.Cells["AN2"].Value = "23";
            ws.Cells["AO2"].Value = "24";
            ws.Cells["AP2"].Value = "25";

            //recepcion real
            var recepcion_realZ = bd.Database.SqlQuery<R_Berries>("SELECT 'TOTAL' as 'SEMANA', round(isnull(V.[37],0),0) as _37,round(isnull(V.[38],0),0) as _38,round(isnull(V.[39],0),0) as _39,round(isnull(V.[40],0),0) as _40,round(isnull(V.[41],0),0) as _41,round(isnull(V.[42],0),0) as _42,round(isnull(V.[43],0),0) as _43,round(isnull(V.[44],0),0) as _44,round(isnull(V.[45],0),0) as _45,round(isnull(V.[46],0),0) as _46,round(isnull(V.[47],0),0) as _47,round(isnull(V.[48],0),0) as _48,round(isnull(V.[49],0),0) as _49,round(isnull(V.[50],0),0) as _50,round(isnull(V.[51],0),0) as _51,round(isnull(V.[52],0),0) as _52,round(isnull(V.[1],0),0) as _1,round(isnull(V.[2],0),0) as _2,round(isnull(V.[3],0),0) as _3,round(isnull(V.[4],0),0) as _4,round(isnull(V.[5],0),0) as _5,round(isnull(V.[6],0),0) as _6,round(isnull(V.[7],0),0) as _7,round(isnull(V.[8],0),0) as _8,round(isnull(V.[9],0),0) as _9,round(isnull(V.[10],0),0) as _10,round(isnull(V.[11],0),0) as _11,round(isnull(V.[12],0),0) as _12,round(isnull(V.[13],0),0) as _13,round(isnull(V.[14],0),0) as _14,round(isnull(V.[15],0),0) as _15,round(isnull(V.[16],0),0) as _16,round(isnull(V.[17],0),0) as _17,round(isnull(V.[18],0),0) as _18,round(isnull(V.[19],0),0) as _19,round(isnull(V.[20],0),0) as _20,round(isnull(V.[21],0),0) as _21,round(isnull(V.[22],0),0) as _22,round(isnull(V.[23],0),0) as _23,round(isnull(V.[24],0),0) as _24,round(isnull(V.[25],0),0) as _25 " +
               "FROM(Select * from(SELECT Semana, SUM(Convertidas) AS Convertidas FROM SEasonsun1213..UV_ProdRecepcion where CodEstatus <> 'C' and temporada = (select temporada from CatSemanas where GETDATE() between Inicio and Fin) and DescProducto = 'ZARZAMORA' GROUP BY Semana Union All " +
               "SELECT Semana, SUM(Convertidas) AS Convertidas FROM SeasonPlan..UV_ProdRecepcion where CodEstatus <> 'C' and temporada = (select temporada from CatSemanas where GETDATE() between Inicio and Fin) and DescProducto = 'ZARZAMORA' GROUP BY Semana" +
               ")V PIVOT(SUM(Convertidas) FOR Semana in ([37],[38],[39],[40],[41],[42],[43],[44],[45],[46],[47],[48],[49],[50],[51],[52],[1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],[17],[18],[19],[20],[21],[22],[23],[24],[25]))X)V").ToList();

            int xRZ = 3;
            foreach (var item in recepcion_realZ)
            {
                ws.Cells[string.Format("A{0}", xRZ)].Value = "RECEPCION REAL";
                ws.Cells[string.Format("B{0}", xRZ)].Value = item._37;
                ws.Cells[string.Format("C{0}", xRZ)].Value = item._38;
                ws.Cells[string.Format("D{0}", xRZ)].Value = item._39;
                ws.Cells[string.Format("E{0}", xRZ)].Value = item._40;
                ws.Cells[string.Format("F{0}", xRZ)].Value = item._41;
                ws.Cells[string.Format("G{0}", xRZ)].Value = item._42;
                ws.Cells[string.Format("H{0}", xRZ)].Value = item._43;
                ws.Cells[string.Format("I{0}", xRZ)].Value = item._44;
                ws.Cells[string.Format("J{0}", xRZ)].Value = item._45;
                ws.Cells[string.Format("K{0}", xRZ)].Value = item._46;
                ws.Cells[string.Format("L{0}", xRZ)].Value = item._47;
                ws.Cells[string.Format("M{0}", xRZ)].Value = item._48;
                ws.Cells[string.Format("N{0}", xRZ)].Value = item._49;
                ws.Cells[string.Format("O{0}", xRZ)].Value = item._50;
                ws.Cells[string.Format("P{0}", xRZ)].Value = item._51;
                ws.Cells[string.Format("Q{0}", xRZ)].Value = item._52;
                ws.Cells[string.Format("R{0}", xRZ)].Value = item._1;
                ws.Cells[string.Format("S{0}", xRZ)].Value = item._2;
                ws.Cells[string.Format("T{0}", xRZ)].Value = item._3;
                ws.Cells[string.Format("U{0}", xRZ)].Value = item._4;
                ws.Cells[string.Format("V{0}", xRZ)].Value = item._5;
                ws.Cells[string.Format("W{0}", xRZ)].Value = item._6;
                ws.Cells[string.Format("X{0}", xRZ)].Value = item._7;
                ws.Cells[string.Format("Y{0}", xRZ)].Value = item._8;
                ws.Cells[string.Format("Z{0}", xRZ)].Value = item._9;
                ws.Cells[string.Format("AA{0}", xRZ)].Value = item._10;
                ws.Cells[string.Format("AB{0}", xRZ)].Value = item._11;
                ws.Cells[string.Format("AC{0}", xRZ)].Value = item._12;
                ws.Cells[string.Format("AD{0}", xRZ)].Value = item._13;
                ws.Cells[string.Format("AE{0}", xRZ)].Value = item._14;
                ws.Cells[string.Format("AF{0}", xRZ)].Value = item._15;
                ws.Cells[string.Format("AG{0}", xRZ)].Value = item._16;
                ws.Cells[string.Format("AH{0}", xRZ)].Value = item._17;
                ws.Cells[string.Format("AI{0}", xRZ)].Value = item._18;
                ws.Cells[string.Format("AJ{0}", xRZ)].Value = item._19;
                ws.Cells[string.Format("AK{0}", xRZ)].Value = item._20;
                ws.Cells[string.Format("AL{0}", xRZ)].Value = item._21;
                ws.Cells[string.Format("AM{0}", xRZ)].Value = item._22;
                ws.Cells[string.Format("AN{0}", xRZ)].Value = item._23;
                ws.Cells[string.Format("AO{0}", xRZ)].Value = item._24;
                ws.Cells[string.Format("AP{0}", xRZ)].Value = item._25;
                xRZ++;
            }

            //recepcion europa
            var recepcion_europaZ = bd.Database.SqlQuery<E_Berries>("Select round(isnull(V.[37],0),0) as _37,round(isnull(V.[38],0),0) as _38,round(isnull(V.[39],0),0) as _39,round(isnull(V.[40],0),0) as _40,round(isnull(V.[41],0),0) as _41,round(isnull(V.[42],0),0) as _42,round(isnull(V.[43],0),0) as _43,round(isnull(V.[44],0),0) as _44,round(isnull(V.[45],0),0) as _45,round(isnull(V.[46],0),0) as _46,round(isnull(V.[47],0),0) as _47,round(isnull(V.[48],0),0) as _48,round(isnull(V.[49],0),0) as _49,round(isnull(V.[50],0),0) as _50,round(isnull(V.[51],0),0) as _51,round(isnull(V.[52],0),0) as _52,round(isnull(V.[1],0),0) as _1,round(isnull(V.[2],0),0) as _2,round(isnull(V.[3],0),0) as _3,round(isnull(V.[4],0),0) as _4,round(isnull(V.[5],0),0) as _5,round(isnull(V.[6],0),0) as _6,round(isnull(V.[7],0),0) as _7,round(isnull(V.[8],0),0) as _8,round(isnull(V.[9],0),0) as _9,round(isnull(V.[10],0),0) as _10,round(isnull(V.[11],0),0) as _11,round(isnull(V.[12],0),0) as _12,round(isnull(V.[13],0),0) as _13,round(isnull(V.[14],0),0) as _14,round(isnull(V.[15],0),0) as _15,round(isnull(V.[16],0),0) as _16,round(isnull(V.[17],0),0) as _17,round(isnull(V.[18],0),0) as _18,round(isnull(V.[19],0),0) as _19,round(isnull(V.[20],0),0) as _20,round(isnull(V.[21],0),0) as _21,round(isnull(V.[22],0),0) as _22,round(isnull(V.[23],0),0) as _23,round(isnull(V.[24],0),0) as _24,round(isnull(V.[25],0),0) as _25 from (Select * from(SELECT E.Concepto, S.Inicio, E.Semanas, E.Cantidad FROM Estimacion_Berries E left join CatSemanas S on E.Temporada=S.Temporada AND E.Semana=S.Semana where E.Temporada=(select temporada from CatSemanas where GETDATE() between Inicio and Fin) and " +
                "E.Concepto = 'EUROPA' and E.Cultivo = 1)V PIVOT(SUM(Cantidad) For Semanas in ([37],[38],[39],[40],[41],[42],[43],[44],[45],[46],[47],[48],[49],[50],[51],[52],[1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],[17],[18],[19],[20],[21],[22],[23],[24],[25]))X)V ORDER BY V.Inicio").ToList();

            int xEZ = 4;
            foreach (var item in recepcion_europaZ)
            {
                ws.Cells[string.Format("A{0}", xEZ)].Value = "RECEPCION EUROPA";
                ws.Cells[string.Format("B{0}", xEZ)].Value = item._37;
                ws.Cells[string.Format("C{0}", xEZ)].Value = item._38;
                ws.Cells[string.Format("D{0}", xEZ)].Value = item._39;
                ws.Cells[string.Format("E{0}", xEZ)].Value = item._40;
                ws.Cells[string.Format("F{0}", xEZ)].Value = item._41;
                ws.Cells[string.Format("G{0}", xEZ)].Value = item._42;
                ws.Cells[string.Format("H{0}", xEZ)].Value = item._43;
                ws.Cells[string.Format("I{0}", xEZ)].Value = item._44;
                ws.Cells[string.Format("J{0}", xEZ)].Value = item._45;
                ws.Cells[string.Format("K{0}", xEZ)].Value = item._46;
                ws.Cells[string.Format("L{0}", xEZ)].Value = item._47;
                ws.Cells[string.Format("M{0}", xEZ)].Value = item._48;
                ws.Cells[string.Format("N{0}", xEZ)].Value = item._49;
                ws.Cells[string.Format("O{0}", xEZ)].Value = item._50;
                ws.Cells[string.Format("P{0}", xEZ)].Value = item._51;
                ws.Cells[string.Format("Q{0}", xEZ)].Value = item._52;
                ws.Cells[string.Format("R{0}", xEZ)].Value = item._1;
                ws.Cells[string.Format("S{0}", xEZ)].Value = item._2;
                ws.Cells[string.Format("T{0}", xEZ)].Value = item._3;
                ws.Cells[string.Format("U{0}", xEZ)].Value = item._4;
                ws.Cells[string.Format("V{0}", xEZ)].Value = item._5;
                ws.Cells[string.Format("W{0}", xEZ)].Value = item._6;
                ws.Cells[string.Format("X{0}", xEZ)].Value = item._7;
                ws.Cells[string.Format("Y{0}", xEZ)].Value = item._8;
                ws.Cells[string.Format("Z{0}", xEZ)].Value = item._9;
                ws.Cells[string.Format("AA{0}", xEZ)].Value = item._10;
                ws.Cells[string.Format("AB{0}", xEZ)].Value = item._11;
                ws.Cells[string.Format("AC{0}", xEZ)].Value = item._12;
                ws.Cells[string.Format("AD{0}", xEZ)].Value = item._13;
                ws.Cells[string.Format("AE{0}", xEZ)].Value = item._14;
                ws.Cells[string.Format("AF{0}", xEZ)].Value = item._15;
                ws.Cells[string.Format("AG{0}", xEZ)].Value = item._16;
                ws.Cells[string.Format("AH{0}", xEZ)].Value = item._17;
                ws.Cells[string.Format("AI{0}", xEZ)].Value = item._18;
                ws.Cells[string.Format("AJ{0}", xEZ)].Value = item._19;
                ws.Cells[string.Format("AK{0}", xEZ)].Value = item._20;
                ws.Cells[string.Format("AL{0}", xEZ)].Value = item._21;
                ws.Cells[string.Format("AM{0}", xEZ)].Value = item._22;
                ws.Cells[string.Format("AN{0}", xEZ)].Value = item._23;
                ws.Cells[string.Format("AO{0}", xEZ)].Value = item._24;
                ws.Cells[string.Format("AP{0}", xEZ)].Value = item._25;
                xEZ++;
            }

            //ALWAYS FRESH
            var always_freshZ = bd.Database.SqlQuery<E_Berries>("Select round(isnull(V.[37],0),0) as _37,round(isnull(V.[38],0),0) as _38,round(isnull(V.[39],0),0) as _39,round(isnull(V.[40],0),0) as _40,round(isnull(V.[41],0),0) as _41,round(isnull(V.[42],0),0) as _42,round(isnull(V.[43],0),0) as _43,round(isnull(V.[44],0),0) as _44,round(isnull(V.[45],0),0) as _45,round(isnull(V.[46],0),0) as _46,round(isnull(V.[47],0),0) as _47,round(isnull(V.[48],0),0) as _48,round(isnull(V.[49],0),0) as _49,round(isnull(V.[50],0),0) as _50,round(isnull(V.[51],0),0) as _51,round(isnull(V.[52],0),0) as _52,round(isnull(V.[1],0),0) as _1,round(isnull(V.[2],0),0) as _2,round(isnull(V.[3],0),0) as _3,round(isnull(V.[4],0),0) as _4,round(isnull(V.[5],0),0) as _5,round(isnull(V.[6],0),0) as _6,round(isnull(V.[7],0),0) as _7,round(isnull(V.[8],0),0) as _8,round(isnull(V.[9],0),0) as _9,round(isnull(V.[10],0),0) as _10,round(isnull(V.[11],0),0) as _11,round(isnull(V.[12],0),0) as _12,round(isnull(V.[13],0),0) as _13,round(isnull(V.[14],0),0) as _14,round(isnull(V.[15],0),0) as _15,round(isnull(V.[16],0),0) as _16,round(isnull(V.[17],0),0) as _17,round(isnull(V.[18],0),0) as _18,round(isnull(V.[19],0),0) as _19,round(isnull(V.[20],0),0) as _20,round(isnull(V.[21],0),0) as _21,round(isnull(V.[22],0),0) as _22,round(isnull(V.[23],0),0) as _23,round(isnull(V.[24],0),0) as _24,round(isnull(V.[25],0),0) as _25 from(Select * from(SELECT S.Inicio, R.Semana, (R.Cantidad-C.Cantidad) Cantidad FROM(SELECT Semana, SUM(Convertidas) AS Cantidad, Temporada FROM SEasonsun1213..UV_ProdRecepcion where CodEstatus<> 'C' and temporada = (select temporada from CatSemanas where GETDATE() between Inicio and Fin) and " +
                "DescProducto = 'ZARZAMORA' GROUP BY Semana, Temporada Union All SELECT Semana, SUM(Convertidas) AS Cantidad, Temporada FROM SeasonPlan..UV_ProdRecepcion where CodEstatus <> 'C' and temporada = (select temporada from CatSemanas where GETDATE() between Inicio and Fin) and " +
                "DescProducto = 'ZARZAMORA' GROUP BY Semana,Temporada)R left join(Select Semanas, Cantidad, Temporada, Cultivo, Semana from Estimacion_Berries " +
                "where Concepto = 'EUROPA' and Cultivo = 1 and Temporada = (select temporada from CatSemanas where GETDATE() between Inicio and Fin))C on R.Temporada = C.Temporada and R.Semana = C.Semanas left join CatSemanas S on R.Temporada = S.Temporada AND C.Semana = S.Semana)V PIVOT(SUM(Cantidad) For Semana in ([37],[38],[39],[40],[41],[42],[43],[44],[45],[46],[47],[48],[49],[50],[51],[52],[1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],[17],[18],[19],[20],[21],[22],[23],[24],[25]))X)V ORDER BY V.Inicio").ToList();

            int xAFZ = 5;
            foreach (var item in always_freshZ)
            {
                ws.Cells[string.Format("A{0}", xAFZ)].Value = "ALWAYS FRESH";
                ws.Cells[string.Format("B{0}", xAFZ)].Value = item._37;
                ws.Cells[string.Format("C{0}", xAFZ)].Value = item._38;
                ws.Cells[string.Format("D{0}", xAFZ)].Value = item._39;
                ws.Cells[string.Format("E{0}", xAFZ)].Value = item._40;
                ws.Cells[string.Format("F{0}", xAFZ)].Value = item._41;
                ws.Cells[string.Format("G{0}", xAFZ)].Value = item._42;
                ws.Cells[string.Format("H{0}", xAFZ)].Value = item._43;
                ws.Cells[string.Format("I{0}", xAFZ)].Value = item._44;
                ws.Cells[string.Format("J{0}", xAFZ)].Value = item._45;
                ws.Cells[string.Format("K{0}", xAFZ)].Value = item._46;
                ws.Cells[string.Format("L{0}", xAFZ)].Value = item._47;
                ws.Cells[string.Format("M{0}", xAFZ)].Value = item._48;
                ws.Cells[string.Format("N{0}", xAFZ)].Value = item._49;
                ws.Cells[string.Format("O{0}", xAFZ)].Value = item._50;
                ws.Cells[string.Format("P{0}", xAFZ)].Value = item._51;
                ws.Cells[string.Format("Q{0}", xAFZ)].Value = item._52;
                ws.Cells[string.Format("R{0}", xAFZ)].Value = item._1;
                ws.Cells[string.Format("S{0}", xAFZ)].Value = item._2;
                ws.Cells[string.Format("T{0}", xAFZ)].Value = item._3;
                ws.Cells[string.Format("U{0}", xAFZ)].Value = item._4;
                ws.Cells[string.Format("V{0}", xAFZ)].Value = item._5;
                ws.Cells[string.Format("W{0}", xAFZ)].Value = item._6;
                ws.Cells[string.Format("X{0}", xAFZ)].Value = item._7;
                ws.Cells[string.Format("Y{0}", xAFZ)].Value = item._8;
                ws.Cells[string.Format("Z{0}", xAFZ)].Value = item._9;
                ws.Cells[string.Format("AA{0}", xAFZ)].Value = item._10;
                ws.Cells[string.Format("AB{0}", xAFZ)].Value = item._11;
                ws.Cells[string.Format("AC{0}", xAFZ)].Value = item._12;
                ws.Cells[string.Format("AD{0}", xAFZ)].Value = item._13;
                ws.Cells[string.Format("AE{0}", xAFZ)].Value = item._14;
                ws.Cells[string.Format("AF{0}", xAFZ)].Value = item._15;
                ws.Cells[string.Format("AG{0}", xAFZ)].Value = item._16;
                ws.Cells[string.Format("AH{0}", xAFZ)].Value = item._17;
                ws.Cells[string.Format("AI{0}", xAFZ)].Value = item._18;
                ws.Cells[string.Format("AJ{0}", xAFZ)].Value = item._19;
                ws.Cells[string.Format("AK{0}", xAFZ)].Value = item._20;
                ws.Cells[string.Format("AL{0}", xAFZ)].Value = item._21;
                ws.Cells[string.Format("AM{0}", xAFZ)].Value = item._22;
                ws.Cells[string.Format("AN{0}", xAFZ)].Value = item._23;
                ws.Cells[string.Format("AO{0}", xAFZ)].Value = item._24;
                ws.Cells[string.Format("AP{0}", xAFZ)].Value = item._25;
                xAFZ++;
            }

            //COYOTES
            var coyotesZ = bd.Database.SqlQuery<E_Berries>("Select V.Concepto, round(isnull(V.[37],0),0) as _37,round(isnull(V.[38],0),0) as _38,round(isnull(V.[39],0),0) as _39,round(isnull(V.[40],0),0) as _40,round(isnull(V.[41],0),0) as _41,round(isnull(V.[42],0),0) as _42,round(isnull(V.[43],0),0) as _43,round(isnull(V.[44],0),0) as _44,round(isnull(V.[45],0),0) as _45,round(isnull(V.[46],0),0) as _46,round(isnull(V.[47],0),0) as _47,round(isnull(V.[48],0),0) as _48,round(isnull(V.[49],0),0) as _49,round(isnull(V.[50],0),0) as _50,round(isnull(V.[51],0),0) as _51,round(isnull(V.[52],0),0) as _52,round(isnull(V.[1],0),0) as _1,round(isnull(V.[2],0),0) as _2,round(isnull(V.[3],0),0) as _3,round(isnull(V.[4],0),0) as _4,round(isnull(V.[5],0),0) as _5,round(isnull(V.[6],0),0) as _6,round(isnull(V.[7],0),0) as _7,round(isnull(V.[8],0),0) as _8,round(isnull(V.[9],0),0) as _9,round(isnull(V.[10],0),0) as _10,round(isnull(V.[11],0),0) as _11,round(isnull(V.[12],0),0) as _12,round(isnull(V.[13],0),0) as _13,round(isnull(V.[14],0),0) as _14,round(isnull(V.[15],0),0) as _15,round(isnull(V.[16],0),0) as _16,round(isnull(V.[17],0),0) as _17,round(isnull(V.[18],0),0) as _18,round(isnull(V.[19],0),0) as _19,round(isnull(V.[20],0),0) as _20,round(isnull(V.[21],0),0) as _21,round(isnull(V.[22],0),0) as _22,round(isnull(V.[23],0),0) as _23,round(isnull(V.[24],0),0) as _24,round(isnull(V.[25],0),0) as _25 " +
                "from(Select * from(SELECT E.Concepto, S.Inicio, E.Semanas, E.Cantidad " +
                "FROM Estimacion_Berries E left join CatSemanas S on E.Temporada = S.Temporada AND E.Semana = S.Semana " +
                "where E.Temporada = (select temporada from CatSemanas where GETDATE() between Inicio and Fin) and E.Concepto = 'COYOTES' and E.Cultivo = 1)V " +
                "PIVOT(SUM(Cantidad) For Semanas in ([37],[38],[39],[40],[41],[42],[43],[44],[45],[46],[47],[48],[49],[50],[51],[52],[1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],[17],[18],[19],[20],[21],[22],[23],[24],[25]))X)V ORDER BY V.Inicio").ToList();

            int xCZ = 6;
            foreach (var item in coyotesZ)
            {
                ws.Cells[string.Format("A{0}", xCZ)].Value = "COYOTES";
                ws.Cells[string.Format("B{0}", xCZ)].Value = item._37;
                ws.Cells[string.Format("C{0}", xCZ)].Value = item._38;
                ws.Cells[string.Format("D{0}", xCZ)].Value = item._39;
                ws.Cells[string.Format("E{0}", xCZ)].Value = item._40;
                ws.Cells[string.Format("F{0}", xCZ)].Value = item._41;
                ws.Cells[string.Format("G{0}", xCZ)].Value = item._42;
                ws.Cells[string.Format("H{0}", xCZ)].Value = item._43;
                ws.Cells[string.Format("I{0}", xCZ)].Value = item._44;
                ws.Cells[string.Format("J{0}", xCZ)].Value = item._45;
                ws.Cells[string.Format("K{0}", xCZ)].Value = item._46;
                ws.Cells[string.Format("L{0}", xCZ)].Value = item._47;
                ws.Cells[string.Format("M{0}", xCZ)].Value = item._48;
                ws.Cells[string.Format("N{0}", xCZ)].Value = item._49;
                ws.Cells[string.Format("O{0}", xCZ)].Value = item._50;
                ws.Cells[string.Format("P{0}", xCZ)].Value = item._51;
                ws.Cells[string.Format("Q{0}", xCZ)].Value = item._52;
                ws.Cells[string.Format("R{0}", xCZ)].Value = item._1;
                ws.Cells[string.Format("S{0}", xCZ)].Value = item._2;
                ws.Cells[string.Format("T{0}", xCZ)].Value = item._3;
                ws.Cells[string.Format("U{0}", xCZ)].Value = item._4;
                ws.Cells[string.Format("V{0}", xCZ)].Value = item._5;
                ws.Cells[string.Format("W{0}", xCZ)].Value = item._6;
                ws.Cells[string.Format("X{0}", xCZ)].Value = item._7;
                ws.Cells[string.Format("Y{0}", xCZ)].Value = item._8;
                ws.Cells[string.Format("Z{0}", xCZ)].Value = item._9;
                ws.Cells[string.Format("AA{0}", xCZ)].Value = item._10;
                ws.Cells[string.Format("AB{0}", xCZ)].Value = item._11;
                ws.Cells[string.Format("AC{0}", xCZ)].Value = item._12;
                ws.Cells[string.Format("AD{0}", xCZ)].Value = item._13;
                ws.Cells[string.Format("AE{0}", xCZ)].Value = item._14;
                ws.Cells[string.Format("AF{0}", xCZ)].Value = item._15;
                ws.Cells[string.Format("AG{0}", xCZ)].Value = item._16;
                ws.Cells[string.Format("AH{0}", xCZ)].Value = item._17;
                ws.Cells[string.Format("AI{0}", xCZ)].Value = item._18;
                ws.Cells[string.Format("AJ{0}", xCZ)].Value = item._19;
                ws.Cells[string.Format("AK{0}", xCZ)].Value = item._20;
                ws.Cells[string.Format("AL{0}", xCZ)].Value = item._21;
                ws.Cells[string.Format("AM{0}", xCZ)].Value = item._22;
                ws.Cells[string.Format("AN{0}", xCZ)].Value = item._23;
                ws.Cells[string.Format("AO{0}", xCZ)].Value = item._24;
                ws.Cells[string.Format("AP{0}", xCZ)].Value = item._25;
                xCZ++;
            }

            //ALWAYS FRESH SIN COYOTES
            var always_freshSCZ = bd.Database.SqlQuery<E_Berries>("Select round(isnull(V.[37],0),0) as _37,round(isnull(V.[38],0),0) as _38,round(isnull(V.[39],0),0) as _39,round(isnull(V.[40],0),0) as _40,round(isnull(V.[41],0),0) as _41,round(isnull(V.[42],0),0) as _42,round(isnull(V.[43],0),0) as _43,round(isnull(V.[44],0),0) as _44,round(isnull(V.[45],0),0) as _45,round(isnull(V.[46],0),0) as _46,round(isnull(V.[47],0),0) as _47,round(isnull(V.[48],0),0) as _48,round(isnull(V.[49],0),0) as _49,round(isnull(V.[50],0),0) as _50,round(isnull(V.[51],0),0) as _51,round(isnull(V.[52],0),0) as _52,round(isnull(V.[1],0),0) as _1,round(isnull(V.[2],0),0) as _2,round(isnull(V.[3],0),0) as _3,round(isnull(V.[4],0),0) as _4,round(isnull(V.[5],0),0) as _5,round(isnull(V.[6],0),0) as _6,round(isnull(V.[7],0),0) as _7,round(isnull(V.[8],0),0) as _8,round(isnull(V.[9],0),0) as _9,round(isnull(V.[10],0),0) as _10,round(isnull(V.[11],0),0) as _11,round(isnull(V.[12],0),0) as _12,round(isnull(V.[13],0),0) as _13,round(isnull(V.[14],0),0) as _14,round(isnull(V.[15],0),0) as _15,round(isnull(V.[16],0),0) as _16,round(isnull(V.[17],0),0) as _17,round(isnull(V.[18],0),0) as _18,round(isnull(V.[19],0),0) as _19,round(isnull(V.[20],0),0) as _20,round(isnull(V.[21],0),0) as _21,round(isnull(V.[22],0),0) as _22,round(isnull(V.[23],0),0) as _23,round(isnull(V.[24],0),0) as _24,round(isnull(V.[25],0),0) as _25 from(Select * from(SELECT S.Inicio, E.Semanas, (E.Cantidad-C.Cantidad) Cantidad FROM(Select Semanas, Cantidad, Temporada, Cultivo, Semana " +
                "from Estimacion_Berries where Concepto = 'ALWAYS FRESH')E left join(Select Semanas, Cantidad, Temporada, Cultivo " +
                "from Estimacion_Berries where Concepto = 'COYOTES')C on E.Temporada = C.Temporada and E.Semanas = C.Semanas AND E.Cultivo = C.Cultivo left join CatSemanas S on E.Temporada = S.Temporada AND E.Semana = S.Semana " +
                "where E.Temporada = (select temporada from CatSemanas where GETDATE() between Inicio and Fin) and E.Cultivo = 1)V PIVOT(SUM(Cantidad) For Semanas in ([37],[38],[39],[40],[41],[42],[43],[44],[45],[46],[47],[48],[49],[50],[51],[52],[1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],[17],[18],[19],[20],[21],[22],[23],[24],[25]))X)V ORDER BY V.Inicio").ToList();

            int xAFCZ = 7;
            foreach (var item in always_freshSCZ)
            {
                ws.Cells[string.Format("A{0}", xAFCZ)].Value = "ALWAYS FRESH SIN COYOTES";
                ws.Cells[string.Format("B{0}", xAFCZ)].Value = item._37;
                ws.Cells[string.Format("C{0}", xAFCZ)].Value = item._38;
                ws.Cells[string.Format("D{0}", xAFCZ)].Value = item._39;
                ws.Cells[string.Format("E{0}", xAFCZ)].Value = item._40;
                ws.Cells[string.Format("F{0}", xAFCZ)].Value = item._41;
                ws.Cells[string.Format("G{0}", xAFCZ)].Value = item._42;
                ws.Cells[string.Format("H{0}", xAFCZ)].Value = item._43;
                ws.Cells[string.Format("I{0}", xAFCZ)].Value = item._44;
                ws.Cells[string.Format("J{0}", xAFCZ)].Value = item._45;
                ws.Cells[string.Format("K{0}", xAFCZ)].Value = item._46;
                ws.Cells[string.Format("L{0}", xAFCZ)].Value = item._47;
                ws.Cells[string.Format("M{0}", xAFCZ)].Value = item._48;
                ws.Cells[string.Format("N{0}", xAFCZ)].Value = item._49;
                ws.Cells[string.Format("O{0}", xAFCZ)].Value = item._50;
                ws.Cells[string.Format("P{0}", xAFCZ)].Value = item._51;
                ws.Cells[string.Format("Q{0}", xAFCZ)].Value = item._52;
                ws.Cells[string.Format("R{0}", xAFCZ)].Value = item._1;
                ws.Cells[string.Format("S{0}", xAFCZ)].Value = item._2;
                ws.Cells[string.Format("T{0}", xAFCZ)].Value = item._3;
                ws.Cells[string.Format("U{0}", xAFCZ)].Value = item._4;
                ws.Cells[string.Format("V{0}", xAFCZ)].Value = item._5;
                ws.Cells[string.Format("W{0}", xAFCZ)].Value = item._6;
                ws.Cells[string.Format("X{0}", xAFCZ)].Value = item._7;
                ws.Cells[string.Format("Y{0}", xAFCZ)].Value = item._8;
                ws.Cells[string.Format("Z{0}", xAFCZ)].Value = item._9;
                ws.Cells[string.Format("AA{0}", xAFCZ)].Value = item._10;
                ws.Cells[string.Format("AB{0}", xAFCZ)].Value = item._11;
                ws.Cells[string.Format("AC{0}", xAFCZ)].Value = item._12;
                ws.Cells[string.Format("AD{0}", xAFCZ)].Value = item._13;
                ws.Cells[string.Format("AE{0}", xAFCZ)].Value = item._14;
                ws.Cells[string.Format("AF{0}", xAFCZ)].Value = item._15;
                ws.Cells[string.Format("AG{0}", xAFCZ)].Value = item._16;
                ws.Cells[string.Format("AH{0}", xAFCZ)].Value = item._17;
                ws.Cells[string.Format("AI{0}", xAFCZ)].Value = item._18;
                ws.Cells[string.Format("AJ{0}", xAFCZ)].Value = item._19;
                ws.Cells[string.Format("AK{0}", xAFCZ)].Value = item._20;
                ws.Cells[string.Format("AL{0}", xAFCZ)].Value = item._21;
                ws.Cells[string.Format("AM{0}", xAFCZ)].Value = item._22;
                ws.Cells[string.Format("AN{0}", xAFCZ)].Value = item._23;
                ws.Cells[string.Format("AO{0}", xAFCZ)].Value = item._24;
                ws.Cells[string.Format("AP{0}", xAFCZ)].Value = item._25;
                xAFCZ++;
            }

            //PROYECCION SEMANAL
            ws.Cells["A9"].Value = "PROYECCION SEMANAL";
            ws.Cells["B9"].Value = "37";
            ws.Cells["C9"].Value = "38";
            ws.Cells["D9"].Value = "39";
            ws.Cells["E9"].Value = "40";
            ws.Cells["F9"].Value = "41";
            ws.Cells["G9"].Value = "42";
            ws.Cells["H9"].Value = "43";
            ws.Cells["I9"].Value = "44";
            ws.Cells["J9"].Value = "45";
            ws.Cells["K9"].Value = "46";
            ws.Cells["L9"].Value = "47";
            ws.Cells["M9"].Value = "48";
            ws.Cells["N9"].Value = "49";
            ws.Cells["O9"].Value = "50";
            ws.Cells["P9"].Value = "51";
            ws.Cells["Q9"].Value = "52";
            ws.Cells["R9"].Value = "1";
            ws.Cells["S9"].Value = "2";
            ws.Cells["T9"].Value = "3";
            ws.Cells["U9"].Value = "4";
            ws.Cells["V9"].Value = "5";
            ws.Cells["W9"].Value = "6";
            ws.Cells["X9"].Value = "7";
            ws.Cells["Y9"].Value = "8";
            ws.Cells["Z9"].Value = "9";
            ws.Cells["AA9"].Value = "10";
            ws.Cells["AB9"].Value = "11";
            ws.Cells["AC9"].Value = "12";
            ws.Cells["AD9"].Value = "13";
            ws.Cells["AE9"].Value = "14";
            ws.Cells["AF9"].Value = "15";
            ws.Cells["AG9"].Value = "16";
            ws.Cells["AH9"].Value = "17";
            ws.Cells["AI9"].Value = "18";
            ws.Cells["AJ9"].Value = "19";
            ws.Cells["AK9"].Value = "20";
            ws.Cells["AL9"].Value = "21";
            ws.Cells["AM9"].Value = "22";
            ws.Cells["AN9"].Value = "23";
            ws.Cells["AO9"].Value = "24";
            ws.Cells["AP9"].Value = "25";

            var dataSZ = bd.Database.SqlQuery<E_Berries>("Select V.Sem AS 'SEMANA', round(isnull(V.[37],0),0) as _37,round(isnull(V.[38],0),0) as _38,round(isnull(V.[39],0),0) as _39,round(isnull(V.[40],0),0) as _40,round(isnull(V.[41],0),0) as _41,round(isnull(V.[42],0),0) as _42,round(isnull(V.[43],0),0) as _43,round(isnull(V.[44],0),0) as _44,round(isnull(V.[45],0),0) as _45,round(isnull(V.[46],0),0) as _46,round(isnull(V.[47],0),0) as _47,round(isnull(V.[48],0),0) as _48,round(isnull(V.[49],0),0) as _49,round(isnull(V.[50],0),0) as _50,round(isnull(V.[51],0),0) as _51,round(isnull(V.[52],0),0) as _52,round(isnull(V.[1],0),0) as _1,round(isnull(V.[2],0),0) as _2,round(isnull(V.[3],0),0) as _3,round(isnull(V.[4],0),0) as _4,round(isnull(V.[5],0),0) as _5,round(isnull(V.[6],0),0) as _6,round(isnull(V.[7],0),0) as _7,round(isnull(V.[8],0),0) as _8,round(isnull(V.[9],0),0) as _9,round(isnull(V.[10],0),0) as _10,round(isnull(V.[11],0),0) as _11,round(isnull(V.[12],0),0) as _12,round(isnull(V.[13],0),0) as _13,round(isnull(V.[14],0),0) as _14,round(isnull(V.[15],0),0) as _15,round(isnull(V.[16],0),0) as _16,round(isnull(V.[17],0),0) as _17,round(isnull(V.[18],0),0) as _18,round(isnull(V.[19],0),0) as _19,round(isnull(V.[20],0),0) as _20,round(isnull(V.[21],0),0) as _21,round(isnull(V.[22],0),0) as _22,round(isnull(V.[23],0),0) as _23,round(isnull(V.[24],0),0) as _24,round(isnull(V.[25],0),0) as _25 " +
                "from(Select * from(SELECT S.Semana AS Sem, S.Inicio, E.Semanas, E.Cantidad FROM Estimacion_Berries E " +
                "left join CatSemanas S on E.Temporada = S.Temporada AND E.Semana = S.Semana " +
                "where E.Temporada = (select temporada from CatSemanas where GETDATE() between Inicio and Fin) and E.Concepto = 'PROYECCION SEMANAL' and E.Cultivo = 1)V " +
                "PIVOT(SUM(Cantidad) For Semanas in ([37],[38],[39],[40],[41],[42],[43],[44],[45],[46],[47],[48],[49],[50],[51],[52],[1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],[17],[18],[19],[20],[21],[22],[23],[24],[25]))X)V ORDER BY V.Inicio").ToList();

            int rowStart = 10;
            foreach (var item in dataSZ)
            {
                ws.Cells[string.Format("A{0}", rowStart)].Value = item.SEMANA;
                ws.Cells[string.Format("B{0}", rowStart)].Value = item._37;
                ws.Cells[string.Format("C{0}", rowStart)].Value = item._38;
                ws.Cells[string.Format("D{0}", rowStart)].Value = item._39;
                ws.Cells[string.Format("E{0}", rowStart)].Value = item._40;
                ws.Cells[string.Format("F{0}", rowStart)].Value = item._41;
                ws.Cells[string.Format("G{0}", rowStart)].Value = item._42;
                ws.Cells[string.Format("H{0}", rowStart)].Value = item._43;
                ws.Cells[string.Format("I{0}", rowStart)].Value = item._44;
                ws.Cells[string.Format("J{0}", rowStart)].Value = item._45;
                ws.Cells[string.Format("K{0}", rowStart)].Value = item._46;
                ws.Cells[string.Format("L{0}", rowStart)].Value = item._47;
                ws.Cells[string.Format("M{0}", rowStart)].Value = item._48;
                ws.Cells[string.Format("N{0}", rowStart)].Value = item._49;
                ws.Cells[string.Format("O{0}", rowStart)].Value = item._50;
                ws.Cells[string.Format("P{0}", rowStart)].Value = item._51;
                ws.Cells[string.Format("Q{0}", rowStart)].Value = item._52;
                ws.Cells[string.Format("R{0}", rowStart)].Value = item._1;
                ws.Cells[string.Format("S{0}", rowStart)].Value = item._2;
                ws.Cells[string.Format("T{0}", rowStart)].Value = item._3;
                ws.Cells[string.Format("U{0}", rowStart)].Value = item._4;
                ws.Cells[string.Format("V{0}", rowStart)].Value = item._5;
                ws.Cells[string.Format("W{0}", rowStart)].Value = item._6;
                ws.Cells[string.Format("X{0}", rowStart)].Value = item._7;
                ws.Cells[string.Format("Y{0}", rowStart)].Value = item._8;
                ws.Cells[string.Format("Z{0}", rowStart)].Value = item._9;
                ws.Cells[string.Format("AA{0}", rowStart)].Value = item._10;
                ws.Cells[string.Format("AB{0}", rowStart)].Value = item._11;
                ws.Cells[string.Format("AC{0}", rowStart)].Value = item._12;
                ws.Cells[string.Format("AD{0}", rowStart)].Value = item._13;
                ws.Cells[string.Format("AE{0}", rowStart)].Value = item._14;
                ws.Cells[string.Format("AF{0}", rowStart)].Value = item._15;
                ws.Cells[string.Format("AG{0}", rowStart)].Value = item._16;
                ws.Cells[string.Format("AH{0}", rowStart)].Value = item._17;
                ws.Cells[string.Format("AI{0}", rowStart)].Value = item._18;
                ws.Cells[string.Format("AJ{0}", rowStart)].Value = item._19;
                ws.Cells[string.Format("AK{0}", rowStart)].Value = item._20;
                ws.Cells[string.Format("AL{0}", rowStart)].Value = item._21;
                ws.Cells[string.Format("AM{0}", rowStart)].Value = item._22;
                ws.Cells[string.Format("AN{0}", rowStart)].Value = item._23;
                ws.Cells[string.Format("AO{0}", rowStart)].Value = item._24;
                ws.Cells[string.Format("AP{0}", rowStart)].Value = item._25;
                rowStart++;
            }

            //PORCENTAJES
            ws.Cells["A53"].Value = "PORCENTAJES";
            ws.Cells["B53"].Value = "37";
            ws.Cells["C53"].Value = "38";
            ws.Cells["D53"].Value = "39";
            ws.Cells["E53"].Value = "40";
            ws.Cells["F53"].Value = "41";
            ws.Cells["G53"].Value = "42";
            ws.Cells["H53"].Value = "43";
            ws.Cells["I53"].Value = "44";
            ws.Cells["J53"].Value = "45";
            ws.Cells["K53"].Value = "46";
            ws.Cells["L53"].Value = "47";
            ws.Cells["M53"].Value = "48";
            ws.Cells["N53"].Value = "49";
            ws.Cells["O53"].Value = "50";
            ws.Cells["P53"].Value = "51";
            ws.Cells["Q53"].Value = "52";
            ws.Cells["R53"].Value = "1";
            ws.Cells["S53"].Value = "2";
            ws.Cells["T53"].Value = "3";
            ws.Cells["U53"].Value = "4";
            ws.Cells["V53"].Value = "5";
            ws.Cells["W53"].Value = "6";
            ws.Cells["X53"].Value = "7";
            ws.Cells["Y53"].Value = "8";
            ws.Cells["Z53"].Value = "9";
            ws.Cells["AA53"].Value = "10";
            ws.Cells["AB53"].Value = "11";
            ws.Cells["AC53"].Value = "12";
            ws.Cells["AD53"].Value = "13";
            ws.Cells["AE53"].Value = "14";
            ws.Cells["AF53"].Value = "15";
            ws.Cells["AG53"].Value = "16";
            ws.Cells["AH53"].Value = "17";
            ws.Cells["AI53"].Value = "18";
            ws.Cells["AJ53"].Value = "19";
            ws.Cells["AK53"].Value = "20";
            ws.Cells["AL53"].Value = "21";
            ws.Cells["AM53"].Value = "22";
            ws.Cells["AN53"].Value = "23";
            ws.Cells["AO53"].Value = "24";
            ws.Cells["AP53"].Value = "25";

            var pSZ = bd.Database.SqlQuery<E_BerriesP>("select V.Sem AS 'SEMANA', (case when V._37='0%' then '' else V._37 end) as _37,(case when V._38='0%' then '' else V._38 end) as _38,(case when V._39='0%' then '' else V._39 end) as _39,(case when V._40='0%' then '' else V._40 end) as _40,(case when V._41='0%' then '' else V._41 end) as _41,(case when V._42='0%' then '' else V._42 end) as _42,(case when V._43='0%' then '' else V._43 end) as _43,(case when V._44='0%' then '' else V._44 end) as _44,(case when V._45='0%' then '' else V._45 end) as _45,(case when V._46='0%' then '' else V._46 end) as _46,(case when V._47='0%' then '' else V._47 end) as _47,(case when V._48='0%' then '' else V._48 end) as _48,(case when V._49='0%' then '' else V._49 end) as _49,(case when V._50='0%' then '' else V._50 end) as _50,(case when V._51='0%' then '' else V._51 end) as _51,(case when V._52='0%' then '' else V._52 end) as _52,(case when V._1='0%' then '' else V._1 end) as _1,(case when V._2='0%' then '' else V._2 end) as _2,(case when V._3='0%' then '' else V._3 end) as _3,(case when V._4='0%' then '' else V._4 end) as _4,(case when V._5='0%' then '' else V._5 end) as _5,(case when V._6='0%' then '' else V._6 end) as _6,(case when V._7='0%' then '' else V._7 end) as _7,(case when V._8='0%' then '' else V._8 end) as _8,(case when V._9='0%' then '' else V._9 end) as _9,(case when V._10='0%' then '' else V._10 end) as _10,(case when V._11='0%' then '' else V._11 end) as _11,(case when V._12='0%' then '' else V._12 end) as _12,(case when V._13='0%' then '' else V._13 end) as _13,(case when V._14='0%' then '' else V._14 end) as _14,(case when V._15='0%' then '' else V._15 end) as _15,(case when V._16='0%' then '' else V._16 end) as _16,(case when V._17='0%' then '' else V._17 end) as _17,(case when V._18='0%' then '' else V._18 end) as _18,(case when V._19='0%' then '' else V._19 end) as _19,(case when V._20='0%' then '' else V._20 end) as _20,(case when V._21='0%' then '' else V._21 end) as _21,(case when V._22='0%' then '' else V._22 end) as _22,(case when V._23='0%' then '' else V._23 end) as _23,(case when V._24='0%' then '' else V._24 end) as _24,(case when V._25='0%' then '' else V._25 end) as _25 from (select V.Inicio, V.Sem, cast(V._37 as varchar)+'%' as _37,cast(V._38 as varchar)+'%' as _38,cast(V._39 as varchar)+'%' as _39,cast(V._40 as varchar)+'%' as _40,cast(V._41 as varchar)+'%' as _41,cast(V._42 as varchar)+'%' as _42,cast(V._43 as varchar)+'%' as _43,cast(V._44 as varchar)+'%' as _44,cast(V._45 as varchar)+'%' as _45,cast(V._46 as varchar)+'%' as _46,cast(V._47 as varchar)+'%' as _47,cast(V._48 as varchar)+'%' as _48,cast(V._49 as varchar)+'%' as _49,cast(V._50 as varchar)+'%' as _50,cast(V._51 as varchar)+'%' as _51,cast(V._52 as varchar)+'%' as _52,cast(V._1 as varchar)+'%' as _1,cast(V._2 as varchar)+'%' _2,cast(V._3 as varchar)+'%' as _3,cast(V._4 as varchar)+'%' as _4,cast(V._5 as varchar)+'%' as _5,cast(V._6 as varchar)+'%' as _6,cast(V._7 as varchar)+'%' as _7,cast(V._8 as varchar)+'%' as _8,cast(V._9 as varchar)+'%' as _9,cast(V._10 as varchar)+'%' as _10,cast(V._11 as varchar)+'%' as _11,cast(V._12 as varchar)+'%' as _12,cast(V._13 as varchar)+'%' as _13,cast(V._14 as varchar)+'%' as _14,cast(V._15 as varchar)+'%' as _15,cast(V._16 as varchar)+'%' as _16,cast(V._17 as varchar)+'%' as _17,cast(V._18 as varchar)+'%' as _18,cast(V._19 as varchar)+'%' as _19,cast(V._20 as varchar)+'%'as _20,cast(V._21 as varchar)+'%' as _21,cast(V._22 as varchar)+'%' as _22,cast(V._23 as varchar)+'%' as _23,cast(V._24 as varchar)+'%' as _24,cast(V._25 as varchar)+'%' as _25 from (SELECT V.Inicio, V.Sem,(case when A._37=0 then 0 else case when V._37=0 then 0 else round((A._37/V._37)*100,0) end end) as _37,(case when A._38=0 then 0 else case when V._38=0 then 0 else round((A._38/V._38)*100,0) end end) as _38,(case when A._39=0 then 0 else case when V._39=0 then 0 else round((A._39/V._39)*100,0) end end) as _39,(case when A._40=0 then 0 else case when V._40=0 then 0 else round((A._40/V._40)*100,0) end end) as _40,(case when A._41=0 then 0 else case when V._41=0 then 0 else round((A._41/V._41)*100,0) end end) as _41,(case when A._42=0 then 0 else case when V._42=0 then 0 else round((A._42/V._42)*100,0) end end) as _42,(case when A._43=0 then 0 else case when V._43=0 then 0 else round((A._43/V._43)*100,0) end end) as _43,(case when A._44=0 then 0 else case when V._44=0 then 0 else round((A._44/V._44)*100,0) end end) as _44,(case when A._45=0 then 0 else case when V._45=0 then 0 else round((A._45/V._45)*100,0) end end) as _45,(case when A._46=0 then 0 else case when V._46=0 then 0 else round((A._46/V._46)*100,0) end end) as _46,(case when A._47=0 then 0 else case when V._47=0 then 0 else round((A._47/V._47)*100,0) end end) as _47,(case when A._48=0 then 0 else case when V._48=0 then 0 else round((A._48/V._48)*100,0) end end) as _48,(case when A._49=0 then 0 else case when V._49=0 then 0 else round((A._49/V._49)*100,0) end end) as _49,(case when A._50=0 then 0 else case when V._50=0 then 0 else round((A._50/V._50)*100,0) end end) as _50,(case when A._51=0 then 0 else case when V._51=0 then 0 else round((A._51/V._51)*100,0) end end) as _51,(case when A._52=0 then 0 else case when V._52=0 then 0 else round((A._52/V._52)*100,0) end end) as _52,(case when A._1=0 then 0 else case when V._1=0 then 0 else round((A._1/V._1)*100,0) end end) as _1,(case when A._2=0 then 0 else case when V._2=0 then 0 else round((A._2/V._2)*100,0) end end) as _2,(case when A._3=0 then 0 else case when V._3=0 then 0 else round((A._3/V._3)*100,0) end end) as _3,(case when A._4=0 then 0 else case when V._4=0 then 0 else round((A._4/V._4)*100,0) end end) as _4,(case when A._5=0 then 0 else case when V._5=0 then 0 else round((A._5/V._5)*100,0) end end) as _5,(case when A._6=0 then 0 else case when V._6=0 then 0 else round((A._6/V._6)*100,0) end end) as _6,(case when A._7=0 then 0 else case when V._7=0 then 0 else round((A._7/V._7)*100,0) end end) as _7,(case when A._8=0 then 0 else case when V._8=0 then 0 else round((A._8/V._8)*100,0) end end) as _8,(case when A._9=0 then 0 else case when V._9=0 then 0 else round((A._9/V._9)*100,0) end end) as _9,(case when A._10=0 then 0 else case when V._10=0 then 0 else round((A._10/V._10)*100,0) end end) as _10,(case when A._11=0 then 0 else case when V._11=0 then 0 else round((A._11/V._11)*100,0) end end) as _11,(case when A._12=0 then 0 else case when V._12=0 then 0 else round((A._12/V._12)*100,0) end end) as _12,(case when A._13=0 then 0 else case when V._13=0 then 0 else round((A._13/V._13)*100,0) end end) as _13,(case when A._14=0 then 0 else case when V._14=0 then 0 else round((A._14/V._14)*100,0) end end) as _14,(case when A._15=0 then 0 else case when V._15=0 then 0 else round((A._15/V._15)*100,0) end end) as _15,(case when A._16=0 then 0 else case when V._16=0 then 0 else round((A._16/V._16)*100,0) end end) as _16,(case when A._17=0 then 0 else case when V._17=0 then 0 else round((A._17/V._17)*100,0) end end) as _17,(case when A._18=0 then 0 else case when V._18=0 then 0 else round((A._18/V._18)*100,0) end end) as _18,(case when A._19=0 then 0 else case when V._19=0 then 0 else round((A._19/V._19)*100,0) end end) as _19,(case when A._20=0 then 0 else case when V._20=0 then 0 else round((A._20/V._20)*100,0) end end) as _20,(case when A._21=0 then 0 else case when V._21=0 then 0 else round((A._21/V._21)*100,0) end end) as _21,(case when A._22=0 then 0 else case when V._22=0 then 0 else round((A._22/V._22)*100,0) end end) as _22,(case when A._23=0 then 0 else case when V._23=0 then 0 else round((A._23/V._23)*100,0) end end) as _23,(case when A._24=0 then 0 else case when V._24=0 then 0 else round((A._24/V._24)*100,0) end end) as _24,(case when A._25=0 then 0 else case when V._25=0 then 0 else round((A._25/V._25)*100,0) end end) as _25 FROM(Select V.Cultivo, isnull(V.[37],0) as _37,isnull(V.[38],0) as _38,isnull(V.[39],0) as _39,isnull(V.[40],0) as _40,isnull(V.[41],0) as _41,isnull(V.[42],0) as _42,isnull(V.[43],0) as _43,isnull(V.[44],0) as _44,isnull(V.[45],0)as _45,isnull(V.[46],0) as _46,isnull(V.[47],0) as _47,isnull(V.[48],0) as _48,isnull(V.[49],0) as _49,isnull(V.[50],0) as _50,isnull(V.[51],0) as _51,isnull(V.[52],0) as _52,isnull(V.[1],0) as _1,isnull(V.[2],0) as _2,isnull(V.[3],0) as _3,isnull(V.[4],0) as _4,isnull(V.[5],0) as _5,isnull(V.[6],0) as _6,isnull(V.[7],0) as _7,isnull(V.[8],0) as _8,isnull(V.[9],0) as _9,isnull(V.[10],0) as _10,isnull(V.[11],0) as _11,isnull(V.[12],0) as _12,isnull(V.[13],0) as _13,isnull(V.[14],0) as _14,isnull(V.[15],0) as _15,isnull(V.[16],0) as _16,isnull(V.[17],0) as _17,isnull(V.[18],0) as _18,isnull(V.[19],0) as _19,isnull(V.[20],0) as _20,isnull(V.[21],0) as _21,isnull(V.[22],0) as _22,isnull(V.[23],0) as _23,isnull(V.[24],0) as _24,isnull(V.[25],0) as _25 from (Select * from(SELECT E.Concepto, S.Inicio, E.Semanas, E.Cantidad, E.Cultivo FROM Estimacion_Berries E left join CatSemanas S on E.Temporada=S.Temporada AND E.Semana=S.Semana where E.Temporada=(select temporada from CatSemanas where GETDATE() between Inicio and Fin) and " +
                "E.Concepto = 'ALWAYS FRESH' and E.Cultivo = 1)V PIVOT(SUM(Cantidad) For Semanas in ([37],[38],[39],[40],[41],[42],[43],[44],[45],[46],[47],[48],[49],[50],[51],[52],[1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],[17],[18],[19],[20],[21],[22],[23],[24],[25]))X)V)A left join(Select V.Inicio, V.Cultivo, V.Sem, isnull(V.[37], 0) as _37, isnull(V.[38], 0) as _38, isnull(V.[39], 0) as _39, isnull(V.[40], 0) as _40, isnull(V.[41], 0) as _41, isnull(V.[42], 0) as _42, isnull(V.[43], 0) as _43, isnull(V.[44], 0) as _44, isnull(V.[45], 0) as _45, isnull(V.[46], 0) as _46, isnull(V.[47], 0) as _47, isnull(V.[48], 0) as _48, isnull(V.[49], 0) as _49, isnull(V.[50], 0) as _50, isnull(V.[51], 0) as _51, isnull(V.[52], 0) as _52, isnull(V.[1], 0) as _1, isnull(V.[2], 0) as _2, isnull(V.[3], 0) as _3, isnull(V.[4], 0) as _4, isnull(V.[5], 0) as _5, isnull(V.[6], 0) as _6, isnull(V.[7], 0) as _7, isnull(V.[8], 0) as _8, isnull(V.[9], 0) as _9, isnull(V.[10], 0) as _10, isnull(V.[11], 0) as _11, isnull(V.[12], 0) as _12, isnull(V.[13], 0) as _13, isnull(V.[14], 0) as _14, isnull(V.[15], 0) as _15, isnull(V.[16], 0) as _16, isnull(V.[17], 0) as _17, isnull(V.[18], 0) as _18, isnull(V.[19], 0) as _19, isnull(V.[20], 0) as _20, isnull(V.[21], 0) as _21, isnull(V.[22], 0) as _22, isnull(V.[23], 0) as _23, isnull(V.[24], 0) as _24, isnull(V.[25], 0) as _25 " +
                "from(select * from(SELECT S.Semana AS Sem, S.Inicio, E.Semanas, E.Cantidad, E.Cultivo FROM Estimacion_Berries E left join CatSemanas S on E.Temporada = S.Temporada AND E.Semana = S.Semana where E.Temporada = (select temporada from CatSemanas where GETDATE() between Inicio and Fin) and " +
                "E.Concepto = 'PROYECCION SEMANAL' and E.Cultivo = 1)V PIVOT(SUM(Cantidad) For Semanas in ([37],[38],[39],[40],[41],[42],[43],[44],[45],[46],[47],[48],[49],[50],[51],[52],[1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],[17],[18],[19],[20],[21],[22],[23],[24],[25]))X)V )V on A.Cultivo = V.Cultivo)V)V ORDER BY V.Inicio").ToList();

            int xPSZ = 54;
            foreach (var item in pSZ)
            {
                ws.Cells[string.Format("A{0}", xPSZ)].Value = item.SEMANA;
                ws.Cells[string.Format("B{0}", xPSZ)].Value = item._37;
                ws.Cells[string.Format("C{0}", xPSZ)].Value = item._38;
                ws.Cells[string.Format("D{0}", xPSZ)].Value = item._39;
                ws.Cells[string.Format("E{0}", xPSZ)].Value = item._40;
                ws.Cells[string.Format("F{0}", xPSZ)].Value = item._41;
                ws.Cells[string.Format("G{0}", xPSZ)].Value = item._42;
                ws.Cells[string.Format("H{0}", xPSZ)].Value = item._43;
                ws.Cells[string.Format("I{0}", xPSZ)].Value = item._44;
                ws.Cells[string.Format("J{0}", xPSZ)].Value = item._45;
                ws.Cells[string.Format("K{0}", xPSZ)].Value = item._46;
                ws.Cells[string.Format("L{0}", xPSZ)].Value = item._47;
                ws.Cells[string.Format("M{0}", xPSZ)].Value = item._48;
                ws.Cells[string.Format("N{0}", xPSZ)].Value = item._49;
                ws.Cells[string.Format("O{0}", xPSZ)].Value = item._50;
                ws.Cells[string.Format("P{0}", xPSZ)].Value = item._51;
                ws.Cells[string.Format("Q{0}", xPSZ)].Value = item._52;
                ws.Cells[string.Format("R{0}", xPSZ)].Value = item._1;
                ws.Cells[string.Format("S{0}", xPSZ)].Value = item._2;
                ws.Cells[string.Format("T{0}", xPSZ)].Value = item._3;
                ws.Cells[string.Format("U{0}", xPSZ)].Value = item._4;
                ws.Cells[string.Format("V{0}", xPSZ)].Value = item._5;
                ws.Cells[string.Format("W{0}", xPSZ)].Value = item._6;
                ws.Cells[string.Format("X{0}", xPSZ)].Value = item._7;
                ws.Cells[string.Format("Y{0}", xPSZ)].Value = item._8;
                ws.Cells[string.Format("Z{0}", xPSZ)].Value = item._9;
                ws.Cells[string.Format("AA{0}", xPSZ)].Value = item._10;
                ws.Cells[string.Format("AB{0}", xPSZ)].Value = item._11;
                ws.Cells[string.Format("AC{0}", xPSZ)].Value = item._12;
                ws.Cells[string.Format("AD{0}", xPSZ)].Value = item._13;
                ws.Cells[string.Format("AE{0}", xPSZ)].Value = item._14;
                ws.Cells[string.Format("AF{0}", xPSZ)].Value = item._15;
                ws.Cells[string.Format("AG{0}", xPSZ)].Value = item._16;
                ws.Cells[string.Format("AH{0}", xPSZ)].Value = item._17;
                ws.Cells[string.Format("AI{0}", xPSZ)].Value = item._18;
                ws.Cells[string.Format("AJ{0}", xPSZ)].Value = item._19;
                ws.Cells[string.Format("AK{0}", xPSZ)].Value = item._20;
                ws.Cells[string.Format("AL{0}", xPSZ)].Value = item._21;
                ws.Cells[string.Format("AM{0}", xPSZ)].Value = item._22;
                ws.Cells[string.Format("AN{0}", xPSZ)].Value = item._23;
                ws.Cells[string.Format("AO{0}", xPSZ)].Value = item._24;
                ws.Cells[string.Format("AP{0}", xPSZ)].Value = item._25;
                xPSZ++;
            }

            ws.Cells["A:AP"].AutoFitColumns();

            ws.Row(2).Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            ws.Row(2).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(string.Format("yellow")));
            ws.Row(9).Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            ws.Row(9).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(string.Format("yellow")));
            ws.Row(53).Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            ws.Row(53).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(string.Format("yellow")));

            //-----------------------------------FRAMBUESA----------------------------------------------
            ExcelWorksheet wsF = excel.Workbook.Worksheets.Add("FRAMBUESA");
            wsF.Cells["A1"].Value = "VARIACION SEMANAL DE LA ESTIMACION DE FRAMBUESA";
            wsF.Cells["A2"].Value = "SEMANA";
            wsF.Cells["B2"].Value = "37";
            wsF.Cells["C2"].Value = "38";
            wsF.Cells["D2"].Value = "39";
            wsF.Cells["E2"].Value = "40";
            wsF.Cells["F2"].Value = "41";
            wsF.Cells["G2"].Value = "42";
            wsF.Cells["H2"].Value = "43";
            wsF.Cells["I2"].Value = "44";
            wsF.Cells["J2"].Value = "45";
            wsF.Cells["K2"].Value = "46";
            wsF.Cells["L2"].Value = "47";
            wsF.Cells["M2"].Value = "48";
            wsF.Cells["N2"].Value = "49";
            wsF.Cells["O2"].Value = "50";
            wsF.Cells["P2"].Value = "51";
            wsF.Cells["Q2"].Value = "52";
            wsF.Cells["R2"].Value = "1";
            wsF.Cells["S2"].Value = "2";
            wsF.Cells["T2"].Value = "3";
            wsF.Cells["U2"].Value = "4";
            wsF.Cells["V2"].Value = "5";
            wsF.Cells["W2"].Value = "6";
            wsF.Cells["X2"].Value = "7";
            wsF.Cells["Y2"].Value = "8";
            wsF.Cells["Z2"].Value = "9";
            wsF.Cells["AA2"].Value = "10";
            wsF.Cells["AB2"].Value = "11";
            wsF.Cells["AC2"].Value = "12";
            wsF.Cells["AD2"].Value = "13";
            wsF.Cells["AE2"].Value = "14";
            wsF.Cells["AF2"].Value = "15";
            wsF.Cells["AG2"].Value = "16";
            wsF.Cells["AH2"].Value = "17";
            wsF.Cells["AI2"].Value = "18";
            wsF.Cells["AJ2"].Value = "19";
            wsF.Cells["AK2"].Value = "20";
            wsF.Cells["AL2"].Value = "21";
            wsF.Cells["AM2"].Value = "22";
            wsF.Cells["AN2"].Value = "23";
            wsF.Cells["AO2"].Value = "24";
            wsF.Cells["AP2"].Value = "25";

            var recepcion_realF = bd.Database.SqlQuery<R_Berries>("SELECT 'TOTAL' as 'SEMANA',  round(isnull(V.[37],0),0) as _37,round(isnull(V.[38],0),0) as _38,round(isnull(V.[39],0),0) as _39,round(isnull(V.[40],0),0) as _40,round(isnull(V.[41],0),0) as _41,round(isnull(V.[42],0),0) as _42,round(isnull(V.[43],0),0) as _43,round(isnull(V.[44],0),0) as _44,round(isnull(V.[45],0),0) as _45,round(isnull(V.[46],0),0) as _46,round(isnull(V.[47],0),0) as _47,round(isnull(V.[48],0),0) as _48,round(isnull(V.[49],0),0) as _49,round(isnull(V.[50],0),0) as _50,round(isnull(V.[51],0),0) as _51,round(isnull(V.[52],0),0) as _52,round(isnull(V.[1],0),0) as _1,round(isnull(V.[2],0),0) as _2,round(isnull(V.[3],0),0) as _3,round(isnull(V.[4],0),0) as _4,round(isnull(V.[5],0),0) as _5,round(isnull(V.[6],0),0) as _6,round(isnull(V.[7],0),0) as _7,round(isnull(V.[8],0),0) as _8,round(isnull(V.[9],0),0) as _9,round(isnull(V.[10],0),0) as _10,round(isnull(V.[11],0),0) as _11,round(isnull(V.[12],0),0) as _12,round(isnull(V.[13],0),0) as _13,round(isnull(V.[14],0),0) as _14,round(isnull(V.[15],0),0) as _15,round(isnull(V.[16],0),0) as _16,round(isnull(V.[17],0),0) as _17,round(isnull(V.[18],0),0) as _18,round(isnull(V.[19],0),0) as _19,round(isnull(V.[20],0),0) as _20,round(isnull(V.[21],0),0) as _21,round(isnull(V.[22],0),0) as _22,round(isnull(V.[23],0),0) as _23,round(isnull(V.[24],0),0) as _24,round(isnull(V.[25],0),0) as _25 " +
               "FROM(Select * from(SELECT Semana, SUM(Convertidas) AS Convertidas FROM SEasonsun1213..UV_ProdRecepcion where CodEstatus <> 'C' and temporada = (select temporada from CatSemanas where GETDATE() between Inicio and Fin) and DescProducto = 'FRAMBUESA' GROUP BY Semana Union All " +
               "SELECT Semana, SUM(Convertidas) AS Convertidas FROM SeasonPlan..UV_ProdRecepcion where CodEstatus <> 'C' and temporada = (select temporada from CatSemanas where GETDATE() between Inicio and Fin) and DescProducto = 'FRAMBUESA' GROUP BY Semana" +
               ")V PIVOT(SUM(Convertidas) FOR Semana in ([37],[38],[39],[40],[41],[42],[43],[44],[45],[46],[47],[48],[49],[50],[51],[52],[1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],[17],[18],[19],[20],[21],[22],[23],[24],[25]))X)V ").ToList();

            int xRF = 3;
            foreach (var item in recepcion_realF)
            {
                wsF.Cells[string.Format("A{0}", xRF)].Value = "RECEPCION REAL";
                wsF.Cells[string.Format("B{0}", xRF)].Value = item._37;
                wsF.Cells[string.Format("C{0}", xRF)].Value = item._38;
                wsF.Cells[string.Format("D{0}", xRF)].Value = item._39;
                wsF.Cells[string.Format("E{0}", xRF)].Value = item._40;
                wsF.Cells[string.Format("F{0}", xRF)].Value = item._41;
                wsF.Cells[string.Format("G{0}", xRF)].Value = item._42;
                wsF.Cells[string.Format("H{0}", xRF)].Value = item._43;
                wsF.Cells[string.Format("I{0}", xRF)].Value = item._44;
                wsF.Cells[string.Format("J{0}", xRF)].Value = item._45;
                wsF.Cells[string.Format("K{0}", xRF)].Value = item._46;
                wsF.Cells[string.Format("L{0}", xRF)].Value = item._47;
                wsF.Cells[string.Format("M{0}", xRF)].Value = item._48;
                wsF.Cells[string.Format("N{0}", xRF)].Value = item._49;
                wsF.Cells[string.Format("O{0}", xRF)].Value = item._50;
                wsF.Cells[string.Format("P{0}", xRF)].Value = item._51;
                wsF.Cells[string.Format("Q{0}", xRF)].Value = item._52;
                wsF.Cells[string.Format("R{0}", xRF)].Value = item._1;
                wsF.Cells[string.Format("S{0}", xRF)].Value = item._2;
                wsF.Cells[string.Format("T{0}", xRF)].Value = item._3;
                wsF.Cells[string.Format("U{0}", xRF)].Value = item._4;
                wsF.Cells[string.Format("V{0}", xRF)].Value = item._5;
                wsF.Cells[string.Format("W{0}", xRF)].Value = item._6;
                wsF.Cells[string.Format("X{0}", xRF)].Value = item._7;
                wsF.Cells[string.Format("Y{0}", xRF)].Value = item._8;
                wsF.Cells[string.Format("Z{0}", xRF)].Value = item._9;
                wsF.Cells[string.Format("AA{0}", xRF)].Value = item._10;
                wsF.Cells[string.Format("AB{0}", xRF)].Value = item._11;
                wsF.Cells[string.Format("AC{0}", xRF)].Value = item._12;
                wsF.Cells[string.Format("AD{0}", xRF)].Value = item._13;
                wsF.Cells[string.Format("AE{0}", xRF)].Value = item._14;
                wsF.Cells[string.Format("AF{0}", xRF)].Value = item._15;
                wsF.Cells[string.Format("AG{0}", xRF)].Value = item._16;
                wsF.Cells[string.Format("AH{0}", xRF)].Value = item._17;
                wsF.Cells[string.Format("AI{0}", xRF)].Value = item._18;
                wsF.Cells[string.Format("AJ{0}", xRF)].Value = item._19;
                wsF.Cells[string.Format("AK{0}", xRF)].Value = item._20;
                wsF.Cells[string.Format("AL{0}", xRF)].Value = item._21;
                wsF.Cells[string.Format("AM{0}", xRF)].Value = item._22;
                wsF.Cells[string.Format("AN{0}", xRF)].Value = item._23;
                wsF.Cells[string.Format("AO{0}", xRF)].Value = item._24;
                wsF.Cells[string.Format("AP{0}", xRF)].Value = item._25;
                xRF++;
            }

            //PROYECCION SEMANAL
            wsF.Cells["A5"].Value = "PROYECCION SEMANAL";
            wsF.Cells["B5"].Value = "37";
            wsF.Cells["C5"].Value = "38";
            wsF.Cells["D5"].Value = "39";
            wsF.Cells["E5"].Value = "40";
            wsF.Cells["F5"].Value = "41";
            wsF.Cells["G5"].Value = "42";
            wsF.Cells["H5"].Value = "43";
            wsF.Cells["I5"].Value = "44";
            wsF.Cells["J5"].Value = "45";
            wsF.Cells["K5"].Value = "46";
            wsF.Cells["L5"].Value = "47";
            wsF.Cells["M5"].Value = "48";
            wsF.Cells["N5"].Value = "49";
            wsF.Cells["O5"].Value = "50";
            wsF.Cells["P5"].Value = "51";
            wsF.Cells["Q5"].Value = "52";
            wsF.Cells["R5"].Value = "1";
            wsF.Cells["S5"].Value = "2";
            wsF.Cells["T5"].Value = "3";
            wsF.Cells["U5"].Value = "4";
            wsF.Cells["V5"].Value = "5";
            wsF.Cells["W5"].Value = "6";
            wsF.Cells["X5"].Value = "7";
            wsF.Cells["Y5"].Value = "8";
            wsF.Cells["Z5"].Value = "9";
            wsF.Cells["AA5"].Value = "10";
            wsF.Cells["AB5"].Value = "11";
            wsF.Cells["AC5"].Value = "12";
            wsF.Cells["AD5"].Value = "13";
            wsF.Cells["AE5"].Value = "14";
            wsF.Cells["AF5"].Value = "15";
            wsF.Cells["AG5"].Value = "16";
            wsF.Cells["AH5"].Value = "17";
            wsF.Cells["AI5"].Value = "18";
            wsF.Cells["AJ5"].Value = "19";
            wsF.Cells["AK5"].Value = "20";
            wsF.Cells["AL5"].Value = "21";
            wsF.Cells["AM5"].Value = "22";
            wsF.Cells["AN5"].Value = "23";
            wsF.Cells["AO5"].Value = "24";
            wsF.Cells["AP5"].Value = "25";

            var dataSF = bd.Database.SqlQuery<E_Berries>("Select V.Sem AS 'SEMANA', round(isnull(V.[37],0),0) as _37,round(isnull(V.[38],0),0) as _38,round(isnull(V.[39],0),0) as _39,round(isnull(V.[40],0),0) as _40,round(isnull(V.[41],0),0) as _41,round(isnull(V.[42],0),0) as _42,round(isnull(V.[43],0),0) as _43,round(isnull(V.[44],0),0) as _44,round(isnull(V.[45],0),0) as _45,round(isnull(V.[46],0),0) as _46,round(isnull(V.[47],0),0) as _47,round(isnull(V.[48],0),0) as _48,round(isnull(V.[49],0),0) as _49,round(isnull(V.[50],0),0) as _50,round(isnull(V.[51],0),0) as _51,round(isnull(V.[52],0),0) as _52,round(isnull(V.[1],0),0) as _1,round(isnull(V.[2],0),0) as _2,round(isnull(V.[3],0),0) as _3,round(isnull(V.[4],0),0) as _4,round(isnull(V.[5],0),0) as _5,round(isnull(V.[6],0),0) as _6,round(isnull(V.[7],0),0) as _7,round(isnull(V.[8],0),0) as _8,round(isnull(V.[9],0),0) as _9,round(isnull(V.[10],0),0) as _10,round(isnull(V.[11],0),0) as _11,round(isnull(V.[12],0),0) as _12,round(isnull(V.[13],0),0) as _13,round(isnull(V.[14],0),0) as _14,round(isnull(V.[15],0),0) as _15,round(isnull(V.[16],0),0) as _16,round(isnull(V.[17],0),0) as _17,round(isnull(V.[18],0),0) as _18,round(isnull(V.[19],0),0) as _19,round(isnull(V.[20],0),0) as _20,round(isnull(V.[21],0),0) as _21,round(isnull(V.[22],0),0) as _22,round(isnull(V.[23],0),0) as _23,round(isnull(V.[24],0),0) as _24,round(isnull(V.[25],0),0) as _25 " +
                "from(Select * from(SELECT S.Semana AS Sem, S.Inicio, E.Semanas, E.Cantidad FROM Estimacion_Berries E left join CatSemanas S on E.Temporada = S.Temporada AND E.Semana = S.Semana " +
                "where E.Temporada = (select temporada from CatSemanas where GETDATE() between Inicio and Fin) and E.Concepto = 'PROYECCION SEMANAL' and E.Cultivo = 2)V " +
                "PIVOT(SUM(Cantidad) For Semanas in ([37],[38],[39],[40],[41],[42],[43],[44],[45],[46],[47],[48],[49],[50],[51],[52],[1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],[17],[18],[19],[20],[21],[22],[23],[24],[25]))X)V ORDER BY V.Inicio").ToList();

            int xSF = 6;
            foreach (var item in dataSF)
            {
                wsF.Cells[string.Format("A{0}", xSF)].Value = item.SEMANA;
                wsF.Cells[string.Format("B{0}", xSF)].Value = item._37;
                wsF.Cells[string.Format("C{0}", xSF)].Value = item._38;
                wsF.Cells[string.Format("D{0}", xSF)].Value = item._39;
                wsF.Cells[string.Format("E{0}", xSF)].Value = item._40;
                wsF.Cells[string.Format("F{0}", xSF)].Value = item._41;
                wsF.Cells[string.Format("G{0}", xSF)].Value = item._42;
                wsF.Cells[string.Format("H{0}", xSF)].Value = item._43;
                wsF.Cells[string.Format("I{0}", xSF)].Value = item._44;
                wsF.Cells[string.Format("J{0}", xSF)].Value = item._45;
                wsF.Cells[string.Format("K{0}", xSF)].Value = item._46;
                wsF.Cells[string.Format("L{0}", xSF)].Value = item._47;
                wsF.Cells[string.Format("M{0}", xSF)].Value = item._48;
                wsF.Cells[string.Format("N{0}", xSF)].Value = item._49;
                wsF.Cells[string.Format("O{0}", xSF)].Value = item._50;
                wsF.Cells[string.Format("P{0}", xSF)].Value = item._51;
                wsF.Cells[string.Format("Q{0}", xSF)].Value = item._52;
                wsF.Cells[string.Format("R{0}", xSF)].Value = item._1;
                wsF.Cells[string.Format("S{0}", xSF)].Value = item._2;
                wsF.Cells[string.Format("T{0}", xSF)].Value = item._3;
                wsF.Cells[string.Format("U{0}", xSF)].Value = item._4;
                wsF.Cells[string.Format("V{0}", xSF)].Value = item._5;
                wsF.Cells[string.Format("W{0}", xSF)].Value = item._6;
                wsF.Cells[string.Format("X{0}", xSF)].Value = item._7;
                wsF.Cells[string.Format("Y{0}", xSF)].Value = item._8;
                wsF.Cells[string.Format("Z{0}", xSF)].Value = item._9;
                wsF.Cells[string.Format("AA{0}", xSF)].Value = item._10;
                wsF.Cells[string.Format("AB{0}", xSF)].Value = item._11;
                wsF.Cells[string.Format("AC{0}", xSF)].Value = item._12;
                wsF.Cells[string.Format("AD{0}", xSF)].Value = item._13;
                wsF.Cells[string.Format("AE{0}", xSF)].Value = item._14;
                wsF.Cells[string.Format("AF{0}", xSF)].Value = item._15;
                wsF.Cells[string.Format("AG{0}", xSF)].Value = item._16;
                wsF.Cells[string.Format("AH{0}", xSF)].Value = item._17;
                wsF.Cells[string.Format("AI{0}", xSF)].Value = item._18;
                wsF.Cells[string.Format("AJ{0}", xSF)].Value = item._19;
                wsF.Cells[string.Format("AK{0}", xSF)].Value = item._20;
                wsF.Cells[string.Format("AL{0}", xSF)].Value = item._21;
                wsF.Cells[string.Format("AM{0}", xSF)].Value = item._22;
                wsF.Cells[string.Format("AN{0}", xSF)].Value = item._23;
                wsF.Cells[string.Format("AO{0}", xSF)].Value = item._24;
                wsF.Cells[string.Format("AP{0}", xSF)].Value = item._25;
                xSF++;
            }

            //PORCENTAJES
            wsF.Cells["A49"].Value = "PORCENTAJES";
            wsF.Cells["B49"].Value = "37";
            wsF.Cells["C49"].Value = "38";
            wsF.Cells["D49"].Value = "39";
            wsF.Cells["E49"].Value = "40";
            wsF.Cells["F49"].Value = "41";
            wsF.Cells["G49"].Value = "42";
            wsF.Cells["H49"].Value = "43";
            wsF.Cells["I49"].Value = "44";
            wsF.Cells["J49"].Value = "45";
            wsF.Cells["K49"].Value = "46";
            wsF.Cells["L49"].Value = "47";
            wsF.Cells["M49"].Value = "48";
            wsF.Cells["N49"].Value = "49";
            wsF.Cells["O49"].Value = "50";
            wsF.Cells["P49"].Value = "51";
            wsF.Cells["Q49"].Value = "52";
            wsF.Cells["R49"].Value = "1";
            wsF.Cells["S49"].Value = "2";
            wsF.Cells["T49"].Value = "3";
            wsF.Cells["U49"].Value = "4";
            wsF.Cells["V49"].Value = "5";
            wsF.Cells["W49"].Value = "6";
            wsF.Cells["X49"].Value = "7";
            wsF.Cells["Y49"].Value = "8";
            wsF.Cells["Z49"].Value = "9";
            wsF.Cells["AA48"].Value = "10";
            wsF.Cells["AB49"].Value = "11";
            wsF.Cells["AC49"].Value = "12";
            wsF.Cells["AD49"].Value = "13";
            wsF.Cells["AE49"].Value = "14";
            wsF.Cells["AF49"].Value = "15";
            wsF.Cells["AG49"].Value = "16";
            wsF.Cells["AH49"].Value = "17";
            wsF.Cells["AI49"].Value = "18";
            wsF.Cells["AJ49"].Value = "19";
            wsF.Cells["AK49"].Value = "20";
            wsF.Cells["AL49"].Value = "21";
            wsF.Cells["AM49"].Value = "22";
            wsF.Cells["AN49"].Value = "23";
            wsF.Cells["AO49"].Value = "24";
            wsF.Cells["AP49"].Value = "25";

            var pSF = bd.Database.SqlQuery<E_BerriesP>("select V.Sem AS 'SEMANA',(case when V._37='0%' then '' else V._37 end) as _37,(case when V._38='0%' then '' else V._38 end) as _38,(case when V._39='0%' then '' else V._39 end) as _39,(case when V._40='0%' then '' else V._40 end) as _40,(case when V._41='0%' then '' else V._41 end) as _41,(case when V._42='0%' then '' else V._42 end) as _42,(case when V._43='0%' then '' else V._43 end) as _43,(case when V._44='0%' then '' else V._44 end) as _44,(case when V._45='0%' then '' else V._45 end) as _45,(case when V._46='0%' then '' else V._46 end) as _46,(case when V._47='0%' then '' else V._47 end) as _47,(case when V._48='0%' then '' else V._48 end) as _48,(case when V._49='0%' then '' else V._49 end) as _49,(case when V._50='0%' then '' else V._50 end) as _50,(case when V._51='0%' then '' else V._51 end) as _51,(case when V._52='0%' then '' else V._52 end) as _52,(case when V._1='0%' then '' else V._1 end) as _1,(case when V._2='0%' then '' else V._2 end) as _2,(case when V._3='0%' then '' else V._3 end) as _3,(case when V._4='0%' then '' else V._4 end) as _4,(case when V._5='0%' then '' else V._5 end) as _5,(case when V._6='0%' then '' else V._6 end) as _6,(case when V._7='0%' then '' else V._7 end) as _7,(case when V._8='0%' then '' else V._8 end) as _8,(case when V._9='0%' then '' else V._9 end) as _9,(case when V._10='0%' then '' else V._10 end) as _10,(case when V._11='0%' then '' else V._11 end) as _11,(case when V._12='0%' then '' else V._12 end) as _12,(case when V._13='0%' then '' else V._13 end) as _13,(case when V._14='0%' then '' else V._14 end) as _14,(case when V._15='0%' then '' else V._15 end) as _15,(case when V._16='0%' then '' else V._16 end) as _16,(case when V._17='0%' then '' else V._17 end) as _17,(case when V._18='0%' then '' else V._18 end) as _18,(case when V._19='0%' then '' else V._19 end) as _19,(case when V._20='0%' then '' else V._20 end) as _20,(case when V._21='0%' then '' else V._21 end) as _21,(case when V._22='0%' then '' else V._22 end) as _22,(case when V._23='0%' then '' else V._23 end) as _23,(case when V._24='0%' then '' else V._24 end) as _24,(case when V._25='0%' then '' else V._25 end) as _25 from(select V.Inicio, V.Sem, cast(V._35 as varchar)+'%' as _35,cast(V._36 as varchar)+'%' as _36,cast(V._37 as varchar)+'%' as _37,cast(V._38 as varchar)+'%' as _38,cast(V._39 as varchar)+'%' as _39,cast(V._40 as varchar)+'%' as _40,cast(V._41 as varchar)+'%' as _41,cast(V._42 as varchar)+'%' as _42,cast(V._43 as varchar)+'%' as _43,cast(V._44 as varchar)+'%' as _44,cast(V._45 as varchar)+'%' as _45,cast(V._46 as varchar)+'%' as _46,cast(V._47 as varchar)+'%' as _47,cast(V._48 as varchar)+'%' as _48,cast(V._49 as varchar)+'%' as _49,cast(V._50 as varchar)+'%' as _50,cast(V._51 as varchar)+'%' as _51,cast(V._52 as varchar)+'%' as _52,cast(V._1 as varchar)+'%' as _1,cast(V._2 as varchar)+'%' _2,cast(V._3 as varchar)+'%' as _3,cast(V._4 as varchar)+'%' as _4,cast(V._5 as varchar)+'%' as _5,cast(V._6 as varchar)+'%' as _6,cast(V._7 as varchar)+'%' as _7,cast(V._8 as varchar)+'%' as _8,cast(V._9 as varchar)+'%' as _9,cast(V._10 as varchar)+'%' as _10,cast(V._11 as varchar)+'%' as _11,cast(V._12 as varchar)+'%' as _12,cast(V._13 as varchar)+'%' as _13,cast(V._14 as varchar)+'%' as _14,cast(V._15 as varchar)+'%' as _15,cast(V._16 as varchar)+'%' as _16,cast(V._17 as varchar)+'%' as _17,cast(V._18 as varchar)+'%' as _18,cast(V._19 as varchar)+'%' as _19,cast(V._20 as varchar)+'%'as _20,cast(V._21 as varchar)+'%' as _21,cast(V._22 as varchar)+'%' as _22,cast(V._23 as varchar)+'%' as _23,cast(V._24 as varchar)+'%' as _24,cast(V._25 as varchar)+'%' as _25 from (SELECT V.Inicio, V.Sem,(case when A._35=0 then '' else case when V._35=0 then '' else round((A._35/V._35)*100,0) end end) as _35,(case when A._36=0 then 0 else case when V._36=0 then 0 else round((A._36/V._36)*100,0) end end) as _36,(case when A._37=0 then 0 else case when V._37=0 then 0 else round((A._37/V._37)*100,0) end end) as _37,(case when A._38=0 then 0 else case when V._38=0 then 0 else round((A._38/V._38)*100,0) end end) as _38,(case when A._39=0 then 0 else case when V._39=0 then 0 else round((A._39/V._39)*100,0) end end) as _39,(case when A._40=0 then 0 else case when V._40=0 then 0 else round((A._40/V._40)*100,0) end end) as _40,(case when A._41=0 then 0 else case when V._41=0 then 0 else round((A._41/V._41)*100,0) end end) as _41,(case when A._42=0 then 0 else case when V._42=0 then 0 else round((A._42/V._42)*100,0) end end) as _42,(case when A._43=0 then 0 else case when V._43=0 then 0 else round((A._43/V._43)*100,0) end end) as _43,(case when A._44=0 then 0 else case when V._44=0 then 0 else round((A._44/V._44)*100,0) end end) as _44,(case when A._45=0 then 0 else case when V._45=0 then 0 else round((A._45/V._45)*100,0) end end) as _45,(case when A._46=0 then 0 else case when V._46=0 then 0 else round((A._46/V._46)*100,0) end end) as _46,(case when A._47=0 then 0 else case when V._47=0 then 0 else round((A._47/V._47)*100,0) end end) as _47,(case when A._48=0 then 0 else case when V._48=0 then 0 else round((A._48/V._48)*100,0) end end) as _48,(case when A._49=0 then 0 else case when V._49=0 then 0 else round((A._49/V._49)*100,0) end end) as _49,(case when A._50=0 then 0 else case when V._50=0 then 0 else round((A._50/V._50)*100,0) end end) as _50,(case when A._51=0 then 0 else case when V._51=0 then 0 else round((A._51/V._51)*100,0) end end) as _51,(case when A._52=0 then 0 else case when V._52=0 then 0 else round((A._52/V._52)*100,0) end end) as _52,(case when A._1=0 then 0 else case when V._1=0 then 0 else round((A._1/V._1)*100,0) end end) as _1,(case when A._2=0 then 0 else case when V._2=0 then 0 else round((A._2/V._2)*100,0) end end) as _2,(case when A._3=0 then 0 else case when V._3=0 then 0 else round((A._3/V._3)*100,0) end end) as _3,(case when A._4=0 then 0 else case when V._4=0 then 0 else round((A._4/V._4)*100,0) end end) as _4,(case when A._5=0 then 0 else case when V._5=0 then 0 else round((A._5/V._5)*100,0) end end) as _5,(case when A._6=0 then 0 else case when V._6=0 then 0 else round((A._6/V._6)*100,0) end end) as _6,(case when A._7=0 then 0 else case when V._7=0 then 0 else round((A._7/V._7)*100,0) end end) as _7,(case when A._8=0 then 0 else case when V._8=0 then 0 else round((A._8/V._8)*100,0) end end) as _8,(case when A._9=0 then 0 else case when V._9=0 then 0 else round((A._9/V._9)*100,0) end end) as _9,(case when A._10=0 then 0 else case when V._10=0 then 0 else round((A._10/V._10)*100,0) end end) as _10,(case when A._11=0 then 0 else case when V._11=0 then 0 else round((A._11/V._11)*100,0) end end) as _11,(case when A._12=0 then 0 else case when V._12=0 then 0 else round((A._12/V._12)*100,0) end end) as _12,(case when A._13=0 then 0 else case when V._13=0 then 0 else round((A._13/V._13)*100,0) end end) as _13,(case when A._14=0 then 0 else case when V._14=0 then 0 else round((A._14/V._14)*100,0) end end) as _14,(case when A._15=0 then 0 else case when V._15=0 then 0 else round((A._15/V._15)*100,0) end end) as _15,(case when A._16=0 then 0 else case when V._16=0 then 0 else round((A._16/V._16)*100,0) end end) as _16,(case when A._17=0 then 0 else case when V._17=0 then 0 else round((A._17/V._17)*100,0) end end) as _17,(case when A._18=0 then 0 else case when V._18=0 then 0 else round((A._18/V._18)*100,0) end end) as _18,(case when A._19=0 then 0 else case when V._19=0 then 0 else round((A._19/V._19)*100,0) end end) as _19,(case when A._20=0 then 0 else case when V._20=0 then 0 else round((A._20/V._20)*100,0) end end) as _20,(case when A._21=0 then 0 else case when V._21=0 then 0 else round((A._21/V._21)*100,0) end end) as _21,(case when A._22=0 then 0 else case when V._22=0 then 0 else round((A._22/V._22)*100,0) end end) as _22,(case when A._23=0 then 0 else case when V._23=0 then 0 else round((A._23/V._23)*100,0) end end) as _23,(case when A._24=0 then 0 else case when V._24=0 then 0 else round((A._24/V._24)*100,0) end end) as _24,(case when A._25=0 then 0 else case when V._25=0 then 0 else round((A._25/V._25)*100,0) end end) as _25 FROM(SELECT V.Temporada, isnull(V.[35],0) as _35,isnull(V.[36],0) as _36,isnull(V.[37],0) as _37,isnull(V.[38],0) as _38,isnull(V.[39],0) as _39,isnull(V.[40],0) as _40,isnull(V.[41],0) as _41,isnull(V.[42],0) as _42,isnull(V.[43],0) as _43,isnull(V.[44],0) as _44,isnull(V.[45],0) as _45,isnull(V.[46],0) as _46,isnull(V.[47],0) as _47,isnull(V.[48],0) as _48,isnull(V.[49],0) as _49,isnull(V.[50],0) as _50,isnull(V.[51],0) as _51,isnull(V.[52],0) as _52,isnull(V.[1],0) as _1,isnull(V.[2],0) as _2,isnull(V.[3],0) as _3,isnull(V.[4],0) as _4,isnull(V.[5],0) as _5,isnull(V.[6],0) as _6,isnull(V.[7],0) as _7,isnull(V.[8],0) as _8,isnull(V.[9],0) as _9,isnull(V.[10],0) as _10,isnull(V.[11],0) as _11,isnull(V.[12],0) as _12,isnull(V.[13],0) as _13,isnull(V.[14],0) as _14,isnull(V.[15],0) as _15,isnull(V.[16],0) as _16,isnull(V.[17],0) as _17,isnull(V.[18],0) as _18,isnull(V.[19],0) as _19,isnull(V.[20],0) as _20,isnull(V.[21],0) as _21,isnull(V.[22],0) as _22,isnull(V.[23],0) as _23,isnull(V.[24],0) as _24,isnull(V.[25],0) as _25 FROM(Select * from(SELECT Temporada, Semana, SUM(Convertidas) AS Convertidas FROM SEasonsun1213..UV_ProdRecepcion where CodEstatus<> 'C' and temporada = (select temporada from CatSemanas where GETDATE() between Inicio and Fin) and " +
           "DescProducto = 'FRAMBUESA' GROUP BY Temporada, Semana Union All SELECT Temporada, Semana, SUM(Convertidas) AS Convertidas FROM SeasonPlan..UV_ProdRecepcion where CodEstatus <> 'C' and temporada = (select temporada from CatSemanas where GETDATE() between Inicio and Fin) and " +
           "DescProducto = 'FRAMBUESA' GROUP BY Temporada, Semana)V PIVOT(SUM(Convertidas) FOR Semana in ([35],[36],[37],[38],[39],[40],[41],[42],[43],[44],[45],[46],[47],[48],[49],[50],[51],[52],[1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],[17],[18],[19],[20],[21],[22],[23],[24],[25]))X)V)A left join(Select V.Inicio, V.Temporada, V.Sem, isnull(V.[35], 0) as _35, isnull(V.[36], 0) as _36, isnull(V.[37], 0) as _37, isnull(V.[38], 0) as _38, isnull(V.[39], 0) as _39, isnull(V.[40], 0) as _40, isnull(V.[41], 0) as _41, isnull(V.[42], 0) as _42, isnull(V.[43], 0) as _43, isnull(V.[44], 0) as _44, isnull(V.[45], 0) as _45, isnull(V.[46], 0) as _46, isnull(V.[47], 0) as _47, isnull(V.[48], 0) as _48, isnull(V.[49], 0) as _49, isnull(V.[50], 0) as _50, isnull(V.[51], 0) as _51, isnull(V.[52], 0) as _52, isnull(V.[1], 0) as _1, isnull(V.[2], 0) as _2, isnull(V.[3], 0) as _3, isnull(V.[4], 0) as _4, isnull(V.[5], 0) as _5, isnull(V.[6], 0) as _6, isnull(V.[7], 0) as _7, isnull(V.[8], 0) as _8, isnull(V.[9], 0) as _9, isnull(V.[10], 0) as _10, isnull(V.[11], 0) as _11, isnull(V.[12], 0) as _12, isnull(V.[13], 0) as _13, isnull(V.[14], 0) as _14, isnull(V.[15], 0) as _15, isnull(V.[16], 0) as _16, isnull(V.[17], 0) as _17, isnull(V.[18], 0) as _18, isnull(V.[19], 0) as _19, isnull(V.[20], 0) as _20, isnull(V.[21], 0) as _21, isnull(V.[22], 0) as _22, isnull(V.[23], 0) as _23, isnull(V.[24], 0) as _24, isnull(V.[25], 0) as _25 from(select * from(SELECT S.Semana AS Sem, S.Inicio, E.Semanas, E.Cantidad, E.Temporada FROM Estimacion_Berries E left join CatSemanas S on E.Temporada = S.Temporada AND E.Semana = S.Semana where E.Temporada = (select temporada from CatSemanas where GETDATE() between Inicio and Fin) and " +
           "E.Concepto = 'PROYECCION SEMANAL' and E.Cultivo = 2)V PIVOT(SUM(Cantidad) For Semanas in ([35],[36],[37],[38],[39],[40],[41],[42],[43],[44],[45],[46],[47],[48],[49],[50],[51],[52],[1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],[17],[18],[19],[20],[21],[22],[23],[24],[25]))X)V)V on A.Temporada = V.Temporada)V)V ORDER BY V.Inicio").ToList();

            int xPSF = 50;
            foreach (var item in pSF)
            {
                wsF.Cells[string.Format("A{0}", xPSF)].Value = item.SEMANA;
                wsF.Cells[string.Format("B{0}", xPSF)].Value = item._37;
                wsF.Cells[string.Format("C{0}", xPSF)].Value = item._38;
                wsF.Cells[string.Format("D{0}", xPSF)].Value = item._39;
                wsF.Cells[string.Format("E{0}", xPSF)].Value = item._40;
                wsF.Cells[string.Format("F{0}", xPSF)].Value = item._41;
                wsF.Cells[string.Format("G{0}", xPSF)].Value = item._42;
                wsF.Cells[string.Format("H{0}", xPSF)].Value = item._43;
                wsF.Cells[string.Format("I{0}", xPSF)].Value = item._44;
                wsF.Cells[string.Format("J{0}", xPSF)].Value = item._45;
                wsF.Cells[string.Format("K{0}", xPSF)].Value = item._46;
                wsF.Cells[string.Format("L{0}", xPSF)].Value = item._47;
                wsF.Cells[string.Format("M{0}", xPSF)].Value = item._48;
                wsF.Cells[string.Format("N{0}", xPSF)].Value = item._49;
                wsF.Cells[string.Format("O{0}", xPSF)].Value = item._50;
                wsF.Cells[string.Format("P{0}", xPSF)].Value = item._51;
                wsF.Cells[string.Format("Q{0}", xPSF)].Value = item._52;
                wsF.Cells[string.Format("R{0}", xPSF)].Value = item._1;
                wsF.Cells[string.Format("S{0}", xPSF)].Value = item._2;
                wsF.Cells[string.Format("T{0}", xPSF)].Value = item._3;
                wsF.Cells[string.Format("U{0}", xPSF)].Value = item._4;
                wsF.Cells[string.Format("V{0}", xPSF)].Value = item._5;
                wsF.Cells[string.Format("W{0}", xPSF)].Value = item._6;
                wsF.Cells[string.Format("X{0}", xPSF)].Value = item._7;
                wsF.Cells[string.Format("Y{0}", xPSF)].Value = item._8;
                wsF.Cells[string.Format("Z{0}", xPSF)].Value = item._9;
                wsF.Cells[string.Format("AA{0}", xPSF)].Value = item._10;
                wsF.Cells[string.Format("AB{0}", xPSF)].Value = item._11;
                wsF.Cells[string.Format("AC{0}", xPSF)].Value = item._12;
                wsF.Cells[string.Format("AD{0}", xPSF)].Value = item._13;
                wsF.Cells[string.Format("AE{0}", xPSF)].Value = item._14;
                wsF.Cells[string.Format("AF{0}", xPSF)].Value = item._15;
                wsF.Cells[string.Format("AG{0}", xPSF)].Value = item._16;
                wsF.Cells[string.Format("AH{0}", xPSF)].Value = item._17;
                wsF.Cells[string.Format("AI{0}", xPSF)].Value = item._18;
                wsF.Cells[string.Format("AJ{0}", xPSF)].Value = item._19;
                wsF.Cells[string.Format("AK{0}", xPSF)].Value = item._20;
                wsF.Cells[string.Format("AL{0}", xPSF)].Value = item._21;
                wsF.Cells[string.Format("AM{0}", xPSF)].Value = item._22;
                wsF.Cells[string.Format("AN{0}", xPSF)].Value = item._23;
                wsF.Cells[string.Format("AO{0}", xPSF)].Value = item._24;
                wsF.Cells[string.Format("AP{0}", xPSF)].Value = item._25;
                xPSF++;
            }

            wsF.Cells["A:AP"].AutoFitColumns();

            wsF.Row(2).Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            wsF.Row(2).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(string.Format("yellow")));
            wsF.Row(5).Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            wsF.Row(5).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(string.Format("yellow")));
            wsF.Row(49).Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            wsF.Row(49).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(string.Format("yellow")));

            //----------------------------------------ARANDANO-----------------------------------------------
            ExcelWorksheet wsA = excel.Workbook.Worksheets.Add("ARANDANO");
            wsA.Cells["A1"].Value = "VARIACION SEMANAL DE LA ESTIMACION DE ARANDANO";
            wsA.Cells["A2"].Value = "SEMANA";
            wsA.Cells["B2"].Value = "35";
            wsA.Cells["C2"].Value = "36";
            wsA.Cells["D2"].Value = "37";
            wsA.Cells["E2"].Value = "38";
            wsA.Cells["F2"].Value = "39";
            wsA.Cells["G2"].Value = "40";
            wsA.Cells["H2"].Value = "41";
            wsA.Cells["I2"].Value = "42";
            wsA.Cells["J2"].Value = "43";
            wsA.Cells["K2"].Value = "44";
            wsA.Cells["L2"].Value = "45";
            wsA.Cells["M2"].Value = "46";
            wsA.Cells["N2"].Value = "47";
            wsA.Cells["O2"].Value = "48";
            wsA.Cells["P2"].Value = "49";
            wsA.Cells["Q2"].Value = "50";
            wsA.Cells["R2"].Value = "51";
            wsA.Cells["S2"].Value = "52";
            wsA.Cells["T2"].Value = "1";
            wsA.Cells["U2"].Value = "2";
            wsA.Cells["V2"].Value = "3";
            wsA.Cells["W2"].Value = "4";
            wsA.Cells["X2"].Value = "5";
            wsA.Cells["Y2"].Value = "6";
            wsA.Cells["Z2"].Value = "7";
            wsA.Cells["AA2"].Value = "8";
            wsA.Cells["AB2"].Value = "9";
            wsA.Cells["AC2"].Value = "10";
            wsA.Cells["AD2"].Value = "11";
            wsA.Cells["AE2"].Value = "12";
            wsA.Cells["AF2"].Value = "13";
            wsA.Cells["AG2"].Value = "14";
            wsA.Cells["AH2"].Value = "15";
            wsA.Cells["AI2"].Value = "16";
            wsA.Cells["AJ2"].Value = "17";
            wsA.Cells["AK2"].Value = "18";
            wsA.Cells["AL2"].Value = "19";
            wsA.Cells["AM2"].Value = "20";
            wsA.Cells["AN2"].Value = "21";
            wsA.Cells["AO2"].Value = "22";
            wsA.Cells["AP2"].Value = "23";
            wsA.Cells["AQ2"].Value = "24";
            wsA.Cells["AR2"].Value = "25";

            var estimacion_realA = bd.Database.SqlQuery<R_Berries>("SELECT 'TOTAL' as 'SEMANA', round(isnull(V.[35],0),0) as _35,round(isnull(V.[36],0),0) as _36,round(isnull(V.[37],0),0) as _37,round(isnull(V.[38],0),0) as _38,round(isnull(V.[39],0),0) as _39,round(isnull(V.[40],0),0) as _40,round(isnull(V.[41],0),0) as _41,round(isnull(V.[42],0),0) as _42,round(isnull(V.[43],0),0) as _43,round(isnull(V.[44],0),0) as _44,round(isnull(V.[45],0),0) as _45,round(isnull(V.[46],0),0) as _46,round(isnull(V.[47],0),0) as _47,round(isnull(V.[48],0),0) as _48,round(isnull(V.[49],0),0) as _49,round(isnull(V.[50],0),0) as _50,round(isnull(V.[51],0),0) as _51,round(isnull(V.[52],0),0) as _52,round(isnull(V.[1],0),0) as _1,round(isnull(V.[2],0),0) as _2,round(isnull(V.[3],0),0) as _3,round(isnull(V.[4],0),0) as _4,round(isnull(V.[5],0),0) as _5,round(isnull(V.[6],0),0) as _6,round(isnull(V.[7],0),0) as _7,round(isnull(V.[8],0),0) as _8,round(isnull(V.[9],0),0) as _9,round(isnull(V.[10],0),0) as _10,round(isnull(V.[11],0),0) as _11,round(isnull(V.[12],0),0) as _12,round(isnull(V.[13],0),0) as _13,round(isnull(V.[14],0),0) as _14,round(isnull(V.[15],0),0) as _15,round(isnull(V.[16],0),0) as _16,round(isnull(V.[17],0),0) as _17,round(isnull(V.[18],0),0) as _18,round(isnull(V.[19],0),0) as _19,round(isnull(V.[20],0),0) as _20,round(isnull(V.[21],0),0) as _21,round(isnull(V.[22],0),0) as _22,round(isnull(V.[23],0),0) as _23,round(isnull(V.[24],0),0) as _24,round(isnull(V.[25],0),0) as _25 " +
                "FROM(Select * from(SELECT Semana, SUM(Convertidas) AS Convertidas FROM SEasonsun1213..UV_ProdRecepcion where CodEstatus <> 'C' and temporada = (select temporada from CatSemanas where GETDATE() between Inicio and Fin) and DescProducto = 'ARANDANO' GROUP BY Semana " +
                "Union All SELECT Semana, SUM(Convertidas) AS Convertidas FROM SeasonPlan..UV_ProdRecepcion where CodEstatus <> 'C' and temporada = (select temporada from CatSemanas where GETDATE() between Inicio and Fin) and DescProducto = 'ARANDANO' GROUP BY Semana)V PIVOT(SUM(Convertidas) FOR Semana in ([35],[36],[37],[38],[39],[40],[41],[42],[43],[44],[45],[46],[47],[48],[49],[50],[51],[52],[1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],[17],[18],[19],[20],[21],[22],[23],[24],[25]))X)V").ToList();

            int xRA = 3;
            foreach (var item in estimacion_realA)
            {
                wsA.Cells[string.Format("A{0}", xRA)].Value = "RECEPCION REAL";
                wsA.Cells[string.Format("B{0}", xRA)].Value = item._35;
                wsA.Cells[string.Format("C{0}", xRA)].Value = item._36;
                wsA.Cells[string.Format("D{0}", xRA)].Value = item._37;
                wsA.Cells[string.Format("E{0}", xRA)].Value = item._38;
                wsA.Cells[string.Format("F{0}", xRA)].Value = item._39;
                wsA.Cells[string.Format("G{0}", xRA)].Value = item._40;
                wsA.Cells[string.Format("H{0}", xRA)].Value = item._41;
                wsA.Cells[string.Format("I{0}", xRA)].Value = item._42;
                wsA.Cells[string.Format("J{0}", xRA)].Value = item._43;
                wsA.Cells[string.Format("K{0}", xRA)].Value = item._44;
                wsA.Cells[string.Format("L{0}", xRA)].Value = item._45;
                wsA.Cells[string.Format("M{0}", xRA)].Value = item._46;
                wsA.Cells[string.Format("N{0}", xRA)].Value = item._47;
                wsA.Cells[string.Format("O{0}", xRA)].Value = item._48;
                wsA.Cells[string.Format("P{0}", xRA)].Value = item._49;
                wsA.Cells[string.Format("Q{0}", xRA)].Value = item._50;
                wsA.Cells[string.Format("R{0}", xRA)].Value = item._51;
                wsA.Cells[string.Format("S{0}", xRA)].Value = item._52;
                wsA.Cells[string.Format("T{0}", xRA)].Value = item._1;
                wsA.Cells[string.Format("U{0}", xRA)].Value = item._2;
                wsA.Cells[string.Format("V{0}", xRA)].Value = item._3;
                wsA.Cells[string.Format("W{0}", xRA)].Value = item._4;
                wsA.Cells[string.Format("X{0}", xRA)].Value = item._5;
                wsA.Cells[string.Format("Y{0}", xRA)].Value = item._6;
                wsA.Cells[string.Format("Z{0}", xRA)].Value = item._7;
                wsA.Cells[string.Format("AA{0}", xRA)].Value = item._8;
                wsA.Cells[string.Format("AB{0}", xRA)].Value = item._9;
                wsA.Cells[string.Format("AC{0}", xRA)].Value = item._10;
                wsA.Cells[string.Format("AD{0}", xRA)].Value = item._11;
                wsA.Cells[string.Format("AE{0}", xRA)].Value = item._12;
                wsA.Cells[string.Format("AF{0}", xRA)].Value = item._13;
                wsA.Cells[string.Format("AG{0}", xRA)].Value = item._14;
                wsA.Cells[string.Format("AH{0}", xRA)].Value = item._15;
                wsA.Cells[string.Format("AI{0}", xRA)].Value = item._16;
                wsA.Cells[string.Format("AJ{0}", xRA)].Value = item._17;
                wsA.Cells[string.Format("AK{0}", xRA)].Value = item._18;
                wsA.Cells[string.Format("AL{0}", xRA)].Value = item._19;
                wsA.Cells[string.Format("AM{0}", xRA)].Value = item._20;
                wsA.Cells[string.Format("AN{0}", xRA)].Value = item._21;
                wsA.Cells[string.Format("AO{0}", xRA)].Value = item._22;
                wsA.Cells[string.Format("AP{0}", xRA)].Value = item._23;
                wsA.Cells[string.Format("AQ{0}", xRA)].Value = item._24;
                wsA.Cells[string.Format("AR{0}", xRA)].Value = item._25;
                xRA++;
            }

            var recepcion_europaA = bd.Database.SqlQuery<E_Berries>("Select round(isnull(V.[35],0),0) as _35,round(isnull(V.[36],0),0) as _36,round(isnull(V.[37],0),0) as _37,round(isnull(V.[38],0),0) as _38,round(isnull(V.[39],0),0) as _39,round(isnull(V.[40],0),0) as _40,round(isnull(V.[41],0),0) as _41,round(isnull(V.[42],0),0) as _42,round(isnull(V.[43],0),0) as _43,round(isnull(V.[44],0),0) as _44,round(isnull(V.[45],0),0) as _45,round(isnull(V.[46],0),0) as _46,round(isnull(V.[47],0),0) as _47,round(isnull(V.[48],0),0) as _48,round(isnull(V.[49],0),0) as _49,round(isnull(V.[50],0),0) as _50,round(isnull(V.[51],0),0) as _51,round(isnull(V.[52],0),0) as _52,round(isnull(V.[1],0),0) as _1,round(isnull(V.[2],0),0) as _2,round(isnull(V.[3],0),0) as _3,round(isnull(V.[4],0),0) as _4,round(isnull(V.[5],0),0) as _5,round(isnull(V.[6],0),0) as _6,round(isnull(V.[7],0),0) as _7,round(isnull(V.[8],0),0) as _8,round(isnull(V.[9],0),0) as _9,round(isnull(V.[10],0),0) as _10,round(isnull(V.[11],0),0) as _11,round(isnull(V.[12],0),0) as _12,round(isnull(V.[13],0),0) as _13,round(isnull(V.[14],0),0) as _14,round(isnull(V.[15],0),0) as _15,round(isnull(V.[16],0),0) as _16,round(isnull(V.[17],0),0) as _17,round(isnull(V.[18],0),0) as _18,round(isnull(V.[19],0),0) as _19,round(isnull(V.[20],0),0) as _20,round(isnull(V.[21],0),0) as _21,round(isnull(V.[22],0),0) as _22,round(isnull(V.[23],0),0) as _23,round(isnull(V.[24],0),0) as _24,round(isnull(V.[25],0),0) as _25 " +
                "from(Select * from(SELECT E.Concepto, S.Inicio, E.Semanas, E.Cantidad FROM Estimacion_Berries E " +
                "left join CatSemanas S on E.Temporada = S.Temporada AND E.Semana = S.Semana " +
                "where E.Temporada = (select temporada from CatSemanas where GETDATE() between Inicio and Fin) and E.Concepto = 'EUROPA' and E.Cultivo = 3)V " +
                "PIVOT(SUM(Cantidad) For Semanas in ([35],[36],[37],[38],[39],[40],[41],[42],[43],[44],[45],[46],[47],[48],[49],[50],[51],[52],[1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],[17],[18],[19],[20],[21],[22],[23],[24],[25]))X)V ORDER BY V.Inicio").ToList();

            int xEA = 4;
            foreach (var item in recepcion_europaA)
            {
                wsA.Cells[string.Format("A{0}", xEA)].Value = "RECEPCION EUROPA";
                wsA.Cells[string.Format("B{0}", xEA)].Value = item._35;
                wsA.Cells[string.Format("C{0}", xEA)].Value = item._36;
                wsA.Cells[string.Format("D{0}", xEA)].Value = item._37;
                wsA.Cells[string.Format("E{0}", xEA)].Value = item._38;
                wsA.Cells[string.Format("F{0}", xEA)].Value = item._39;
                wsA.Cells[string.Format("G{0}", xEA)].Value = item._40;
                wsA.Cells[string.Format("H{0}", xEA)].Value = item._41;
                wsA.Cells[string.Format("I{0}", xEA)].Value = item._42;
                wsA.Cells[string.Format("J{0}", xEA)].Value = item._43;
                wsA.Cells[string.Format("K{0}", xEA)].Value = item._44;
                wsA.Cells[string.Format("L{0}", xEA)].Value = item._45;
                wsA.Cells[string.Format("M{0}", xEA)].Value = item._46;
                wsA.Cells[string.Format("N{0}", xEA)].Value = item._47;
                wsA.Cells[string.Format("O{0}", xEA)].Value = item._48;
                wsA.Cells[string.Format("P{0}", xEA)].Value = item._49;
                wsA.Cells[string.Format("Q{0}", xEA)].Value = item._50;
                wsA.Cells[string.Format("R{0}", xEA)].Value = item._51;
                wsA.Cells[string.Format("S{0}", xEA)].Value = item._52;
                wsA.Cells[string.Format("T{0}", xEA)].Value = item._1;
                wsA.Cells[string.Format("U{0}", xEA)].Value = item._2;
                wsA.Cells[string.Format("V{0}", xEA)].Value = item._3;
                wsA.Cells[string.Format("W{0}", xEA)].Value = item._4;
                wsA.Cells[string.Format("X{0}", xEA)].Value = item._5;
                wsA.Cells[string.Format("Y{0}", xEA)].Value = item._6;
                wsA.Cells[string.Format("Z{0}", xEA)].Value = item._7;
                wsA.Cells[string.Format("AA{0}", xEA)].Value = item._8;
                wsA.Cells[string.Format("AB{0}", xEA)].Value = item._9;
                wsA.Cells[string.Format("AC{0}", xEA)].Value = item._10;
                wsA.Cells[string.Format("AD{0}", xEA)].Value = item._11;
                wsA.Cells[string.Format("AE{0}", xEA)].Value = item._12;
                wsA.Cells[string.Format("AF{0}", xEA)].Value = item._13;
                wsA.Cells[string.Format("AG{0}", xEA)].Value = item._14;
                wsA.Cells[string.Format("AH{0}", xEA)].Value = item._15;
                wsA.Cells[string.Format("AI{0}", xEA)].Value = item._16;
                wsA.Cells[string.Format("AJ{0}", xEA)].Value = item._17;
                wsA.Cells[string.Format("AK{0}", xEA)].Value = item._18;
                wsA.Cells[string.Format("AL{0}", xEA)].Value = item._19;
                wsA.Cells[string.Format("AM{0}", xEA)].Value = item._20;
                wsA.Cells[string.Format("AN{0}", xEA)].Value = item._21;
                wsA.Cells[string.Format("AO{0}", xEA)].Value = item._22;
                wsA.Cells[string.Format("AP{0}", xEA)].Value = item._23;
                wsA.Cells[string.Format("AQ{0}", xEA)].Value = item._24;
                wsA.Cells[string.Format("AR{0}", xEA)].Value = item._25;
                xEA++;
            }

            //PROYECCION SEMANAL
            wsA.Cells["A6"].Value = "PROYECCION SEMANAL";
            wsA.Cells["B6"].Value = "35";
            wsA.Cells["C6"].Value = "36";
            wsA.Cells["D6"].Value = "37";
            wsA.Cells["E6"].Value = "38";
            wsA.Cells["F6"].Value = "39";
            wsA.Cells["G6"].Value = "40";
            wsA.Cells["H6"].Value = "41";
            wsA.Cells["I6"].Value = "42";
            wsA.Cells["J6"].Value = "43";
            wsA.Cells["K6"].Value = "44";
            wsA.Cells["L6"].Value = "45";
            wsA.Cells["M6"].Value = "46";
            wsA.Cells["N6"].Value = "47";
            wsA.Cells["O6"].Value = "48";
            wsA.Cells["P6"].Value = "49";
            wsA.Cells["Q6"].Value = "50";
            wsA.Cells["R6"].Value = "51";
            wsA.Cells["S6"].Value = "52";
            wsA.Cells["T6"].Value = "1";
            wsA.Cells["U6"].Value = "2";
            wsA.Cells["V6"].Value = "3";
            wsA.Cells["W6"].Value = "4";
            wsA.Cells["X6"].Value = "5";
            wsA.Cells["Y6"].Value = "6";
            wsA.Cells["Z6"].Value = "7";
            wsA.Cells["AA6"].Value = "8";
            wsA.Cells["AB6"].Value = "9";
            wsA.Cells["AC6"].Value = "10";
            wsA.Cells["AD6"].Value = "11";
            wsA.Cells["AE6"].Value = "12";
            wsA.Cells["AF6"].Value = "13";
            wsA.Cells["AG6"].Value = "14";
            wsA.Cells["AH6"].Value = "15";
            wsA.Cells["AI6"].Value = "16";
            wsA.Cells["AJ6"].Value = "17";
            wsA.Cells["AK6"].Value = "18";
            wsA.Cells["AL6"].Value = "19";
            wsA.Cells["AM6"].Value = "20";
            wsA.Cells["AN6"].Value = "21";
            wsA.Cells["AO6"].Value = "22";
            wsA.Cells["AP6"].Value = "23";
            wsA.Cells["AQ6"].Value = "24";
            wsA.Cells["AR6"].Value = "25";

            var dataSA = bd.Database.SqlQuery<E_Berries>("Select V.Sem AS 'SEMANA', round(isnull(V.[35],0),0) as _35, round(isnull(V.[36],0),0) as _36, round(isnull(V.[37],0),0) as _37,round(isnull(V.[38],0),0) as _38,round(isnull(V.[39],0),0) as _39,round(isnull(V.[40],0),0) as _40,round(isnull(V.[41],0),0) as _41,round(isnull(V.[42],0),0) as _42,round(isnull(V.[43],0),0) as _43,round(isnull(V.[44],0),0) as _44,round(isnull(V.[45],0),0) as _45,round(isnull(V.[46],0),0) as _46,round(isnull(V.[47],0),0) as _47,round(isnull(V.[48],0),0) as _48,round(isnull(V.[49],0),0) as _49,round(isnull(V.[50],0),0) as _50,round(isnull(V.[51],0),0) as _51,round(isnull(V.[52],0),0) as _52,round(isnull(V.[1],0),0) as _1,round(isnull(V.[2],0),0) as _2,round(isnull(V.[3],0),0) as _3,round(isnull(V.[4],0),0) as _4,round(isnull(V.[5],0),0) as _5,round(isnull(V.[6],0),0) as _6,round(isnull(V.[7],0),0) as _7,round(isnull(V.[8],0),0) as _8,round(isnull(V.[9],0),0) as _9,round(isnull(V.[10],0),0) as _10,round(isnull(V.[11],0),0) as _11,round(isnull(V.[12],0),0) as _12,round(isnull(V.[13],0),0) as _13,round(isnull(V.[14],0),0) as _14,round(isnull(V.[15],0),0) as _15,round(isnull(V.[16],0),0) as _16,round(isnull(V.[17],0),0) as _17,round(isnull(V.[18],0),0) as _18,round(isnull(V.[19],0),0) as _19,round(isnull(V.[20],0),0) as _20,round(isnull(V.[21],0),0) as _21,round(isnull(V.[22],0),0) as _22,round(isnull(V.[23],0),0) as _23,round(isnull(V.[24],0),0) as _24,round(isnull(V.[25],0),0) as _25 " +
                "from(Select * from(SELECT S.Semana AS Sem, S.Inicio, E.Semanas, E.Cantidad FROM Estimacion_Berries E left join CatSemanas S on E.Temporada = S.Temporada AND E.Semana = S.Semana " +
                "where E.Temporada = (select temporada from CatSemanas where GETDATE() between Inicio and Fin) and E.Concepto = 'PROYECCION SEMANAL' and E.Cultivo = 3)V " +
                "PIVOT(SUM(Cantidad) For Semanas in ([35],[36],[37],[38],[39],[40],[41],[42],[43],[44],[45],[46],[47],[48],[49],[50],[51],[52],[1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],[17],[18],[19],[20],[21],[22],[23],[24],[25]))X)V ORDER BY V.Inicio").ToList();

            int xSA = 7;
            foreach (var item in dataSA)
            {
                wsA.Cells[string.Format("A{0}", xSA)].Value = item.SEMANA;
                wsA.Cells[string.Format("B{0}", xSA)].Value = item._35;
                wsA.Cells[string.Format("C{0}", xSA)].Value = item._36;
                wsA.Cells[string.Format("D{0}", xSA)].Value = item._37;
                wsA.Cells[string.Format("E{0}", xSA)].Value = item._38;
                wsA.Cells[string.Format("F{0}", xSA)].Value = item._39;
                wsA.Cells[string.Format("G{0}", xSA)].Value = item._40;
                wsA.Cells[string.Format("H{0}", xSA)].Value = item._41;
                wsA.Cells[string.Format("I{0}", xSA)].Value = item._42;
                wsA.Cells[string.Format("J{0}", xSA)].Value = item._43;
                wsA.Cells[string.Format("K{0}", xSA)].Value = item._44;
                wsA.Cells[string.Format("L{0}", xSA)].Value = item._45;
                wsA.Cells[string.Format("M{0}", xSA)].Value = item._46;
                wsA.Cells[string.Format("N{0}", xSA)].Value = item._47;
                wsA.Cells[string.Format("O{0}", xSA)].Value = item._48;
                wsA.Cells[string.Format("P{0}", xSA)].Value = item._49;
                wsA.Cells[string.Format("Q{0}", xSA)].Value = item._50;
                wsA.Cells[string.Format("R{0}", xSA)].Value = item._51;
                wsA.Cells[string.Format("S{0}", xSA)].Value = item._52;
                wsA.Cells[string.Format("T{0}", xSA)].Value = item._1;
                wsA.Cells[string.Format("U{0}", xSA)].Value = item._2;
                wsA.Cells[string.Format("V{0}", xSA)].Value = item._3;
                wsA.Cells[string.Format("W{0}", xSA)].Value = item._4;
                wsA.Cells[string.Format("X{0}", xSA)].Value = item._5;
                wsA.Cells[string.Format("Y{0}", xSA)].Value = item._6;
                wsA.Cells[string.Format("Z{0}", xSA)].Value = item._7;
                wsA.Cells[string.Format("AA{0}", xSA)].Value = item._8;
                wsA.Cells[string.Format("AB{0}", xSA)].Value = item._9;
                wsA.Cells[string.Format("AC{0}", xSA)].Value = item._10;
                wsA.Cells[string.Format("AD{0}", xSA)].Value = item._11;
                wsA.Cells[string.Format("AE{0}", xSA)].Value = item._12;
                wsA.Cells[string.Format("AF{0}", xSA)].Value = item._13;
                wsA.Cells[string.Format("AG{0}", xSA)].Value = item._14;
                wsA.Cells[string.Format("AH{0}", xSA)].Value = item._15;
                wsA.Cells[string.Format("AI{0}", xSA)].Value = item._16;
                wsA.Cells[string.Format("AJ{0}", xSA)].Value = item._17;
                wsA.Cells[string.Format("AK{0}", xSA)].Value = item._18;
                wsA.Cells[string.Format("AL{0}", xSA)].Value = item._19;
                wsA.Cells[string.Format("AM{0}", xSA)].Value = item._20;
                wsA.Cells[string.Format("AN{0}", xSA)].Value = item._21;
                wsA.Cells[string.Format("AO{0}", xSA)].Value = item._22;
                wsA.Cells[string.Format("AP{0}", xSA)].Value = item._23;
                wsA.Cells[string.Format("AQ{0}", xSA)].Value = item._24;
                wsA.Cells[string.Format("AR{0}", xSA)].Value = item._25;
                xSA++;
            }

            //PORCENTAJES
            wsA.Cells["A50"].Value = "PORCENTAJES";
            wsA.Cells["B50"].Value = "35";
            wsA.Cells["C50"].Value = "36";
            wsA.Cells["D50"].Value = "37";
            wsA.Cells["E50"].Value = "38";
            wsA.Cells["F50"].Value = "39";
            wsA.Cells["G50"].Value = "40";
            wsA.Cells["H50"].Value = "41";
            wsA.Cells["I50"].Value = "42";
            wsA.Cells["J50"].Value = "43";
            wsA.Cells["K50"].Value = "44";
            wsA.Cells["L50"].Value = "45";
            wsA.Cells["M50"].Value = "46";
            wsA.Cells["N50"].Value = "47";
            wsA.Cells["O50"].Value = "48";
            wsA.Cells["P50"].Value = "49";
            wsA.Cells["Q50"].Value = "50";
            wsA.Cells["R50"].Value = "51";
            wsA.Cells["S50"].Value = "52";
            wsA.Cells["T50"].Value = "1";
            wsA.Cells["U50"].Value = "2";
            wsA.Cells["V50"].Value = "3";
            wsA.Cells["W50"].Value = "4";
            wsA.Cells["X50"].Value = "5";
            wsA.Cells["Y50"].Value = "6";
            wsA.Cells["Z50"].Value = "7";
            wsA.Cells["AA50"].Value = "8";
            wsA.Cells["AB50"].Value = "9";
            wsA.Cells["AC50"].Value = "10";
            wsA.Cells["AD50"].Value = "11";
            wsA.Cells["AE50"].Value = "12";
            wsA.Cells["AF50"].Value = "13";
            wsA.Cells["AG50"].Value = "14";
            wsA.Cells["AH50"].Value = "15";
            wsA.Cells["AI50"].Value = "16";
            wsA.Cells["AJ50"].Value = "17";
            wsA.Cells["AK50"].Value = "18";
            wsA.Cells["AL50"].Value = "19";
            wsA.Cells["AM50"].Value = "20";
            wsA.Cells["AN50"].Value = "21";
            wsA.Cells["AO50"].Value = "22";
            wsA.Cells["AP50"].Value = "23";
            wsA.Cells["AQ50"].Value = "24";
            wsA.Cells["AR50"].Value = "25";

            var pSA = bd.Database.SqlQuery<E_BerriesP>("select V.Sem AS 'SEMANA', (case when V._35='0%' then '' else V._35 end) as _35,(case when V._36='0%' then '' else V._36 end) as _36,(case when V._37='0%' then '' else V._37 end) as _37,(case when V._38='0%' then '' else V._38 end) as _38,(case when V._39='0%' then '' else V._39 end) as _39,(case when V._40='0%' then '' else V._40 end) as _40,(case when V._41='0%' then '' else V._41 end) as _41,(case when V._42='0%' then '' else V._42 end) as _42,(case when V._43='0%' then '' else V._43 end) as _43,(case when V._44='0%' then '' else V._44 end) as _44,(case when V._45='0%' then '' else V._45 end) as _45,(case when V._46='0%' then '' else V._46 end) as _46,(case when V._47='0%' then '' else V._47 end) as _47,(case when V._48='0%' then '' else V._48 end) as _48,(case when V._49='0%' then '' else V._49 end) as _49,(case when V._50='0%' then '' else V._50 end) as _50,(case when V._51='0%' then '' else V._51 end) as _51,(case when V._52='0%' then '' else V._52 end) as _52,(case when V._1='0%' then '' else V._1 end) as _1,(case when V._2='0%' then '' else V._2 end) as _2,(case when V._3='0%' then '' else V._3 end) as _3,(case when V._4='0%' then '' else V._4 end) as _4,(case when V._5='0%' then '' else V._5 end) as _5,(case when V._6='0%' then '' else V._6 end) as _6,(case when V._7='0%' then '' else V._7 end) as _7,(case when V._8='0%' then '' else V._8 end) as _8,(case when V._9='0%' then '' else V._9 end) as _9,(case when V._10='0%' then '' else V._10 end) as _10,(case when V._11='0%' then '' else V._11 end) as _11,(case when V._12='0%' then '' else V._12 end) as _12,(case when V._13='0%' then '' else V._13 end) as _13,(case when V._14='0%' then '' else V._14 end) as _14,(case when V._15='0%' then '' else V._15 end) as _15,(case when V._16='0%' then '' else V._16 end) as _16,(case when V._17='0%' then '' else V._17 end) as _17,(case when V._18='0%' then '' else V._18 end) as _18,(case when V._19='0%' then '' else V._19 end) as _19,(case when V._20='0%' then '' else V._20 end) as _20,(case when V._21='0%' then '' else V._21 end) as _21,(case when V._22='0%' then '' else V._22 end) as _22,(case when V._23='0%' then '' else V._23 end) as _23,(case when V._24='0%' then '' else V._24 end) as _24,(case when V._25='0%' then '' else V._25 end) as _25 from(select V.Inicio, V.Sem, cast(V._35 as varchar)+'%' as _35,cast(V._36 as varchar)+'%' as _36,cast(V._37 as varchar)+'%' as _37,cast(V._38 as varchar)+'%' as _38,cast(V._39 as varchar)+'%' as _39,cast(V._40 as varchar)+'%' as _40,cast(V._41 as varchar)+'%' as _41,cast(V._42 as varchar)+'%' as _42,cast(V._43 as varchar)+'%' as _43,cast(V._44 as varchar)+'%' as _44,cast(V._45 as varchar)+'%' as _45,cast(V._46 as varchar)+'%' as _46,cast(V._47 as varchar)+'%' as _47,cast(V._48 as varchar)+'%' as _48,cast(V._49 as varchar)+'%' as _49,cast(V._50 as varchar)+'%' as _50,cast(V._51 as varchar)+'%' as _51,cast(V._52 as varchar)+'%' as _52,cast(V._1 as varchar)+'%' as _1,cast(V._2 as varchar)+'%' _2,cast(V._3 as varchar)+'%' as _3,cast(V._4 as varchar)+'%' as _4,cast(V._5 as varchar)+'%' as _5,cast(V._6 as varchar)+'%' as _6,cast(V._7 as varchar)+'%' as _7,cast(V._8 as varchar)+'%' as _8,cast(V._9 as varchar)+'%' as _9,cast(V._10 as varchar)+'%' as _10,cast(V._11 as varchar)+'%' as _11,cast(V._12 as varchar)+'%' as _12,cast(V._13 as varchar)+'%' as _13,cast(V._14 as varchar)+'%' as _14,cast(V._15 as varchar)+'%' as _15,cast(V._16 as varchar)+'%' as _16,cast(V._17 as varchar)+'%' as _17,cast(V._18 as varchar)+'%' as _18,cast(V._19 as varchar)+'%' as _19,cast(V._20 as varchar)+'%'as _20,cast(V._21 as varchar)+'%' as _21,cast(V._22 as varchar)+'%' as _22,cast(V._23 as varchar)+'%' as _23,cast(V._24 as varchar)+'%' as _24,cast(V._25 as varchar)+'%' as _25 from (SELECT V.Inicio, V.Sem,(case when A._35=0 then '' else case when V._35=0 then '' else round((A._35/V._35)*100,0) end end) as _35,(case when A._36=0 then 0 else case when V._36=0 then 0 else round((A._36/V._36)*100,0) end end) as _36,(case when A._37=0 then 0 else case when V._37=0 then 0 else round((A._37/V._37)*100,0) end end) as _37,(case when A._38=0 then 0 else case when V._38=0 then 0 else round((A._38/V._38)*100,0) end end) as _38,(case when A._39=0 then 0 else case when V._39=0 then 0 else round((A._39/V._39)*100,0) end end) as _39,(case when A._40=0 then 0 else case when V._40=0 then 0 else round((A._40/V._40)*100,0) end end) as _40,(case when A._41=0 then 0 else case when V._41=0 then 0 else round((A._41/V._41)*100,0) end end) as _41,(case when A._42=0 then 0 else case when V._42=0 then 0 else round((A._42/V._42)*100,0) end end) as _42,(case when A._43=0 then 0 else case when V._43=0 then 0 else round((A._43/V._43)*100,0) end end) as _43,(case when A._44=0 then 0 else case when V._44=0 then 0 else round((A._44/V._44)*100,0) end end) as _44,(case when A._45=0 then 0 else case when V._45=0 then 0 else round((A._45/V._45)*100,0) end end) as _45,(case when A._46=0 then 0 else case when V._46=0 then 0 else round((A._46/V._46)*100,0) end end) as _46,(case when A._47=0 then 0 else case when V._47=0 then 0 else round((A._47/V._47)*100,0) end end) as _47,(case when A._48=0 then 0 else case when V._48=0 then 0 else round((A._48/V._48)*100,0) end end) as _48,(case when A._49=0 then 0 else case when V._49=0 then 0 else round((A._49/V._49)*100,0) end end) as _49,(case when A._50=0 then 0 else case when V._50=0 then 0 else round((A._50/V._50)*100,0) end end) as _50,(case when A._51=0 then 0 else case when V._51=0 then 0 else round((A._51/V._51)*100,0) end end) as _51,(case when A._52=0 then 0 else case when V._52=0 then 0 else round((A._52/V._52)*100,0) end end) as _52,(case when A._1=0 then 0 else case when V._1=0 then 0 else round((A._1/V._1)*100,0) end end) as _1,(case when A._2=0 then 0 else case when V._2=0 then 0 else round((A._2/V._2)*100,0) end end) as _2,(case when A._3=0 then 0 else case when V._3=0 then 0 else round((A._3/V._3)*100,0) end end) as _3,(case when A._4=0 then 0 else case when V._4=0 then 0 else round((A._4/V._4)*100,0) end end) as _4,(case when A._5=0 then 0 else case when V._5=0 then 0 else round((A._5/V._5)*100,0) end end) as _5,(case when A._6=0 then 0 else case when V._6=0 then 0 else round((A._6/V._6)*100,0) end end) as _6,(case when A._7=0 then 0 else case when V._7=0 then 0 else round((A._7/V._7)*100,0) end end) as _7,(case when A._8=0 then 0 else case when V._8=0 then 0 else round((A._8/V._8)*100,0) end end) as _8,(case when A._9=0 then 0 else case when V._9=0 then 0 else round((A._9/V._9)*100,0) end end) as _9,(case when A._10=0 then 0 else case when V._10=0 then 0 else round((A._10/V._10)*100,0) end end) as _10,(case when A._11=0 then 0 else case when V._11=0 then 0 else round((A._11/V._11)*100,0) end end) as _11,(case when A._12=0 then 0 else case when V._12=0 then 0 else round((A._12/V._12)*100,0) end end) as _12,(case when A._13=0 then 0 else case when V._13=0 then 0 else round((A._13/V._13)*100,0) end end) as _13,(case when A._14=0 then 0 else case when V._14=0 then 0 else round((A._14/V._14)*100,0) end end) as _14,(case when A._15=0 then 0 else case when V._15=0 then 0 else round((A._15/V._15)*100,0) end end) as _15,(case when A._16=0 then 0 else case when V._16=0 then 0 else round((A._16/V._16)*100,0) end end) as _16,(case when A._17=0 then 0 else case when V._17=0 then 0 else round((A._17/V._17)*100,0) end end) as _17,(case when A._18=0 then 0 else case when V._18=0 then 0 else round((A._18/V._18)*100,0) end end) as _18,(case when A._19=0 then 0 else case when V._19=0 then 0 else round((A._19/V._19)*100,0) end end) as _19,(case when A._20=0 then 0 else case when V._20=0 then 0 else round((A._20/V._20)*100,0) end end) as _20,(case when A._21=0 then 0 else case when V._21=0 then 0 else round((A._21/V._21)*100,0) end end) as _21,(case when A._22=0 then 0 else case when V._22=0 then 0 else round((A._22/V._22)*100,0) end end) as _22,(case when A._23=0 then 0 else case when V._23=0 then 0 else round((A._23/V._23)*100,0) end end) as _23,(case when A._24=0 then 0 else case when V._24=0 then 0 else round((A._24/V._24)*100,0) end end) as _24,(case when A._25=0 then 0 else case when V._25=0 then 0 else round((A._25/V._25)*100,0) end end) as _25 FROM(SELECT V.Temporada, isnull(V.[35],0) as _35,isnull(V.[36],0) as _36,isnull(V.[37],0) as _37,isnull(V.[38],0) as _38,isnull(V.[39],0) as _39,isnull(V.[40],0) as _40,isnull(V.[41],0) as _41,isnull(V.[42],0) as _42,isnull(V.[43],0) as _43,isnull(V.[44],0) as _44,isnull(V.[45],0) as _45,isnull(V.[46],0) as _46,isnull(V.[47],0) as _47,isnull(V.[48],0) as _48,isnull(V.[49],0) as _49,isnull(V.[50],0) as _50,isnull(V.[51],0) as _51,isnull(V.[52],0) as _52,isnull(V.[1],0) as _1,isnull(V.[2],0) as _2,isnull(V.[3],0) as _3,isnull(V.[4],0) as _4,isnull(V.[5],0) as _5,isnull(V.[6],0) as _6,isnull(V.[7],0) as _7,isnull(V.[8],0) as _8,isnull(V.[9],0) as _9,isnull(V.[10],0) as _10,isnull(V.[11],0) as _11,isnull(V.[12],0) as _12,isnull(V.[13],0) as _13,isnull(V.[14],0) as _14,isnull(V.[15],0) as _15,isnull(V.[16],0) as _16,isnull(V.[17],0) as _17,isnull(V.[18],0) as _18,isnull(V.[19],0) as _19,isnull(V.[20],0) as _20,isnull(V.[21],0) as _21,isnull(V.[22],0) as _22,isnull(V.[23],0) as _23,isnull(V.[24],0) as _24,isnull(V.[25],0) as _25 FROM(Select * from(SELECT Temporada, Semana, SUM(Convertidas) AS Convertidas FROM SEasonsun1213..UV_ProdRecepcion where CodEstatus<> 'C' and temporada = (select temporada from CatSemanas where GETDATE() between Inicio and Fin) and " +
               "DescProducto = 'ARANDANO' GROUP BY Temporada, Semana Union All SELECT Temporada, Semana, SUM(Convertidas) AS Convertidas FROM SeasonPlan..UV_ProdRecepcion where CodEstatus <> 'C' and temporada = (select temporada from CatSemanas where GETDATE() between Inicio and Fin) and " +
               "DescProducto = 'ARANDANO' GROUP BY Temporada, Semana)V PIVOT(SUM(Convertidas) FOR Semana in ([35],[36],[37],[38],[39],[40],[41],[42],[43],[44],[45],[46],[47],[48],[49],[50],[51],[52],[1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],[17],[18],[19],[20],[21],[22],[23],[24],[25]))X)V)A left join(Select V.Inicio, V.Temporada, V.Sem, isnull(V.[35], 0) as _35, isnull(V.[36], 0) as _36, isnull(V.[37], 0) as _37, isnull(V.[38], 0) as _38, isnull(V.[39], 0) as _39, isnull(V.[40], 0) as _40, isnull(V.[41], 0) as _41, isnull(V.[42], 0) as _42, isnull(V.[43], 0) as _43, isnull(V.[44], 0) as _44, isnull(V.[45], 0) as _45, isnull(V.[46], 0) as _46, isnull(V.[47], 0) as _47, isnull(V.[48], 0) as _48, isnull(V.[49], 0) as _49, isnull(V.[50], 0) as _50, isnull(V.[51], 0) as _51, isnull(V.[52], 0) as _52, isnull(V.[1], 0) as _1, isnull(V.[2], 0) as _2, isnull(V.[3], 0) as _3, isnull(V.[4], 0) as _4, isnull(V.[5], 0) as _5, isnull(V.[6], 0) as _6, isnull(V.[7], 0) as _7, isnull(V.[8], 0) as _8, isnull(V.[9], 0) as _9, isnull(V.[10], 0) as _10, isnull(V.[11], 0) as _11, isnull(V.[12], 0) as _12, isnull(V.[13], 0) as _13, isnull(V.[14], 0) as _14, isnull(V.[15], 0) as _15, isnull(V.[16], 0) as _16, isnull(V.[17], 0) as _17, isnull(V.[18], 0) as _18, isnull(V.[19], 0) as _19, isnull(V.[20], 0) as _20, isnull(V.[21], 0) as _21, isnull(V.[22], 0) as _22, isnull(V.[23], 0) as _23, isnull(V.[24], 0) as _24, isnull(V.[25], 0) as _25 from(select * from(SELECT S.Semana AS Sem, S.Inicio, E.Semanas, E.Cantidad, E.Temporada FROM Estimacion_Berries E left join CatSemanas S on E.Temporada = S.Temporada AND E.Semana = S.Semana where E.Temporada = (select temporada from CatSemanas where GETDATE() between Inicio and Fin) and " +
               "E.Concepto = 'PROYECCION SEMANAL' and E.Cultivo = 3)V PIVOT(SUM(Cantidad) For Semanas in ([35],[36],[37],[38],[39],[40],[41],[42],[43],[44],[45],[46],[47],[48],[49],[50],[51],[52],[1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],[17],[18],[19],[20],[21],[22],[23],[24],[25]))X)V)V on A.Temporada = V.Temporada)V)V ORDER BY V.Inicio").ToList();

            int xPSA = 51;
            foreach (var item in pSA)
            {
                wsA.Cells[string.Format("A{0}", xPSA)].Value = item.SEMANA;
                wsA.Cells[string.Format("B{0}", xPSA)].Value = item._35;
                wsA.Cells[string.Format("C{0}", xPSA)].Value = item._36;
                wsA.Cells[string.Format("D{0}", xPSA)].Value = item._37;
                wsA.Cells[string.Format("E{0}", xPSA)].Value = item._38;
                wsA.Cells[string.Format("F{0}", xPSA)].Value = item._39;
                wsA.Cells[string.Format("G{0}", xPSA)].Value = item._40;
                wsA.Cells[string.Format("H{0}", xPSA)].Value = item._41;
                wsA.Cells[string.Format("I{0}", xPSA)].Value = item._42;
                wsA.Cells[string.Format("J{0}", xPSA)].Value = item._43;
                wsA.Cells[string.Format("K{0}", xPSA)].Value = item._44;
                wsA.Cells[string.Format("L{0}", xPSA)].Value = item._45;
                wsA.Cells[string.Format("M{0}", xPSA)].Value = item._46;
                wsA.Cells[string.Format("N{0}", xPSA)].Value = item._47;
                wsA.Cells[string.Format("O{0}", xPSA)].Value = item._48;
                wsA.Cells[string.Format("P{0}", xPSA)].Value = item._49;
                wsA.Cells[string.Format("Q{0}", xPSA)].Value = item._50;
                wsA.Cells[string.Format("R{0}", xPSA)].Value = item._51;
                wsA.Cells[string.Format("S{0}", xPSA)].Value = item._52;
                wsA.Cells[string.Format("T{0}", xPSA)].Value = item._1;
                wsA.Cells[string.Format("U{0}", xPSA)].Value = item._2;
                wsA.Cells[string.Format("V{0}", xPSA)].Value = item._3;
                wsA.Cells[string.Format("W{0}", xPSA)].Value = item._4;
                wsA.Cells[string.Format("X{0}", xPSA)].Value = item._5;
                wsA.Cells[string.Format("Y{0}", xPSA)].Value = item._6;
                wsA.Cells[string.Format("Z{0}", xPSA)].Value = item._7;
                wsA.Cells[string.Format("AA{0}", xPSA)].Value = item._8;
                wsA.Cells[string.Format("AB{0}", xPSA)].Value = item._9;
                wsA.Cells[string.Format("AC{0}", xPSA)].Value = item._10;
                wsA.Cells[string.Format("AD{0}", xPSA)].Value = item._11;
                wsA.Cells[string.Format("AE{0}", xPSA)].Value = item._12;
                wsA.Cells[string.Format("AF{0}", xPSA)].Value = item._13;
                wsA.Cells[string.Format("AG{0}", xPSA)].Value = item._14;
                wsA.Cells[string.Format("AH{0}", xPSA)].Value = item._15;
                wsA.Cells[string.Format("AI{0}", xPSA)].Value = item._16;
                wsA.Cells[string.Format("AJ{0}", xPSA)].Value = item._17;
                wsA.Cells[string.Format("AK{0}", xPSA)].Value = item._18;
                wsA.Cells[string.Format("AL{0}", xPSA)].Value = item._19;
                wsA.Cells[string.Format("AM{0}", xPSA)].Value = item._20;
                wsA.Cells[string.Format("AN{0}", xPSA)].Value = item._21;
                wsA.Cells[string.Format("AO{0}", xPSA)].Value = item._22;
                wsA.Cells[string.Format("AP{0}", xPSA)].Value = item._23;
                wsA.Cells[string.Format("AO{0}", xPSA)].Value = item._24;
                wsA.Cells[string.Format("AP{0}", xPSA)].Value = item._25;
                xPSA++;
            }

            wsA.Cells["A:AR"].AutoFitColumns();
            wsA.Row(2).Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            wsA.Row(2).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(string.Format("yellow")));
            wsA.Row(6).Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            wsA.Row(6).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(string.Format("yellow")));
            wsA.Row(50).Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            wsA.Row(50).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(string.Format("yellow")));


            //-------------------------------------------FRESA-----------------------------------------------------
            ExcelWorksheet wsFR = excel.Workbook.Worksheets.Add("FRESA");
            wsFR.Cells["A1"].Value = "VARIACION SEMANAL DE LA ESTIMACION DE FRESA";
            wsFR.Cells["A2"].Value = "SEMANA";
            wsFR.Cells["B2"].Value = "37";
            wsFR.Cells["C2"].Value = "38";
            wsFR.Cells["D2"].Value = "39";
            wsFR.Cells["E2"].Value = "40";
            wsFR.Cells["F2"].Value = "41";
            wsFR.Cells["G2"].Value = "42";
            wsFR.Cells["H2"].Value = "43";
            wsFR.Cells["I2"].Value = "44";
            wsFR.Cells["J2"].Value = "45";
            wsFR.Cells["K2"].Value = "46";
            wsFR.Cells["L2"].Value = "47";
            wsFR.Cells["M2"].Value = "48";
            wsFR.Cells["N2"].Value = "49";
            wsFR.Cells["O2"].Value = "50";
            wsFR.Cells["P2"].Value = "51";
            wsFR.Cells["Q2"].Value = "52";
            wsFR.Cells["R2"].Value = "1";
            wsFR.Cells["S2"].Value = "2";
            wsFR.Cells["T2"].Value = "3";
            wsFR.Cells["U2"].Value = "4";
            wsFR.Cells["V2"].Value = "5";
            wsFR.Cells["W2"].Value = "6";
            wsFR.Cells["X2"].Value = "7";
            wsFR.Cells["Y2"].Value = "8";
            wsFR.Cells["Z2"].Value = "9";
            wsFR.Cells["AA2"].Value = "10";
            wsFR.Cells["AB2"].Value = "11";
            wsFR.Cells["AC2"].Value = "12";
            wsFR.Cells["AD2"].Value = "13";
            wsFR.Cells["AE2"].Value = "14";
            wsFR.Cells["AF2"].Value = "15";
            wsFR.Cells["AG2"].Value = "16";
            wsFR.Cells["AH2"].Value = "17";
            wsFR.Cells["AI2"].Value = "18";
            wsFR.Cells["AJ2"].Value = "19";
            wsFR.Cells["AK2"].Value = "20";
            wsFR.Cells["AL2"].Value = "21";
            wsFR.Cells["AM2"].Value = "22";
            wsFR.Cells["AN2"].Value = "23";
            wsFR.Cells["AO2"].Value = "24";
            wsFR.Cells["AP2"].Value = "25";

            var recepcion_realFR = bd.Database.SqlQuery<R_Berries>("SELECT 'TOTAL' as 'SEMANA',  round(isnull(V.[37],0),0) as _37,round(isnull(V.[38],0),0) as _38,round(isnull(V.[39],0),0) as _39,round(isnull(V.[40],0),0) as _40,round(isnull(V.[41],0),0) as _41,round(isnull(V.[42],0),0) as _42,round(isnull(V.[43],0),0) as _43,round(isnull(V.[44],0),0) as _44,round(isnull(V.[45],0),0) as _45,round(isnull(V.[46],0),0) as _46,round(isnull(V.[47],0),0) as _47,round(isnull(V.[48],0),0) as _48,round(isnull(V.[49],0),0) as _49,round(isnull(V.[50],0),0) as _50,round(isnull(V.[51],0),0) as _51,round(isnull(V.[52],0),0) as _52,round(isnull(V.[1],0),0) as _1,round(isnull(V.[2],0),0) as _2,round(isnull(V.[3],0),0) as _3,round(isnull(V.[4],0),0) as _4,round(isnull(V.[5],0),0) as _5,round(isnull(V.[6],0),0) as _6,round(isnull(V.[7],0),0) as _7,round(isnull(V.[8],0),0) as _8,round(isnull(V.[9],0),0) as _9,round(isnull(V.[10],0),0) as _10,round(isnull(V.[11],0),0) as _11,round(isnull(V.[12],0),0) as _12,round(isnull(V.[13],0),0) as _13,round(isnull(V.[14],0),0) as _14,round(isnull(V.[15],0),0) as _15,round(isnull(V.[16],0),0) as _16,round(isnull(V.[17],0),0) as _17,round(isnull(V.[18],0),0) as _18,round(isnull(V.[19],0),0) as _19,round(isnull(V.[20],0),0) as _20,round(isnull(V.[21],0),0) as _21,round(isnull(V.[22],0),0) as _22,round(isnull(V.[23],0),0) as _23,round(isnull(V.[24],0),0) as _24,round(isnull(V.[25],0),0) as _25 " +
                "FROM(Select * from(SELECT Semana, SUM(Convertidas) AS Convertidas FROM SEasonsun1213..UV_ProdRecepcion where CodEstatus <> 'C' and temporada = (select temporada from CatSemanas where GETDATE() between Inicio and Fin) and DescProducto = 'FRESA' GROUP BY Semana Union All " +
                "SELECT Semana, SUM(Convertidas) AS Convertidas FROM SeasonPlan..UV_ProdRecepcion where CodEstatus <> 'C' and temporada = (select temporada from CatSemanas where GETDATE() between Inicio and Fin) and DescProducto = 'FRESA' GROUP BY Semana" +
                ")V PIVOT(SUM(Convertidas) FOR Semana in ([37],[38],[39],[40],[41],[42],[43],[44],[45],[46],[47],[48],[49],[50],[51],[52],[1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],[17],[18],[19],[20],[21],[22],[23],[24],[25]))X)V ").ToList();

            int xRFR = 3;
            foreach (var item in recepcion_realFR)
            {
                wsFR.Cells[string.Format("A{0}", xRFR)].Value = "RECEPCION REAL";
                wsFR.Cells[string.Format("B{0}", xRFR)].Value = item._37;
                wsFR.Cells[string.Format("C{0}", xRFR)].Value = item._38;
                wsFR.Cells[string.Format("D{0}", xRFR)].Value = item._39;
                wsFR.Cells[string.Format("E{0}", xRFR)].Value = item._40;
                wsFR.Cells[string.Format("F{0}", xRFR)].Value = item._41;
                wsFR.Cells[string.Format("G{0}", xRFR)].Value = item._42;
                wsFR.Cells[string.Format("H{0}", xRFR)].Value = item._43;
                wsFR.Cells[string.Format("I{0}", xRFR)].Value = item._44;
                wsFR.Cells[string.Format("J{0}", xRFR)].Value = item._45;
                wsFR.Cells[string.Format("K{0}", xRFR)].Value = item._46;
                wsFR.Cells[string.Format("L{0}", xRFR)].Value = item._47;
                wsFR.Cells[string.Format("M{0}", xRFR)].Value = item._48;
                wsFR.Cells[string.Format("N{0}", xRFR)].Value = item._49;
                wsFR.Cells[string.Format("O{0}", xRFR)].Value = item._50;
                wsFR.Cells[string.Format("P{0}", xRFR)].Value = item._51;
                wsFR.Cells[string.Format("Q{0}", xRFR)].Value = item._52;
                wsFR.Cells[string.Format("R{0}", xRFR)].Value = item._1;
                wsFR.Cells[string.Format("S{0}", xRFR)].Value = item._2;
                wsFR.Cells[string.Format("T{0}", xRFR)].Value = item._3;
                wsFR.Cells[string.Format("U{0}", xRFR)].Value = item._4;
                wsFR.Cells[string.Format("V{0}", xRFR)].Value = item._5;
                wsFR.Cells[string.Format("W{0}", xRFR)].Value = item._6;
                wsFR.Cells[string.Format("X{0}", xRFR)].Value = item._7;
                wsFR.Cells[string.Format("Y{0}", xRFR)].Value = item._8;
                wsFR.Cells[string.Format("Z{0}", xRFR)].Value = item._9;
                wsFR.Cells[string.Format("AA{0}", xRFR)].Value = item._10;
                wsFR.Cells[string.Format("AB{0}", xRFR)].Value = item._11;
                wsFR.Cells[string.Format("AC{0}", xRFR)].Value = item._12;
                wsFR.Cells[string.Format("AD{0}", xRFR)].Value = item._13;
                wsFR.Cells[string.Format("AE{0}", xRFR)].Value = item._14;
                wsFR.Cells[string.Format("AF{0}", xRFR)].Value = item._15;
                wsFR.Cells[string.Format("AG{0}", xRFR)].Value = item._16;
                wsFR.Cells[string.Format("AH{0}", xRFR)].Value = item._17;
                wsFR.Cells[string.Format("AI{0}", xRFR)].Value = item._18;
                wsFR.Cells[string.Format("AJ{0}", xRFR)].Value = item._19;
                wsFR.Cells[string.Format("AK{0}", xRFR)].Value = item._20;
                wsFR.Cells[string.Format("AL{0}", xRFR)].Value = item._21;
                wsFR.Cells[string.Format("AM{0}", xRFR)].Value = item._22;
                wsFR.Cells[string.Format("AN{0}", xRFR)].Value = item._23;
                wsFR.Cells[string.Format("AO{0}", xRFR)].Value = item._24;
                wsFR.Cells[string.Format("AP{0}", xRFR)].Value = item._25;
                xRFR++;
            }

            //PROYECCION SEMANAL
            wsFR.Cells["A5"].Value = "PROYECCION SEMANAL";
            wsFR.Cells["B5"].Value = "37";
            wsFR.Cells["C5"].Value = "38";
            wsFR.Cells["D5"].Value = "39";
            wsFR.Cells["E5"].Value = "40";
            wsFR.Cells["F5"].Value = "41";
            wsFR.Cells["G5"].Value = "42";
            wsFR.Cells["H5"].Value = "43";
            wsFR.Cells["I5"].Value = "44";
            wsFR.Cells["J5"].Value = "45";
            wsFR.Cells["K5"].Value = "46";
            wsFR.Cells["L5"].Value = "47";
            wsFR.Cells["M5"].Value = "48";
            wsFR.Cells["N5"].Value = "49";
            wsFR.Cells["O5"].Value = "50";
            wsFR.Cells["P5"].Value = "51";
            wsFR.Cells["Q5"].Value = "52";
            wsFR.Cells["R5"].Value = "1";
            wsFR.Cells["S5"].Value = "2";
            wsFR.Cells["T5"].Value = "3";
            wsFR.Cells["U5"].Value = "4";
            wsFR.Cells["V5"].Value = "5";
            wsFR.Cells["W5"].Value = "6";
            wsFR.Cells["X5"].Value = "7";
            wsFR.Cells["Y5"].Value = "8";
            wsFR.Cells["Z5"].Value = "9";
            wsFR.Cells["AA5"].Value = "10";
            wsFR.Cells["AB5"].Value = "11";
            wsFR.Cells["AC5"].Value = "12";
            wsFR.Cells["AD5"].Value = "13";
            wsFR.Cells["AE5"].Value = "14";
            wsFR.Cells["AF5"].Value = "15";
            wsFR.Cells["AG5"].Value = "16";
            wsFR.Cells["AH5"].Value = "17";
            wsFR.Cells["AI5"].Value = "18";
            wsFR.Cells["AJ5"].Value = "19";
            wsFR.Cells["AK5"].Value = "20";
            wsFR.Cells["AL5"].Value = "21";
            wsFR.Cells["AM5"].Value = "22";
            wsFR.Cells["AN5"].Value = "23";
            wsFR.Cells["AO5"].Value = "24";
            wsFR.Cells["AP5"].Value = "25";

            var dataSFR = bd.Database.SqlQuery<E_Berries>("Select V.Sem AS 'SEMANA', round(isnull(V.[37],0),0) as _37,round(isnull(V.[38],0),0) as _38,round(isnull(V.[39],0),0) as _39,round(isnull(V.[40],0),0) as _40,round(isnull(V.[41],0),0) as _41,round(isnull(V.[42],0),0) as _42,round(isnull(V.[43],0),0) as _43,round(isnull(V.[44],0),0) as _44,round(isnull(V.[45],0),0) as _45,round(isnull(V.[46],0),0) as _46,round(isnull(V.[47],0),0) as _47,round(isnull(V.[48],0),0) as _48,round(isnull(V.[49],0),0) as _49,round(isnull(V.[50],0),0) as _50,round(isnull(V.[51],0),0) as _51,round(isnull(V.[52],0),0) as _52,round(isnull(V.[1],0),0) as _1,round(isnull(V.[2],0),0) as _2,round(isnull(V.[3],0),0) as _3,round(isnull(V.[4],0),0) as _4,round(isnull(V.[5],0),0) as _5,round(isnull(V.[6],0),0) as _6,round(isnull(V.[7],0),0) as _7,round(isnull(V.[8],0),0) as _8,round(isnull(V.[9],0),0) as _9,round(isnull(V.[10],0),0) as _10,round(isnull(V.[11],0),0) as _11,round(isnull(V.[12],0),0) as _12,round(isnull(V.[13],0),0) as _13,round(isnull(V.[14],0),0) as _14,round(isnull(V.[15],0),0) as _15,round(isnull(V.[16],0),0) as _16,round(isnull(V.[17],0),0) as _17,round(isnull(V.[18],0),0) as _18,round(isnull(V.[19],0),0) as _19,round(isnull(V.[20],0),0) as _20,round(isnull(V.[21],0),0) as _21,round(isnull(V.[22],0),0) as _22,round(isnull(V.[23],0),0) as _23,round(isnull(V.[24],0),0) as _24,round(isnull(V.[25],0),0) as _25 " +
                "from(Select * from(SELECT S.Semana AS Sem, S.Inicio, E.Semanas, E.Cantidad FROM Estimacion_Berries E left join CatSemanas S on E.Temporada = S.Temporada AND E.Semana = S.Semana " +
                "where E.Temporada = (select temporada from CatSemanas where GETDATE() between Inicio and Fin) and E.Concepto = 'PROYECCION SEMANAL' and E.Cultivo = 4)V " +
                "PIVOT(SUM(Cantidad) For Semanas in ([37],[38],[39],[40],[41],[42],[43],[44],[45],[46],[47],[48],[49],[50],[51],[52],[1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],[17],[18],[19],[20],[21],[22],[23],[24],[25]))X)V ORDER BY V.Inicio").ToList();

            int xSFR = 6;
            foreach (var item in dataSFR)
            {
                wsFR.Cells[string.Format("A{0}", xSFR)].Value = item.SEMANA;
                wsFR.Cells[string.Format("B{0}", xSFR)].Value = item._37;
                wsFR.Cells[string.Format("C{0}", xSFR)].Value = item._38;
                wsFR.Cells[string.Format("D{0}", xSFR)].Value = item._39;
                wsFR.Cells[string.Format("E{0}", xSFR)].Value = item._40;
                wsFR.Cells[string.Format("F{0}", xSFR)].Value = item._41;
                wsFR.Cells[string.Format("G{0}", xSFR)].Value = item._42;
                wsFR.Cells[string.Format("H{0}", xSFR)].Value = item._43;
                wsFR.Cells[string.Format("I{0}", xSFR)].Value = item._44;
                wsFR.Cells[string.Format("J{0}", xSFR)].Value = item._45;
                wsFR.Cells[string.Format("K{0}", xSFR)].Value = item._46;
                wsFR.Cells[string.Format("L{0}", xSFR)].Value = item._47;
                wsFR.Cells[string.Format("M{0}", xSFR)].Value = item._48;
                wsFR.Cells[string.Format("N{0}", xSFR)].Value = item._49;
                wsFR.Cells[string.Format("O{0}", xSFR)].Value = item._50;
                wsFR.Cells[string.Format("P{0}", xSFR)].Value = item._51;
                wsFR.Cells[string.Format("Q{0}", xSFR)].Value = item._52;
                wsFR.Cells[string.Format("R{0}", xSFR)].Value = item._1;
                wsFR.Cells[string.Format("S{0}", xSFR)].Value = item._2;
                wsFR.Cells[string.Format("T{0}", xSFR)].Value = item._3;
                wsFR.Cells[string.Format("U{0}", xSFR)].Value = item._4;
                wsFR.Cells[string.Format("V{0}", xSFR)].Value = item._5;
                wsFR.Cells[string.Format("W{0}", xSFR)].Value = item._6;
                wsFR.Cells[string.Format("X{0}", xSFR)].Value = item._7;
                wsFR.Cells[string.Format("Y{0}", xSFR)].Value = item._8;
                wsFR.Cells[string.Format("Z{0}", xSFR)].Value = item._9;
                wsFR.Cells[string.Format("AA{0}", xSFR)].Value = item._10;
                wsFR.Cells[string.Format("AB{0}", xSFR)].Value = item._11;
                wsFR.Cells[string.Format("AC{0}", xSFR)].Value = item._12;
                wsFR.Cells[string.Format("AD{0}", xSFR)].Value = item._13;
                wsFR.Cells[string.Format("AE{0}", xSFR)].Value = item._14;
                wsFR.Cells[string.Format("AF{0}", xSFR)].Value = item._15;
                wsFR.Cells[string.Format("AG{0}", xSFR)].Value = item._16;
                wsFR.Cells[string.Format("AH{0}", xSFR)].Value = item._17;
                wsFR.Cells[string.Format("AI{0}", xSFR)].Value = item._18;
                wsFR.Cells[string.Format("AJ{0}", xSFR)].Value = item._19;
                wsFR.Cells[string.Format("AK{0}", xSFR)].Value = item._20;
                wsFR.Cells[string.Format("AL{0}", xSFR)].Value = item._21;
                wsFR.Cells[string.Format("AM{0}", xSFR)].Value = item._22;
                wsFR.Cells[string.Format("AN{0}", xSFR)].Value = item._23;
                wsFR.Cells[string.Format("AO{0}", xSFR)].Value = item._24;
                wsFR.Cells[string.Format("AP{0}", xSFR)].Value = item._25;
                xSFR++;
            }

            //PORCENTAJES
            wsFR.Cells["A49"].Value = "PORCENTAJES";
            wsFR.Cells["B49"].Value = "37";
            wsFR.Cells["C49"].Value = "38";
            wsFR.Cells["D49"].Value = "39";
            wsFR.Cells["E49"].Value = "40";
            wsFR.Cells["F49"].Value = "41";
            wsFR.Cells["G49"].Value = "42";
            wsFR.Cells["H49"].Value = "43";
            wsFR.Cells["I49"].Value = "44";
            wsFR.Cells["J49"].Value = "45";
            wsFR.Cells["K49"].Value = "46";
            wsFR.Cells["L49"].Value = "47";
            wsFR.Cells["M49"].Value = "48";
            wsFR.Cells["N49"].Value = "49";
            wsFR.Cells["O49"].Value = "50";
            wsFR.Cells["P49"].Value = "51";
            wsFR.Cells["Q49"].Value = "52";
            wsFR.Cells["R49"].Value = "1";
            wsFR.Cells["S49"].Value = "2";
            wsFR.Cells["T49"].Value = "3";
            wsFR.Cells["U49"].Value = "4";
            wsFR.Cells["V49"].Value = "5";
            wsFR.Cells["W49"].Value = "6";
            wsFR.Cells["X49"].Value = "7";
            wsFR.Cells["Y49"].Value = "8";
            wsFR.Cells["Z49"].Value = "9";
            wsFR.Cells["AA49"].Value = "10";
            wsFR.Cells["AB49"].Value = "11";
            wsFR.Cells["AC49"].Value = "12";
            wsFR.Cells["AD49"].Value = "13";
            wsFR.Cells["AE49"].Value = "14";
            wsFR.Cells["AF49"].Value = "15";
            wsFR.Cells["AG49"].Value = "16";
            wsFR.Cells["AH49"].Value = "17";
            wsFR.Cells["AI49"].Value = "18";
            wsFR.Cells["AJ49"].Value = "19";
            wsFR.Cells["AK49"].Value = "20";
            wsFR.Cells["AL49"].Value = "21";
            wsFR.Cells["AM49"].Value = "22";
            wsFR.Cells["AN49"].Value = "23";
            wsFR.Cells["AO49"].Value = "24";
            wsFR.Cells["AP49"].Value = "25";

            var pSFR = bd.Database.SqlQuery<E_BerriesP>("select V.Sem AS 'SEMANA', (case when V._37='0%' then '' else V._37 end) as _37,(case when V._38='0%' then '' else V._38 end) as _38,(case when V._39='0%' then '' else V._39 end) as _39,(case when V._40='0%' then '' else V._40 end) as _40,(case when V._41='0%' then '' else V._41 end) as _41,(case when V._42='0%' then '' else V._42 end) as _42,(case when V._43='0%' then '' else V._43 end) as _43,(case when V._44='0%' then '' else V._44 end) as _44,(case when V._45='0%' then '' else V._45 end) as _45,(case when V._46='0%' then '' else V._46 end) as _46,(case when V._47='0%' then '' else V._47 end) as _47,(case when V._48='0%' then '' else V._48 end) as _48,(case when V._49='0%' then '' else V._49 end) as _49,(case when V._50='0%' then '' else V._50 end) as _50,(case when V._51='0%' then '' else V._51 end) as _51,(case when V._52='0%' then '' else V._52 end) as _52,(case when V._1='0%' then '' else V._1 end) as _1,(case when V._2='0%' then '' else V._2 end) as _2,(case when V._3='0%' then '' else V._3 end) as _3,(case when V._4='0%' then '' else V._4 end) as _4,(case when V._5='0%' then '' else V._5 end) as _5,(case when V._6='0%' then '' else V._6 end) as _6,(case when V._7='0%' then '' else V._7 end) as _7,(case when V._8='0%' then '' else V._8 end) as _8,(case when V._9='0%' then '' else V._9 end) as _9,(case when V._10='0%' then '' else V._10 end) as _10,(case when V._11='0%' then '' else V._11 end) as _11,(case when V._12='0%' then '' else V._12 end) as _12,(case when V._13='0%' then '' else V._13 end) as _13,(case when V._14='0%' then '' else V._14 end) as _14,(case when V._15='0%' then '' else V._15 end) as _15,(case when V._16='0%' then '' else V._16 end) as _16,(case when V._17='0%' then '' else V._17 end) as _17,(case when V._18='0%' then '' else V._18 end) as _18,(case when V._19='0%' then '' else V._19 end) as _19,(case when V._20='0%' then '' else V._20 end) as _20,(case when V._21='0%' then '' else V._21 end) as _21,(case when V._22='0%' then '' else V._22 end) as _22,(case when V._23='0%' then '' else V._23 end) as _23,(case when V._24='0%' then '' else V._24 end) as _24,(case when V._25='0%' then '' else V._25 end) as _25 from(select V.Inicio, V.Sem, cast(V._35 as varchar)+'%' as _35,cast(V._36 as varchar)+'%' as _36,cast(V._37 as varchar)+'%' as _37,cast(V._38 as varchar)+'%' as _38,cast(V._39 as varchar)+'%' as _39,cast(V._40 as varchar)+'%' as _40,cast(V._41 as varchar)+'%' as _41,cast(V._42 as varchar)+'%' as _42,cast(V._43 as varchar)+'%' as _43,cast(V._44 as varchar)+'%' as _44,cast(V._45 as varchar)+'%' as _45,cast(V._46 as varchar)+'%' as _46,cast(V._47 as varchar)+'%' as _47,cast(V._48 as varchar)+'%' as _48,cast(V._49 as varchar)+'%' as _49,cast(V._50 as varchar)+'%' as _50,cast(V._51 as varchar)+'%' as _51,cast(V._52 as varchar)+'%' as _52,cast(V._1 as varchar)+'%' as _1,cast(V._2 as varchar)+'%' _2,cast(V._3 as varchar)+'%' as _3,cast(V._4 as varchar)+'%' as _4,cast(V._5 as varchar)+'%' as _5,cast(V._6 as varchar)+'%' as _6,cast(V._7 as varchar)+'%' as _7,cast(V._8 as varchar)+'%' as _8,cast(V._9 as varchar)+'%' as _9,cast(V._10 as varchar)+'%' as _10,cast(V._11 as varchar)+'%' as _11,cast(V._12 as varchar)+'%' as _12,cast(V._13 as varchar)+'%' as _13,cast(V._14 as varchar)+'%' as _14,cast(V._15 as varchar)+'%' as _15,cast(V._16 as varchar)+'%' as _16,cast(V._17 as varchar)+'%' as _17,cast(V._18 as varchar)+'%' as _18,cast(V._19 as varchar)+'%' as _19,cast(V._20 as varchar)+'%'as _20,cast(V._21 as varchar)+'%' as _21,cast(V._22 as varchar)+'%' as _22,cast(V._23 as varchar)+'%' as _23,cast(V._24 as varchar)+'%' as _24,cast(V._25 as varchar)+'%' as _25 from (SELECT V.Inicio, V.Sem,(case when A._35=0 then '' else case when V._35=0 then '' else round((A._35/V._35)*100,0) end end) as _35,(case when A._36=0 then 0 else case when V._36=0 then 0 else round((A._36/V._36)*100,0) end end) as _36,(case when A._37=0 then 0 else case when V._37=0 then 0 else round((A._37/V._37)*100,0) end end) as _37,(case when A._38=0 then 0 else case when V._38=0 then 0 else round((A._38/V._38)*100,0) end end) as _38,(case when A._39=0 then 0 else case when V._39=0 then 0 else round((A._39/V._39)*100,0) end end) as _39,(case when A._40=0 then 0 else case when V._40=0 then 0 else round((A._40/V._40)*100,0) end end) as _40,(case when A._41=0 then 0 else case when V._41=0 then 0 else round((A._41/V._41)*100,0) end end) as _41,(case when A._42=0 then 0 else case when V._42=0 then 0 else round((A._42/V._42)*100,0) end end) as _42,(case when A._43=0 then 0 else case when V._43=0 then 0 else round((A._43/V._43)*100,0) end end) as _43,(case when A._44=0 then 0 else case when V._44=0 then 0 else round((A._44/V._44)*100,0) end end) as _44,(case when A._45=0 then 0 else case when V._45=0 then 0 else round((A._45/V._45)*100,0) end end) as _45,(case when A._46=0 then 0 else case when V._46=0 then 0 else round((A._46/V._46)*100,0) end end) as _46,(case when A._47=0 then 0 else case when V._47=0 then 0 else round((A._47/V._47)*100,0) end end) as _47,(case when A._48=0 then 0 else case when V._48=0 then 0 else round((A._48/V._48)*100,0) end end) as _48,(case when A._49=0 then 0 else case when V._49=0 then 0 else round((A._49/V._49)*100,0) end end) as _49,(case when A._50=0 then 0 else case when V._50=0 then 0 else round((A._50/V._50)*100,0) end end) as _50,(case when A._51=0 then 0 else case when V._51=0 then 0 else round((A._51/V._51)*100,0) end end) as _51,(case when A._52=0 then 0 else case when V._52=0 then 0 else round((A._52/V._52)*100,0) end end) as _52,(case when A._1=0 then 0 else case when V._1=0 then 0 else round((A._1/V._1)*100,0) end end) as _1,(case when A._2=0 then 0 else case when V._2=0 then 0 else round((A._2/V._2)*100,0) end end) as _2,(case when A._3=0 then 0 else case when V._3=0 then 0 else round((A._3/V._3)*100,0) end end) as _3,(case when A._4=0 then 0 else case when V._4=0 then 0 else round((A._4/V._4)*100,0) end end) as _4,(case when A._5=0 then 0 else case when V._5=0 then 0 else round((A._5/V._5)*100,0) end end) as _5,(case when A._6=0 then 0 else case when V._6=0 then 0 else round((A._6/V._6)*100,0) end end) as _6,(case when A._7=0 then 0 else case when V._7=0 then 0 else round((A._7/V._7)*100,0) end end) as _7,(case when A._8=0 then 0 else case when V._8=0 then 0 else round((A._8/V._8)*100,0) end end) as _8,(case when A._9=0 then 0 else case when V._9=0 then 0 else round((A._9/V._9)*100,0) end end) as _9,(case when A._10=0 then 0 else case when V._10=0 then 0 else round((A._10/V._10)*100,0) end end) as _10,(case when A._11=0 then 0 else case when V._11=0 then 0 else round((A._11/V._11)*100,0) end end) as _11,(case when A._12=0 then 0 else case when V._12=0 then 0 else round((A._12/V._12)*100,0) end end) as _12,(case when A._13=0 then 0 else case when V._13=0 then 0 else round((A._13/V._13)*100,0) end end) as _13,(case when A._14=0 then 0 else case when V._14=0 then 0 else round((A._14/V._14)*100,0) end end) as _14,(case when A._15=0 then 0 else case when V._15=0 then 0 else round((A._15/V._15)*100,0) end end) as _15,(case when A._16=0 then 0 else case when V._16=0 then 0 else round((A._16/V._16)*100,0) end end) as _16,(case when A._17=0 then 0 else case when V._17=0 then 0 else round((A._17/V._17)*100,0) end end) as _17,(case when A._18=0 then 0 else case when V._18=0 then 0 else round((A._18/V._18)*100,0) end end) as _18,(case when A._19=0 then 0 else case when V._19=0 then 0 else round((A._19/V._19)*100,0) end end) as _19,(case when A._20=0 then 0 else case when V._20=0 then 0 else round((A._20/V._20)*100,0) end end) as _20,(case when A._21=0 then 0 else case when V._21=0 then 0 else round((A._21/V._21)*100,0) end end) as _21,(case when A._22=0 then 0 else case when V._22=0 then 0 else round((A._22/V._22)*100,0) end end) as _22,(case when A._23=0 then 0 else case when V._23=0 then 0 else round((A._23/V._23)*100,0) end end) as _23,(case when A._24=0 then 0 else case when V._24=0 then 0 else round((A._24/V._24)*100,0) end end) as _24,(case when A._25=0 then 0 else case when V._25=0 then 0 else round((A._25/V._25)*100,0) end end) as _25 FROM(SELECT V.Temporada, isnull(V.[35],0) as _35,isnull(V.[36],0) as _36,isnull(V.[37],0) as _37,isnull(V.[38],0) as _38,isnull(V.[39],0) as _39,isnull(V.[40],0) as _40,isnull(V.[41],0) as _41,isnull(V.[42],0) as _42,isnull(V.[43],0) as _43,isnull(V.[44],0) as _44,isnull(V.[45],0) as _45,isnull(V.[46],0) as _46,isnull(V.[47],0) as _47,isnull(V.[48],0) as _48,isnull(V.[49],0) as _49,isnull(V.[50],0) as _50,isnull(V.[51],0) as _51,isnull(V.[52],0) as _52,isnull(V.[1],0) as _1,isnull(V.[2],0) as _2,isnull(V.[3],0) as _3,isnull(V.[4],0) as _4,isnull(V.[5],0) as _5,isnull(V.[6],0) as _6,isnull(V.[7],0) as _7,isnull(V.[8],0) as _8,isnull(V.[9],0) as _9,isnull(V.[10],0) as _10,isnull(V.[11],0) as _11,isnull(V.[12],0) as _12,isnull(V.[13],0) as _13,isnull(V.[14],0) as _14,isnull(V.[15],0) as _15,isnull(V.[16],0) as _16,isnull(V.[17],0) as _17,isnull(V.[18],0) as _18,isnull(V.[19],0) as _19,isnull(V.[20],0) as _20,isnull(V.[21],0) as _21,isnull(V.[22],0) as _22,isnull(V.[23],0) as _23,isnull(V.[24],0) as _24,isnull(V.[25],0) as _25 FROM(Select * from(SELECT Temporada, Semana, SUM(Convertidas) AS Convertidas FROM SEasonsun1213..UV_ProdRecepcion where CodEstatus<> 'C' and temporada = (select temporada from CatSemanas where GETDATE() between Inicio and Fin) and " +
                "DescProducto = 'FRESA' GROUP BY Temporada, Semana Union All SELECT Temporada, Semana, SUM(Convertidas) AS Convertidas FROM SeasonPlan..UV_ProdRecepcion where CodEstatus <> 'C' and temporada = (select temporada from CatSemanas where GETDATE() between Inicio and Fin) and " +
                "DescProducto = 'FRESA' GROUP BY Temporada, Semana)V PIVOT(SUM(Convertidas) FOR Semana in ([35],[36],[37],[38],[39],[40],[41],[42],[43],[44],[45],[46],[47],[48],[49],[50],[51],[52],[1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],[17],[18],[19],[20],[21],[22],[23],[24],[25]))X)V)A left join(Select V.Inicio, V.Temporada, V.Sem, isnull(V.[35], 0) as _35, isnull(V.[36], 0) as _36, isnull(V.[37], 0) as _37, isnull(V.[38], 0) as _38, isnull(V.[39], 0) as _39, isnull(V.[40], 0) as _40, isnull(V.[41], 0) as _41, isnull(V.[42], 0) as _42, isnull(V.[43], 0) as _43, isnull(V.[44], 0) as _44, isnull(V.[45], 0) as _45, isnull(V.[46], 0) as _46, isnull(V.[47], 0) as _47, isnull(V.[48], 0) as _48, isnull(V.[49], 0) as _49, isnull(V.[50], 0) as _50, isnull(V.[51], 0) as _51, isnull(V.[52], 0) as _52, isnull(V.[1], 0) as _1, isnull(V.[2], 0) as _2, isnull(V.[3], 0) as _3, isnull(V.[4], 0) as _4, isnull(V.[5], 0) as _5, isnull(V.[6], 0) as _6, isnull(V.[7], 0) as _7, isnull(V.[8], 0) as _8, isnull(V.[9], 0) as _9, isnull(V.[10], 0) as _10, isnull(V.[11], 0) as _11, isnull(V.[12], 0) as _12, isnull(V.[13], 0) as _13, isnull(V.[14], 0) as _14, isnull(V.[15], 0) as _15, isnull(V.[16], 0) as _16, isnull(V.[17], 0) as _17, isnull(V.[18], 0) as _18, isnull(V.[19], 0) as _19, isnull(V.[20], 0) as _20, isnull(V.[21], 0) as _21, isnull(V.[22], 0) as _22, isnull(V.[23], 0) as _23, isnull(V.[24], 0) as _24, isnull(V.[25], 0) as _25 from(select * from(SELECT S.Semana AS Sem, S.Inicio, E.Semanas, E.Cantidad, E.Temporada FROM Estimacion_Berries E left join CatSemanas S on E.Temporada = S.Temporada AND E.Semana = S.Semana where E.Temporada = (select temporada from CatSemanas where GETDATE() between Inicio and Fin) and " +
                "E.Concepto = 'PROYECCION SEMANAL' and E.Cultivo = 4)V PIVOT(SUM(Cantidad) For Semanas in ([35],[36],[37],[38],[39],[40],[41],[42],[43],[44],[45],[46],[47],[48],[49],[50],[51],[52],[1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],[17],[18],[19],[20],[21],[22],[23],[24],[25]))X)V)V on A.Temporada = V.Temporada)V)V ORDER BY V.Inicio").ToList();

            int xPSFR = 50;
            foreach (var item in pSFR)
            {
                wsFR.Cells[string.Format("A{0}", xPSFR)].Value = item.SEMANA;
                wsFR.Cells[string.Format("B{0}", xPSFR)].Value = item._37;
                wsFR.Cells[string.Format("C{0}", xPSFR)].Value = item._38;
                wsFR.Cells[string.Format("D{0}", xPSFR)].Value = item._39;
                wsFR.Cells[string.Format("E{0}", xPSFR)].Value = item._40;
                wsFR.Cells[string.Format("F{0}", xPSFR)].Value = item._41;
                wsFR.Cells[string.Format("G{0}", xPSFR)].Value = item._42;
                wsFR.Cells[string.Format("H{0}", xPSFR)].Value = item._43;
                wsFR.Cells[string.Format("I{0}", xPSFR)].Value = item._44;
                wsFR.Cells[string.Format("J{0}", xPSFR)].Value = item._45;
                wsFR.Cells[string.Format("K{0}", xPSFR)].Value = item._46;
                wsFR.Cells[string.Format("L{0}", xPSFR)].Value = item._47;
                wsFR.Cells[string.Format("M{0}", xPSFR)].Value = item._48;
                wsFR.Cells[string.Format("N{0}", xPSFR)].Value = item._49;
                wsFR.Cells[string.Format("O{0}", xPSFR)].Value = item._50;
                wsFR.Cells[string.Format("P{0}", xPSFR)].Value = item._51;
                wsFR.Cells[string.Format("Q{0}", xPSFR)].Value = item._52;
                wsFR.Cells[string.Format("R{0}", xPSFR)].Value = item._1;
                wsFR.Cells[string.Format("S{0}", xPSFR)].Value = item._2;
                wsFR.Cells[string.Format("T{0}", xPSFR)].Value = item._3;
                wsFR.Cells[string.Format("U{0}", xPSFR)].Value = item._4;
                wsFR.Cells[string.Format("V{0}", xPSFR)].Value = item._5;
                wsFR.Cells[string.Format("W{0}", xPSFR)].Value = item._6;
                wsFR.Cells[string.Format("X{0}", xPSFR)].Value = item._7;
                wsFR.Cells[string.Format("Y{0}", xPSFR)].Value = item._8;
                wsFR.Cells[string.Format("Z{0}", xPSFR)].Value = item._9;
                wsFR.Cells[string.Format("AA{0}", xPSFR)].Value = item._10;
                wsFR.Cells[string.Format("AB{0}", xPSFR)].Value = item._11;
                wsFR.Cells[string.Format("AC{0}", xPSFR)].Value = item._12;
                wsFR.Cells[string.Format("AD{0}", xPSFR)].Value = item._13;
                wsFR.Cells[string.Format("AE{0}", xPSFR)].Value = item._14;
                wsFR.Cells[string.Format("AF{0}", xPSFR)].Value = item._15;
                wsFR.Cells[string.Format("AG{0}", xPSFR)].Value = item._16;
                wsFR.Cells[string.Format("AH{0}", xPSFR)].Value = item._17;
                wsFR.Cells[string.Format("AI{0}", xPSFR)].Value = item._18;
                wsFR.Cells[string.Format("AJ{0}", xPSFR)].Value = item._19;
                wsFR.Cells[string.Format("AK{0}", xPSFR)].Value = item._20;
                wsFR.Cells[string.Format("AL{0}", xPSFR)].Value = item._21;
                wsFR.Cells[string.Format("AM{0}", xPSFR)].Value = item._22;
                wsFR.Cells[string.Format("AN{0}", xPSFR)].Value = item._23;
                wsFR.Cells[string.Format("AO{0}", xPSFR)].Value = item._24;
                wsFR.Cells[string.Format("AP{0}", xPSFR)].Value = item._25;
                xPSFR++;
            }

            wsFR.Cells["A:AP"].AutoFitColumns();

            wsFR.Row(2).Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            wsFR.Row(2).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(string.Format("yellow")));
            wsFR.Row(5).Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            wsFR.Row(5).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(string.Format("yellow")));
            wsFR.Row(49).Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            wsFR.Row(49).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(string.Format("yellow")));

            Response.Clear();
            Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            Response.AddHeader("content-disposition", "attachment: filename=" + "ExcelReport.xlsx");
            Response.BinaryWrite(excel.GetAsByteArray());
            Response.End();
        }
    }
}