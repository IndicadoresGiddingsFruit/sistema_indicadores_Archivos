using iTextSharp.text;
using iTextSharp.text.pdf;
using OfficeOpenXml;
using Sistema_Indicadores.Clases;
using Sistema_Indicadores.Models;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Threading.Tasks;
using System.Web.Mvc;

namespace Sistema_Indicadores.Controllers
{
    public class IndicadoresController : Controller
    {
        public static string constr = ConfigurationManager.ConnectionStrings["SeasonSun1213ConnectionString"].ConnectionString;
        SqlConnection con = new SqlConnection(constr);
        SeasonSun1213Entities16 bd = new SeasonSun1213Entities16();
        SIPGComentarios SIPGComentarios = new SIPGComentarios();
        SIPGProyeccion SIPGProyeccion = new SIPGProyeccion();
        SIPGVisitas SIPGVisitas = new SIPGVisitas();
        Image img = null;

        List<ClassVisitas> visitas = new List<ClassVisitas>();
        string idregion = "";
        public ActionResult Index()
        {
            return View();
        }

        //VISITAS
        public ActionResult Visitas()
        {
            if (Session["Nombre"] != null)
            {
                ViewData["Nombre"] = Session["Nombre"].ToString();
                List<ProdAgenteCat> agentesP = new List<ProdAgenteCat>();
                agentesP = bd.ProdAgenteCat.Where(x => x.Depto == "P" && x.Activo == true).OrderBy(x => x.Nombre).ToList();
                ViewBag.agentesP = agentesP;
            }
            else
            {
                return RedirectToAction("Index", "Login");
            }
            return View();
        }

        public JsonResult Grafica(int idAgen = 0)
        {
            int asesor;
            if ((short)Session["IdAgen"] == 1 && idAgen != 0)
            {
                asesor = idAgen;
            }
            else
            {
                asesor = (short)Session["IdAgen"];
            }

            //List<ClassVisitas> data = bd.Database.SqlQuery<ClassVisitas>("SET LANGUAGE Spanish; select SUM(C.TotalProductoresVisit)as Total , C.Mes from(select V.IdAgen, V.Mes, COUNT(V.Cod_Prod) as TotalProductoresVisit " +
            //    "from(select distinct V.IdAgen, V.Cod_Prod, CONVERT(VARCHAR(10), V.Fecha, 23) AS Fecha, DATENAME(month, V.Fecha) as Mes " +
            //    "from ProdVisitasCab V left join ProdCamposCat C on V.Cod_Prod = C.Cod_Prod LEFT join CatSemanas S ON V.Fecha between S.Inicio and S.Fin " +
            //    "where S.Temporada = (select temporada from catsemanas where getdate() between inicio and fin) and C.Activo = 'S')V group by V.IdAgen, V.Mes)C " +
            //    "where c.IdAgen = " + asesor + " group by C.IdAgen, C.Mes " +
            //    "order by(case when C.Mes = 'Julio' then 1 when C.Mes = 'Agosto' then 2 when C.Mes = 'Septiembre' then 3 when C.Mes = 'Octubre' then 4 when C.Mes = 'Noviembre' then 5 when C.Mes = 'Diciembre' then 6 when C.Mes = 'Enero' then 7 when C.Mes = 'Febrero' then 8 when C.Mes = 'Marzo' then 9 when C.Mes = 'Abril' then 10 when C.Mes = 'Mayo' then 11 when C.Mes = 'Junio' then 12 else '0' end)").ToList();
            List<ClassVisitas> data = bd.Database.SqlQuery<ClassVisitas>("SET LANGUAGE Spanish; select SUM(C.TotalCamposVisit) as Total, C.Mes from(select V.IdAgen, V.Mes, COUNT(V.Cod_Campo) as TotalCamposVisit " +
                "from(select distinct V.IdAgen, V.Cod_Prod, V.Cod_Campo, CONVERT(VARCHAR(10), V.Fecha, 23) AS Fecha, DATENAME(month, V.Fecha) as Mes " +
                "from ProdVisitasCab V LEFT JOIN ProdCamposCat C on V.IdAgen = C.IdAgen and V.Cod_Prod = C.Cod_Prod and V.Cod_Campo = C.Cod_Campo " +
                "LEFT join CatSemanas S ON V.Fecha between S.Inicio and S.Fin " +
                "where S.Temporada = (select temporada from catsemanas where getdate() between inicio and fin) and C.Activo = 'S')V group by V.IdAgen, V.Mes)C " +
                "where c.IdAgen = " + asesor + " group by C.IdAgen, C.Mes " +
                "order by(case when C.Mes = 'Julio' then 1 when C.Mes = 'Agosto' then 2 when C.Mes = 'Septiembre' then 3 when C.Mes = 'Octubre' then 4 when C.Mes = 'Noviembre' then 5 when C.Mes = 'Diciembre' then 6 when C.Mes = 'Enero' then 7 when C.Mes = 'Febrero' then 8 when C.Mes = 'Marzo' then 9 when C.Mes = 'Abril' then 10 when C.Mes = 'Mayo' then 11 when C.Mes = 'Junio' then 12 else '0' end) ").ToList();
            return Json(data, JsonRequestBehavior.AllowGet);
        }

        public JsonResult VisitasList(int idAgen = 0)
        {
            int asesor;
            if ((short)Session["IdAgen"] == 1 && idAgen != 0)
            {
                asesor = idAgen;
            }
            else
            {
                asesor = (short)Session["IdAgen"];
            }

            List<ClassVisitas> visitas = bd.Database.SqlQuery<ClassVisitas>("SET LANGUAGE Spanish; select V.IdAgen, V.Mes, V.TotalCampos, V.TotalCamposVisit, V.Eficiencia, cast(round(ISNULL(V.Efectividad*100,0),0) as varchar)+'%' as Efectividad, " +
                "round(V.Suma / (D1 + D2 + D3 + D4 + D5 + D6 + D7 + D8 + D9 + D10 + D11 + D12 + D13 + D14 + D15 + D16 + D17 + D18 + D19 + D20 + D21 + D22 + D23 + D24 + D25 + D26 + D27 + D28 + D29 + D30 + D31), 0) as Promedio, " +
                "V._1, V._2, V._3, V._4, V._5, V._6, V._7, V._8, V._9, V._10, V._11, V._12, V._13, V._14, V._15, V._16, V._17, V._18, V._19, V._20, V._21, V._22, V._23, V._24, V._25, V._26, V._27, V._28, V._29, V._30, V._31 " +
                "from(select V.IdAgen, V.Mes, V.TotalCampos, V.TotalCamposVisit, cast(round((V.Eficiencia * 100), 0) as varchar) + '%' as Eficiencia, (V.TotalCamposVisit / V.TotalCampos) as Efectividad, " +
                "(isnull(V.[1], 0) + isnull(V.[2], 0) + isnull(V.[3], 0) + isnull(V.[4], 0) + isnull(V.[5], 0) + isnull(V.[6], 0) + isnull(V.[7], 0) + isnull(V.[8], 0) + isnull(V.[9], 0) + isnull(V.[10], 0) + isnull(V.[11], 0) + isnull(V.[12], 0)+" +
                "isnull(V.[13], 0) + isnull(V.[14], 0) + isnull(V.[15], 0) + isnull(V.[16], 0) + isnull(V.[17], 0) + isnull(V.[18], 0) + isnull(V.[19], 0) + isnull(V.[20], 0) + isnull(V.[21], 0) + isnull(V.[22], 0) + isnull(V.[23], 0)+ isnull(V.[24], 0)+ " +
                "isnull(V.[25], 0) + isnull(V.[26], 0) + isnull(V.[27], 0) + isnull(V.[28], 0) + isnull(V.[29], 0) + isnull(V.[30], 0) + isnull(V.[31], 0)) AS Suma, " +
                "(case when V.[1] is not null then 1 else 0 END)AS D1,(case when V.[2] is not null then 1 else 0 END) AS D2,(case when V.[3] is not null then 1 else 0 END) AS D3,(case when V.[4] is not null then 1 else 0 END) AS D4, " +
                "(case when V.[5] is not null then 1 else 0 END) AS D5,(case when V.[6] is not null then 1 else 0 END) AS D6,(case when V.[7] is not null then 1 else 0 END) AS D7,(case when V.[8] is not null then 1 else 0 END) AS D8, " +
                "(case when V.[9] is not null then 1 else 0 END) AS D9,(case when V.[10] is not null then 1 else 0 END) AS D10,(case when V.[11] is not null then 1 else 0 END) AS D11,(case when V.[12] is not null then 1 else 0 END) AS D12, " +
                "(case when V.[13] is not null then 1 else 0 END) AS D13,(case when V.[14] is not null then 1 else 0 END) AS D14,(case when V.[15] is not null then 1 else 0 END) AS D15,(case when V.[16] is not null then 1 else 0 END)AS D16, " +
                "(case when V.[17] is not null then 1 else 0 END) AS D17, (case when V.[18] is not null then 1 else 0 END) AS D18, (case when V.[19] is not null then 1 else 0 END) AS D19,(case when V.[20] is not null then 1 else 0 END) AS D20, " +
                "(case when V.[21] is not null then 1 else 0 END) AS D21, (case when V.[22] is not null then 1 else 0 END) AS D22, (case when V.[23] is not null then 1 else 0 END) AS D23,(case when V.[24] is not null then 1 else 0 END) AS D24, " +
                "(case when V.[25] is not null then 1 else 0 END) AS D25,(case when V.[26] is not null then 1 else 0 END) AS D26, (case when V.[27] is not null then 1 else 0 END) AS D27, (case when V.[28] is not null then 1 else 0 END) AS D28, " +
                "(case when V.[29] is not null then 1 else 0 END) AS D29, (case when V.[30] is not null then 1 else 0 END) AS D30,(case when V.[31] is not null then 1 else 0 END) AS D31, " +
                "isnull(LTRIM(STR(V.[1], 3, 0)), '') as _1,isnull(LTRIM(STR(V.[2], 3, 0)), '') AS _2, isnull(LTRIM(STR(V.[3], 3, 0)), '') AS _3, isnull(LTRIM(STR(V.[4], 3, 0)), '') AS _4, isnull(LTRIM(STR(V.[5], 3, 0)), '') as _5,isnull(LTRIM(STR(V.[6], 3, 0)), '') AS _6, isnull(LTRIM(STR(V.[7], 3, 0)), '') as _7, " +
                "isnull(LTRIM(STR(V.[8], 3, 0)), '') AS _8, isnull(LTRIM(STR(V.[9], 3, 0)), '') AS _9, isnull(LTRIM(STR(V.[10], 3, 0)), '') AS _10, isnull(LTRIM(STR(V.[11], 3, 0)), '') AS _11, isnull(LTRIM(STR(V.[12], 3, 0)), '') AS _12, isnull(LTRIM(STR(V.[13], 3, 0)), '') AS _13, isnull(LTRIM(STR(V.[14], 3, 0)), '') AS _14, " +
                "isnull(LTRIM(STR(V.[15], 3, 0)), '') AS _15, isnull(LTRIM(STR(V.[16], 3, 0)), '') AS _16, isnull(LTRIM(STR(V.[17], 3, 0)), '') AS _17, isnull(LTRIM(STR(V.[18], 3, 0)), '') AS _18, isnull(LTRIM(STR(V.[19], 3, 0)), '') AS _19, isnull(LTRIM(STR(V.[20], 3, 0)), '') AS _20, isnull(LTRIM(STR(V.[21], 3, 0)), '') AS _21, " +
                "isnull(LTRIM(STR(V.[22], 3, 0)), '') AS _22, isnull(LTRIM(STR(V.[23], 3, 0)), '') AS _23, isnull(LTRIM(STR(V.[24], 3, 0)), '') AS _24, isnull(LTRIM(STR(V.[25], 3, 0)), '') AS _25, isnull(LTRIM(STR(V.[26], 3, 0)), '') AS _26, isnull(LTRIM(STR(V.[27], 3, 0)), '') AS _27, isnull(LTRIM(STR(V.[28], 3, 0)), '') AS _28, " +
                "isnull(LTRIM(STR(V.[29], 3, 0)), '') AS _29, isnull(LTRIM(STR(V.[30], 3, 0)), '') AS _30, isnull(LTRIM(STR(V.[31], 3, 0)), '') AS _31 " +
                "from(select * from(select V.IdAgen, cast(sum(V.Campos) as float) as Campos, V.Mes, V.Dia, T.TotalCampos, C.TotalCamposVisit, E.Eficiencia " +
                "FROM(select V.IdAgen, count(V.Cod_Campo) as Campos, V.Año, V.Mes, V.Semana, V.Dia, V.Fecha " +
                "from(select distinct V.IdAgen, V.Cod_Prod, V.Cod_Campo, CONVERT(VARCHAR(10), V.Fecha, 23) AS Fecha, DATENAME(year, V.Fecha) as Año, DATENAME(month, V.Fecha) as Mes, S.Semana, DATENAME(day, V.Fecha) as Dia " +
                "from ProdVisitasCab V left join ProdCamposCat C on V.IdAgen = C.IdAgen and V.Cod_Prod = C.Cod_Prod and V.Cod_Campo = C.Cod_Campo " +
                "LEFT join CatSemanas S ON V.Fecha between S.Inicio and S.Fin where S.Temporada = (select Temporada from Catsemanas where getdate() between Inicio and Fin) and C.Activo = 'S')V GROUP BY V.IdAgen, V.Año, V.Mes, V.Semana, V.Dia, V.Fecha)V " +
                "left join(select C.IdAgen, C.Mes, cast(SUM(C.TotalCamposVisit) as float) as TotalCamposVisit from(select V.IdAgen, V.Mes, COUNT(V.Cod_Prod) as TotalProductoresVisit, COUNT(V.Cod_Campo) as TotalCamposVisit " +
                "from(select distinct V.IdAgen, V.Cod_Prod, V.Cod_Campo, CONVERT(VARCHAR(10), V.Fecha, 23) AS Fecha, DATENAME(month, V.Fecha) as Mes from ProdVisitasCab V left join ProdCamposCat C on V.IdAgen = C.IdAgen and V.Cod_Prod = C.Cod_Prod and V.Cod_Campo = C.Cod_Campo " +
                "LEFT join CatSemanas S ON V.Fecha between S.Inicio and S.Fin where S.Temporada = (select Temporada from Catsemanas where getdate() between Inicio and Fin) and C.Activo = 'S')V group by V.IdAgen, V.Mes)C group by C.IdAgen, C.Mes)C ON  V.IdAgen = C.IdAgen AND V.Mes = C.Mes " +
                "LEFT JOIN(select C.IdAgen, SUM(C.Campos)as TotalCampos from(select C.IdAgen, count(C.Productores) AS Productores, count(C.Campos) AS Campos from(SELECT distinct IdAgen, Cod_Prod as Productores, Cod_Campo as Campos " +
                "from ProdCamposCat where Activo = 'S')C group by C.IdAgen)C group by C.IdAgen)T ON  V.IdAgen = T.IdAgen left join(select (CAST(sum(V.Dia) AS float)/ 26) as Eficiencia, V.Mes, V.IdAgen from(select V.IdAgen, V.Mes, count(V.Dia) as Dia " +
                "from(select distinct V.IdAgen, CONVERT(VARCHAR(10), V.Fecha, 23) AS Fecha, DATENAME(month, V.Fecha) as Mes, DATENAME(day, V.Fecha) as Dia from ProdVisitasCab V left join ProdCamposCat C on V.IdAgen = C.IdAgen and V.Cod_Prod = C.Cod_Prod and V.Cod_Campo = C.Cod_Campo " +
                "LEFT join CatSemanas S ON V.Fecha between S.Inicio and S.Fin where S.Temporada = (select temporada from catsemanas where getdate() between inicio and fin) and C.Activo = 'S' )V GROUP BY V.IdAgen, V.Fecha, V.Mes, V.Dia )V GROUP BY V.Mes, V.Dia, V.IdAgen )E ON  V.IdAgen = E.IdAgen AND V.Mes = E.Mes " +
                "WHERE V.IdAgen  = " + asesor + " " +
                "GROUP BY V.IdAgen, V.Fecha, V.Año, V.Mes, V.Dia, T.TotalCampos, C.TotalCamposVisit, E.Eficiencia)V " +
                "pivot(sum(Campos) FOR Dia in ([1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],[17],[18],[19],[20],[21],[22],[23],[24],[25],[26],[27],[28],[29],[30],[31])) as P)V " +
                ")V order by(case when V.Mes = 'Julio' then 1 when V.Mes = 'Agosto' then 2 when V.Mes = 'Septiembre' then 3 when V.Mes = 'Octubre' then 4 when V.Mes = 'Noviembre' then 5 when V.Mes = 'Diciembre' then 6 when V.Mes = 'Enero' then 7 when V.Mes = 'Febrero' then 8 when V.Mes = 'Marzo' then 9 when V.Mes = 'Abril' then 10 when V.Mes = 'Mayo' then 11 when V.Mes = 'Junio' then 12 else '0' end)").ToList();
            return Json(visitas, JsonRequestBehavior.AllowGet);
        }

        public ActionResult ReporteVisitas()
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

        public ActionResult ReporteView(string fecha)
        {
            if (Session["Nombre"] != null)
            {
                ViewData["Nombre"] = Session["Nombre"].ToString();
                RptView(fecha);
            }
            else
            {
                return RedirectToAction("Index", "Login");
            }
            return View();
        }

        public void RptView(string fecha)
        {
            Document doc = new Document(PageSize.LETTER, 20, 20, 10,20);//left, rigth, top, bottom           
            var output = new MemoryStream();
            PdfWriter pw = PdfWriter.GetInstance(doc, output);
            Image logo = Image.GetInstance("/Image/GIDDINGS_PRIMARY_STACKED_LOGO_DRIFT_RGB.png");
            pw.PageEvent = new HeaderFooter();

            Font _standardFont = new Font(Font.FontFamily.HELVETICA, 8, Font.NORMAL, BaseColor.BLACK);

            doc.AddTitle(fecha); 
            doc.AddCreator(Session["Nombre"].ToString());
            doc.Open();
            PdfPTable pdfTab = new PdfPTable(3);

            doc.Add(new Paragraph("Asesor: " + Session["Nombre"].ToString()));
            doc.Add(new Paragraph("Fecha: " + fecha));
            doc.Add(Chunk.NEWLINE);
            
            PdfPTable tbl = new PdfPTable(9);
            tbl.WidthPercentage = 100;
            tbl.HorizontalAlignment = Element.ALIGN_CENTER;

            float[] values = new float[9];
            values[0] = 80;
            values[1] = 200;
            values[2] = 150;
            values[3] = 75;
            values[4] = 200;
            values[5] = 150;
            values[6] = 150;
            values[7] = 200;
            values[8] = 300;
            tbl.SetWidths(values);           

            // Configuramos el título de las columnas de la tabla  
            PdfPCell clCod_prod = new PdfPCell(new Phrase("Código", _standardFont));
            clCod_prod.BorderWidth = 0;
            clCod_prod.BorderWidthBottom = 0.75f;

            PdfPCell clProductor = new PdfPCell(new Phrase("Productor", _standardFont));
            clProductor.BorderWidth = 0;
            clProductor.BorderWidthBottom = 0.75f;

            PdfPCell clCampo = new PdfPCell(new Phrase("Campo", _standardFont));
            clCampo.BorderWidth = 0;
            clCampo.BorderWidthBottom = 0.75f;

            PdfPCell clSector = new PdfPCell(new Phrase("Sector", _standardFont));
            clSector.BorderWidth = 0;
            clSector.BorderWidthBottom = 0.75f;

            PdfPCell clProducto = new PdfPCell(new Phrase("Producto", _standardFont));
            clProducto.BorderWidth = 0;
            clProducto.BorderWidthBottom = 0.75f;    

            PdfPCell clAtendio = new PdfPCell(new Phrase("Atendió", _standardFont));
            clAtendio.BorderWidth = 0;
            clAtendio.BorderWidthBottom = 0.75f;

            PdfPCell clEtapa = new PdfPCell(new Phrase("Etapa", _standardFont));
            clEtapa.BorderWidth = 0;
            clEtapa.BorderWidthBottom = 0.75f;

            PdfPCell clComentarios = new PdfPCell(new Phrase("Comentarios", _standardFont));
            clComentarios.BorderWidth = 0;
            clComentarios.BorderWidthBottom = 0.75f;

            PdfPCell X = new PdfPCell(new Phrase("Fotografía", _standardFont));
            X.BorderWidth = 0;
            X.BorderWidthBottom = 0.75f;

            // Añadimos las celdas a la tabla            
            tbl.AddCell(clCod_prod);
            tbl.AddCell(clProductor);
            tbl.AddCell(clCampo);
            tbl.AddCell(clSector);
            tbl.AddCell(clProducto);
            tbl.AddCell(clAtendio);
            tbl.AddCell(clEtapa);
            tbl.AddCell(clComentarios);
            tbl.AddCell(X);

            ReporteVisitasList(fecha);
            
            int x = 2;
            foreach (var item in visitas)
            {
                clCod_prod = new PdfPCell(new Phrase(item.Cod_Prod, _standardFont));
                clCod_prod.BorderWidth = 0;
               
                clProductor = new PdfPCell(new Phrase(item.Productor, _standardFont));
                clProductor.BorderWidth = 0;

                clCampo = new PdfPCell(new Phrase(Convert.ToString(item.Cod_Campo + " - " + item.Campo), _standardFont));
                clCampo.BorderWidth = 0;

                clSector = new PdfPCell(new Phrase(Convert.ToString(item.IdSector), _standardFont));
                clSector.BorderWidth = 0;

                clProducto = new PdfPCell(new Phrase(item.Tipo, _standardFont));
                clProducto.BorderWidth = 0;

                clAtendio = new PdfPCell(new Phrase(item.Atendio, _standardFont));
                clAtendio.BorderWidth = 0;

                clEtapa = new PdfPCell(new Phrase(item.Etapa, _standardFont));
                clEtapa.BorderWidth = 0;

                clComentarios = new PdfPCell(new Phrase(item.Comentarios, _standardFont));
                clComentarios.BorderWidth = 0;

                tbl.AddCell(clCod_prod);
                tbl.AddCell(clProductor);
                tbl.AddCell(clCampo);
                tbl.AddCell(clSector);
                tbl.AddCell(clProducto);
                tbl.AddCell(clAtendio);
                tbl.AddCell(clEtapa);
                tbl.AddCell(clComentarios);

                if (ImgExist(item.IdVisita))
                {
                    img.BorderWidth = 0;
                    img.Alignment = Element.ALIGN_CENTER;
                    float percentage = 0.0f;
                    percentage = 150 / img.Width;
                    img.ScalePercent(percentage * 100);
                }
                tbl.AddCell(img);
                x++;
            }

            //Logo            
           
            logo.SetAbsolutePosition(0f, 0f);
            logo.ScaleAbsolute(50f, 50f);

            doc.Add(logo);
            doc.Add(tbl);
            doc.Close();
            pw.Close();

            Response.Clear();
            Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            Response.AddHeader("Content-Disposition", string.Format("attachment;filename="+fecha+".pdf", "some string"));
            Response.BinaryWrite(output.ToArray());
            Response.End();
        }

        class HeaderFooter : PdfPageEventHelper {
            public override void OnEndPage(PdfWriter writer, Document document)
            {
                //base.OnEndPage(writer, document);
                PdfPTable tbHeader = new PdfPTable(3);
                tbHeader.TotalWidth = document.PageSize.Width - document.LeftMargin - document.RightMargin;
                tbHeader.DefaultCell.Border = 0;
                
                PdfPCell _cell = new PdfPCell(new Paragraph("Visitas diarias"));
                _cell.HorizontalAlignment = Element.ALIGN_CENTER;
                tbHeader.AddCell(_cell);
                tbHeader.AddCell(new Paragraph());

                PdfPTable tbFooter = new PdfPTable(3);
                tbFooter.TotalWidth = document.PageSize.Width - document.LeftMargin - document.RightMargin;
                tbFooter.DefaultCell.Border = 0;

                _cell = new PdfPCell(new Paragraph("Fruits - Giddings S.A.de C.V"));
                _cell.HorizontalAlignment = Element.ALIGN_CENTER;
                tbFooter.AddCell(_cell);
                tbFooter.AddCell(new Paragraph());
            }
        }
        public bool ImgExist(int IdVisita)
        {
            try
            {
                img = Image.GetInstance("//192.168.0.21/recursos season/VisitasProd/" + IdVisita + "/1.jpg");
                return true;
            }
            catch (Exception e) {
                e.ToString();
                return false;
            }            
        }

        public void ReportVisitList(string fecha)
        {
            visitas = bd.Database.SqlQuery<ClassVisitas>("Select cat.Nombre as Asesor,cab.IdVisita, sec.IdSector,cab.Cod_prod, prod.Nombre as Productor,cab.Cod_Campo, cam.Descripcion as Campo, rtrim(tpo.Descripcion) + ' - ' + rtrim(pto.Descripcion) as Tipo, pto.Descripcion as Producto, convert(VARCHAR(20), cab.Fecha, 103) as Fecha, " +
                "cab.Comentarios, cab.Atendio, isnull(icat.DescIncidencia,'') as DescIncidencia, isnull(etp.DescEtapa,'') as Etapa, isnull(det.Comentario,'') as Folio " +
                "FROM ProdVisitasCab cab " +
                "inner join ProdAgenteCat cat on cab.IdAgen = cat.IdAgen " +
                "left join ProdProductoresCat prod on cab.Cod_prod = prod.Cod_Prod " +
                "left join ProdCamposCat cam on cab.Cod_prod = cam.Cod_Prod and cab.Cod_Campo = cam.Cod_Campo " +
                "left join CatTiposProd tpo on cam.Tipo = tpo.Tipo " +
                "left join CatProductos pto on cam.Tipo = pto.Tipo and cam.Producto = pto.Producto " +
                "left join ProdVisitasDet det on cab.IdVisita = det.IdVisita " +
                "left join(select cab.IdVisita, i.IdIncidencia, STUFF((SELECT distinct ',' + i.DescIncidencia FROM ProdVisitasCab cab left join ProdVisitasDet det on cab.IdVisita = det.IdVisita left join ProdIncidenciasCat i on det.IdIncidencia = i.IdIncidencia " +
                "where convert(varchar, cab.Fecha, 23) = '" + fecha + "' and cab.IdAgen = " + (short)Session["IdAgen"] + " FOR XML PATH('')), 1, 1, '') as DescIncidencia " +
                "FROM ProdVisitasCab cab left join ProdVisitasDet det on cab.IdVisita = det.IdVisita left join ProdIncidenciasCat i on det.IdIncidencia = i.IdIncidencia)icat on icat.IdVisita = cab.IdVisita and det.IdIncidencia = icat.IdIncidencia " +
                "left join ProdVisitasSectores sec on cab.IdVisita = sec.IdVisita " +
                "left join ProdEtapasCat etp on det.IdEtapa = etp.IdEtapa " +
                "where convert(varchar, cab.Fecha, 23) = '" + fecha + "' and cat.IdAgen = " + (short)Session["IdAgen"] + " " +
                "group by cat.Nombre, cab.IdVisita, sec.IdSector, cab.Cod_prod, prod.Nombre, cab.Cod_Campo, cam.Descripcion, tpo.Descripcion, pto.Descripcion, cab.Fecha, cab.Comentarios, cab.Atendio, icat.DescIncidencia,etp.DescEtapa, det.Comentario, cab.IdSector " +
                "order by cab.Fecha, cab.Cod_prod, cab.Cod_Campo, sec.IdSector").ToList();
        }
        public JsonResult ReporteVisitasList(string fecha)
        {
            ReportVisitList(fecha);
            return Json(visitas, JsonRequestBehavior.AllowGet);
        }

        //imagen
        public JsonResult obtenerImagen(int IdVisita = 0)
        {
            string ruta = "//192.168.0.21/recursos season/VisitasProd/" + IdVisita + "/1.jpg";
            string Imagen = getImage(ruta);
            return Json(Imagen, JsonRequestBehavior.AllowGet);
        }

        public string getImage(string ruta)
        {
            try
            {
                byte[] bytesImagen = System.IO.File.ReadAllBytes(ruta);
                string imagenBase64 = Convert.ToBase64String(bytesImagen);

                string tipoContenido;

                switch (Path.GetExtension(ruta))
                {
                    case ".jpg":
                        {
                            tipoContenido = "image/jpg";
                            break;
                        }

                    default:
                        {
                            return null;
                        }
                }

                return string.Format("data:{0};base64,{1}", tipoContenido, imagenBase64);
            }
            catch (Exception e)
            {
                e.ToString();
                return null;
            }
        }

        //Excel Reporte visitas
        public void ReporteVisitasExcel(string fecha)
        {
            ViewData["Nombre"] = Session["Nombre"].ToString();

            //Range formatRange;
            //Application xlApp = new Application();
            //Excel.Workbook xlWorkBook;
            //Excel.Worksheet xlWorkSheet;
            //object misValue = Missing.Value;

            //xlWorkBook = xlApp.Workbooks.Add(misValue);
            //xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            ExcelPackage excel = new ExcelPackage();
            ExcelWorksheet ws = excel.Workbook.Worksheets.Add(fecha);
            ws.Cells["A1"].Value = "Asesor";
            ws.Cells["B1"].Value = "Codigo";
            ws.Cells["C1"].Value = "Productor";
            ws.Cells["D1"].Value = "Cod_Campo";
            ws.Cells["E1"].Value = "Campo";
            ws.Cells["F1"].Value = "Sector";
            ws.Cells["G1"].Value = "Producto";
            ws.Cells["H1"].Value = "Fecha";
            ws.Cells["I1"].Value = "Comentarios";
            ws.Cells["J1"].Value = "Atendió";
            ws.Cells["K1"].Value = "Etapa";
            ws.Cells["L1"].Value = "Incidencia";
            ws.Cells["M1"].Value = "Folio";

            //xlWorkSheet.Cells[1, 1] = "Codigo";
            //xlWorkSheet.Cells[1, 2] = "Productor";
            //xlWorkSheet.Cells[1, 3] = "Cod_Campo";
            //xlWorkSheet.Cells[1, 4] = "Campo";
            //xlWorkSheet.Cells[1, 5] = "Sector";
            //xlWorkSheet.Cells[1, 6] = "Producto";
            //xlWorkSheet.Cells[1, 7] = "Fecha";
            //xlWorkSheet.Cells[1, 8] = "Comentarios";
            //xlWorkSheet.Cells[1, 9] = "Atendió";
            //xlWorkSheet.Cells[1, 10] = "Etapa";
            //xlWorkSheet.Cells[1, 11] = "Incidencia";
            //xlWorkSheet.Cells[1, 12] = "Folio";
            //xlWorkSheet.Cells[1, 13] = "";

            ReportVisitList(fecha);
            int x = 2;
            foreach (var item in visitas)
            {
                ws.Cells[string.Format("A{0}", x)].Value = item.Asesor;
                ws.Cells[string.Format("B{0}", x)].Value = item.Cod_Prod;
                ws.Cells[string.Format("C{0}", x)].Value = item.Productor;
                ws.Cells[string.Format("D{0}", x)].Value = item.Cod_Campo;
                ws.Cells[string.Format("E{0}", x)].Value = item.Campo;
                ws.Cells[string.Format("F{0}", x)].Value = item.IdSector;
                ws.Cells[string.Format("G{0}", x)].Value = item.Tipo;
                ws.Cells[string.Format("H{0}", x)].Value = item.Fecha;
                ws.Cells[string.Format("I{0}", x)].Value = item.Comentarios;
                ws.Cells[string.Format("J{0}", x)].Value = item.Atendio;
                ws.Cells[string.Format("K{0}", x)].Value = item.Etapa;
                ws.Cells[string.Format("L{0}", x)].Value = item.DescIncidencia;
                ws.Cells[string.Format("M{0}", x)].Value = item.Folio;
                //xlWorkSheet.Cells[x,1] = item.Cod_Prod;
                //xlWorkSheet.Cells[x,2] = item.Productor;
                //xlWorkSheet.Cells[x,3] = item.Cod_Campo;
                //xlWorkSheet.Cells[x,4] = item.Campo;
                //xlWorkSheet.Cells[x,5] = item.IdSector;
                //xlWorkSheet.Cells[x,6] = item.Tipo;
                //xlWorkSheet.Cells[x,7] = item.Fecha;
                //xlWorkSheet.Cells[x,8] = item.Comentarios;
                //xlWorkSheet.Cells[x,9] = item.Atendio;
                //xlWorkSheet.Cells[x,10] = item.Etapa;
                //xlWorkSheet.Cells[x,11] = item.DescIncidencia;
                //xlWorkSheet.Cells[x,12] = item.Folio;
                //xlWorkSheet.Shapes.AddPicture("\\192.168.0.21\recursos season\\VisitasProd\\" + item.IdVisita + "\\1.jpg", MsoTriState.msoFalse,MsoTriState.msoCTrue, 200, 20, 100, 100);//rigth,top,bottom, left "C:\\Users\\marholym\\Pictures\\Captura.png"
                x++;
            }
            ws.Cells["A:M"].AutoFitColumns();

            //xlWorkBook.SaveAs("C:\\Users\\ReporteVisitasExcel.xls", Excel.XlFileFormat.xlWorkbookNormal, misValue,misValue,misValue,misValue,Excel.XlSaveAsAccessMode.xlExclusive,misValue, misValue, misValue, misValue);
            //xlWorkBook.Close(true,misValue,misValue);
            //xlApp.Quit();
            //Marshal.ReleaseComObject(xlWorkSheet);
            //Marshal.ReleaseComObject(xlWorkBook);
            //Marshal.ReleaseComObject(xlApp);


            Response.Clear();
            Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            Response.AddHeader("Content-Disposition", string.Format("attachment;filename=" + fecha + ".xlsx"));
            Response.BinaryWrite(excel.GetAsByteArray());
            Response.End();
        }

        //PRODUCCION
        public ActionResult Produccion()
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

        public JsonResult ProduccionList()
        {
            //List<ClassCurva> datos = bd.Database.SqlQuery<ClassCurva>("select V.Cod_Prod, V.Nombre, cast('$'+V.SaldoFinal as varchar) as SaldoFinal, V.Pronostico, V.Entregado, V.Diferencia, V.PronosticoAA, V.DiferenciaAA, isnull(V.Semana,0) as Semana, V.PronosticoSA, V.EntregadoSA, V.DiferenciaSA from(select * from (select V.Cod_Prod, V.Nombre, round(isnull(V.SaldoFinal,0),0) as SaldoFinal, round(isnull(V.Pronostico,0),0) as Pronostico, round(isnull(V.Entregado,0),0) as Entregado, round(isnull(((V.Diferencia*100)-100),0),0) as Diferencia, round(isnull(V.PronosticoAA,0),0) as PronosticoAA, round(isnull(((V.DiferenciaAA*100)-100),0),0) as DiferenciaAA, V.Semana, round(isnull(V.PronosticoSA,0),0) as PronosticoSA, round(isnull(V.EntregadoSA,0),0) as EntregadoSA, round(isnull(((V.DiferenciaSA*100)-100),0),0) as DiferenciaSA from(select * from(select *, (V.Entregado/ V.Pronostico) AS Diferencia, (V.Entregado/ V.PronosticoAA) AS DiferenciaAA, (V.EntregadoSA / V.PronosticoSA) AS DiferenciaSA from(select F.Cod_Prod, F.Nombre, F.SaldoFinal, V.Pronostico as Pronostico, V.Entregado, SA.PronosticoSA, SA.EntregadoSA, SA.Semana, AA.PronosticoAA from(select Cod_Prod, Nombre, (sum(Saldo) + sum(SaldoAGQ)) as SaldoFinal from dbo.fnRptSaldosFinanciamiento(GetDate(),GetDate(),GetDate(),50) group by Cod_Prod, Nombre)F left join(select V.IdAgen, V.Cod_prod, V.Cod_Campo, E.Entregado, sum(V.Pronostico) as Pronostico from(select * from(select * from(select IdAgen, Cod_Prod,Cod_Campo, SUM(isnull([27], 0)) as [27], sum(isnull([28], 0)) as [28], sum(isnull([29], 0)) as[29], sum(isnull([30], 0)) as [30], sum(isnull([31], 0)) as [31], sum(isnull([32], 0)) as [32], sum(isnull([33], 0)) as [33], sum(isnull([34], 0)) as [34], sum(isnull([35], 0)) as [35], sum(isnull([36], 0)) as [36], sum(isnull([37], 0)) as [37], sum(isnull([38], 0)) as [38], sum(isnull([39], 0)) as [39], sum(isnull([40], 0)) as [40], sum(isnull([41], 0)) as [41], sum(isnull([42], 0)) as [42], sum(isnull([43], 0)) as [43], sum(isnull([44], 0)) as [44], sum(isnull([45], 0)) as [45], sum(isnull([46], 0)) as [46], sum(isnull([47], 0)) as [47], sum(isnull([48], 0)) as [48], sum(isnull([49], 0)) as [49], sum(isnull([50], 0)) as [50], sum(isnull([51], 0)) as [51], sum(isnull([52], 0)) as [52], sum(isnull([1], 0)) as [1], sum(isnull([2], 0)) as [2], sum(isnull([3], 0)) as [3], sum(isnull([4], 0)) as [4], sum(isnull([5], 0)) as [5], sum(isnull([6], 0)) as [6], sum(isnull([7], 0)) as [7], sum(isnull([8], 0)) as [8], sum(isnull([9], 0)) as [9], sum(isnull([10], 0)) as [10], sum(isnull([11], 0)) as [11], sum(isnull([12], 0)) as [12], sum(isnull([13], 0)) as [13], sum(isnull([14], 0)) as [14], sum(isnull([15], 0)) as [15],sum(isnull([16], 0)) as [16], sum(isnull([17], 0)) as [17], sum(isnull([18], 0)) as [18], sum(isnull([19], 0)) as [19], sum(isnull([20], 0)) as [20], sum(isnull([21], 0)) as [21], sum(isnull([22], 0)) as [22], sum(isnull([23], 0)) as [23], sum(isnull([24], 0)) as [24], sum(isnull([25], 0)) as [25], sum(isnull([26], 0)) as [26] from SIPGProyeccion P where Temporada = (select Temporada from CatSemanas where getdate() between Inicio and Fin) and Estado='A' group by IdAgen, Cod_Prod,Cod_Campo)V UNPIVOT(Pronostico FOR Semana in ([27],[28],[29],[30],[31],[32],[33],[34],[35],[36],[37],[38],[39],[40],[41],[42],[43],[44],[45],[46],[47],[48],[49],[50],[51],[52],[1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],[17],[18],[19],[20],[21],[22],[23],[24],[25],[26]))X)V where Pronostico <> 0)V left join(select C.IdAgen, C.Cod_Prod, C.Cod_Campo, sum(C.Convertidas) as Entregado FROM(select IdAgen, Cod_Prod, Cod_Campo, sum(Convertidas) as Convertidas from SEasonsun1213..UV_ProdRecepcion where CodEstatus <> 'C' and Temporada = (select Temporada from CatSemanas where getdate() between Inicio and Fin) group by IdAgen, Cod_prod, Cod_Campo Union All select IdAgen, Cod_Prod, Cod_Campo, sum(Convertidas) as Convertidas from SeasonPlan..UV_ProdRecepcion where CodEstatus <> 'C' and Temporada = (select Temporada from CatSemanas where getdate() between Inicio and Fin) group by IdAgen, Cod_prod, Cod_Campo)C GROUP BY C.IdAgen, C.Cod_Prod, C.Cod_Campo)E on V.IdAgen = E.IdAgen AND V.Cod_Prod = E.Cod_Prod and V.Cod_Campo = E.Cod_Campo group by V.IdAgen, V.Cod_prod, V.Cod_Campo, E.Entregado)V ON F.Cod_Prod = V.Cod_Prod left join(select V.IdAgen, V.Cod_Prod, V.Cod_Campo, sum(V.Pronostico) as PronosticoAA FROM CatSemanas S left join (select * from(select * from(select IdAgen, Cod_Prod, Cod_Campo, SUM(isnull([27], 0)) as [27], sum(isnull([28], 0)) as [28], sum(isnull([29], 0)) as [29], sum(isnull([30], 0)) as [30], sum(isnull([31], 0)) as [31], sum(isnull([32], 0)) as [32], sum(isnull([33], 0)) as [33], sum(isnull([34], 0)) as [34], sum(isnull([35], 0)) as [35], sum(isnull([36], 0)) as [36], sum(isnull([37], 0)) as [37], sum(isnull([38], 0)) as [38], sum(isnull([39], 0)) as [39], sum(isnull([40], 0)) as [40], sum(isnull([41], 0)) as [41], sum(isnull([42], 0)) as [42], sum(isnull([43], 0)) as [43], sum(isnull([44], 0)) as [44], sum(isnull([45], 0)) as [45], sum(isnull([46], 0)) as [46], sum(isnull([47], 0)) as [47], sum(isnull([48], 0)) as [48], sum(isnull([49], 0)) as [49], sum(isnull([50], 0)) as [50], sum(isnull([51], 0)) as [51], sum(isnull([52], 0)) as [52], sum(isnull([1], 0)) as [1], sum(isnull([2], 0)) as [2], sum(isnull([3], 0)) as [3], sum(isnull([4], 0)) as [4], sum(isnull([5], 0)) as [5], sum(isnull([6], 0)) as [6], sum(isnull([7], 0)) as [7], sum(isnull([8], 0)) as [8], sum(isnull([9], 0)) as [9], sum(isnull([10], 0)) as [10], sum(isnull([11], 0)) as [11], sum(isnull([12], 0)) as [12], sum(isnull([13], 0)) as [13], sum(isnull([14], 0)) as [14], sum(isnull([15], 0)) as [15],sum(isnull([16], 0)) as [16], sum(isnull([17], 0)) as [17], sum(isnull([18], 0)) as [18], sum(isnull([19], 0)) as [19], sum(isnull([20], 0)) as [20], sum(isnull([21], 0)) as [21], sum(isnull([22], 0)) as [22], sum(isnull([23], 0)) as [23], sum(isnull([24], 0)) as [24], sum(isnull([25], 0)) as [25], sum(isnull([26], 0)) as [26] from SIPGProyeccion P where Temporada = (select Temporada from CatSemanas where getdate() between Inicio and Fin) and Estado='A' group by IdAgen, Cod_Prod, Cod_Campo)V UNPIVOT(Pronostico FOR Semana in ([27],[28],[29],[30],[31],[32],[33],[34],[35],[36],[37],[38],[39],[40],[41],[42],[43],[44],[45],[46],[47],[48],[49],[50],[51],[52],[1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],[17],[18],[19],[20],[21],[22],[23],[24],[25],[26]))X)V where Pronostico <> 0)V on S.Semana=V.Semana " +
            //    "where S.Inicio between '20190701' and getdate() group by  V.IdAgen, V.Cod_Prod, V.Cod_Campo)AA ON V.IdAgen = AA.IdAgen and V.Cod_Prod = AA.Cod_Prod AND V.Cod_Campo = AA.Cod_Campo LEFT JOIN(select PA.IdAgen, PA.Cod_Prod, PA.Cod_Campo, round(PA.Pronostico,0) AS PronosticoSA, EA.Entregado AS EntregadoSA, S.Semana FROM CatSemanas S left join(select* from(select* from(select IdAgen, Cod_Prod, Cod_Campo, SUM(isnull([27], 0)) as [27], sum(isnull([28], 0)) as [28], sum(isnull([29], 0)) as[29], sum(isnull([30], 0)) as [30], sum(isnull([31], 0)) as [31], sum(isnull([32], 0)) as [32], sum(isnull([33], 0)) as [33], sum(isnull([34], 0)) as [34], sum(isnull([35], 0)) as [35], sum(isnull([36], 0)) as [36], sum(isnull([37], 0)) as [37], sum(isnull([38], 0)) as [38], sum(isnull([39], 0)) as [39], sum(isnull([40], 0)) as [40], sum(isnull([41], 0)) as [41], sum(isnull([42], 0)) as [42], sum(isnull([43], 0)) as [43], sum(isnull([44], 0)) as [44], sum(isnull([45], 0)) as [45], sum(isnull([46], 0)) as [46], sum(isnull([47], 0)) as [47], sum(isnull([48], 0)) as [48], sum(isnull([49], 0)) as [49], sum(isnull([50], 0)) as [50], sum(isnull([51], 0)) as [51], sum(isnull([52], 0)) as [52], sum(isnull([1], 0)) as [1], sum(isnull([2], 0)) as [2], sum(isnull([3], 0)) as [3], sum(isnull([4], 0)) as [4], sum(isnull([5], 0)) as [5], sum(isnull([6], 0)) as [6], sum(isnull([7], 0)) as [7], sum(isnull([8], 0)) as [8], sum(isnull([9], 0)) as [9], sum(isnull([10], 0)) as [10], sum(isnull([11], 0)) as [11], sum(isnull([12], 0)) as [12], sum(isnull([13], 0)) as [13], sum(isnull([14], 0)) as [14], sum(isnull([15], 0)) as [15],sum(isnull([16], 0)) as [16], sum(isnull([17], 0)) as [17], sum(isnull([18], 0)) as [18], sum(isnull([19], 0)) as [19], sum(isnull([20], 0)) as [20], sum(isnull([21], 0)) as [21], sum(isnull([22], 0)) as [22], sum(isnull([23], 0)) as [23], sum(isnull([24], 0)) as [24], sum(isnull([25], 0)) as [25], sum(isnull([26], 0)) as [26] FROM SIPGProyeccion P where Temporada = (select Temporada from CatSemanas where getdate() between Inicio and Fin) and Estado = 'A' group by IdAgen, Cod_Prod, Cod_Campo)V UNPIVOT(Pronostico FOR Semana in ([27],[28],[29],[30],[31],[32],[33],[34],[35],[36],[37],[38],[39],[40],[41],[42],[43],[44],[45],[46],[47],[48],[49],[50],[51],[52],[1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],[17],[18],[19],[20],[21],[22],[23],[24],[25],[26]))X)V where Pronostico <> 0)PA ON S.Semana = PA.Semana Left Join(select C.IdAgen, C.Cod_Prod, C.Cod_Campo, sum(C.Convertidas) as Entregado, C.Semana FROM(select IdAgen, Cod_Prod, Cod_Campo, sum(Convertidas) as Convertidas, Semana FROM SEasonsun1213..UV_ProdRecepcion where CodEstatus<> 'C' and Temporada = (select Temporada from CatSemanas where getdate() between Inicio and Fin) group by IdAgen, Cod_Prod, Cod_Campo, Semana Union All SELECT IdAgen, Cod_Prod, Cod_Campo, sum(Convertidas) as Convertidas, Semana FROM SeasonPlan..UV_ProdRecepcion where CodEstatus<> 'C' and Temporada = (select Temporada from CatSemanas where getdate() between Inicio and Fin) group by IdAgen, Cod_Prod, Cod_Campo, Semana)C GROUP BY C.IdAgen, C.Cod_Prod, C.Cod_Campo, C.Semana)EA ON PA.IdAgen = EA.IdAgen and PA.Cod_Prod = EA.Cod_Prod and PA.Cod_Campo = EA.Cod_Campo and S.Semana = EA.Semana WHERE getdate() between S.Inicio and S.Fin)SA on V.IdAgen = SA.IdAgen and V.Cod_Prod = SA.Cod_Prod AND V.Cod_Campo = SA.Cod_Campo " +
            //    "WHERE V.IdAgen = " + (short)Session["IdAgen"] + " GROUP BY F.SaldoFinal, F.Cod_Prod, F.Nombre, V.Pronostico, V.Entregado, SA.PronosticoSA, SA.EntregadoSA, SA.Semana, AA.PronosticoAA)V)V GROUP BY V.Cod_Prod, V.Nombre, V.SaldoFinal, V.Pronostico, V.Entregado, V.Diferencia, V.PronosticoAA, V.DiferenciaAA, V.PronosticoSA, V.EntregadoSA, DiferenciaSA, V.Semana)V)V)V ORDER BY V.Diferencia").ToList();
            //return Json(datos, JsonRequestBehavior.AllowGet);

            List<ClassCurva> datos = bd.Database.SqlQuery<ClassCurva>("select V.Cod_Prod, V.Nombre as Productor, CONVERT(VARCHAR(50), CAST(V.SaldoFinal AS MONEY),1) as SaldoFinal, V.Entregado as EntregadoT, isnull(V.Semana,0) as Semana, V.EntregadoSA " +
                "FROM(select V.Cod_Prod, V.Nombre, V.SaldoFinal, V.Entregado, V.Semana, V.EntregadoSA " +
                "from(select V.Cod_Prod, V.Nombre, round(isnull(V.SaldoFinal, 0), 0) as SaldoFinal, round(isnull(V.Entregado, 0), 0) as Entregado, V.Semana, round(isnull(V.EntregadoSA, 0), 0) as EntregadoSA " +
                "from(select F.Cod_Prod, F.Nombre, F.SaldoFinal, V.Entregado, SA.EntregadoSA, SA.Semana " +
                "from(select Cod_Prod, Nombre, (sum(Saldo) + sum(SaldoAGQ)) as SaldoFinal " +
                "from dbo.fnRptSaldosFinanciamiento(GetDate(), GetDate(), GetDate(), 50) group by Cod_Prod, Nombre)F " +
                "left join(select E.IdAgen, E.Cod_prod, sum(E.Entregado) as Entregado " +
                "from(select C.IdAgen, C.Cod_Prod, C.Cod_Campo, sum(C.Convertidas) as Entregado " +
                "FROM(select IdAgen, Cod_Prod, Cod_Campo, sum(Convertidas) as Convertidas from SEasonsun1213..UV_ProdRecepcion " +
                "where CodEstatus <> 'C' and Temporada = (select Temporada from CatSemanas where getdate() between Inicio and Fin) group by IdAgen, Cod_prod, Cod_Campo " +
                ")C GROUP BY C.IdAgen, C.Cod_Prod, C.Cod_Campo)E group by E.IdAgen, E.Cod_prod)V ON F.Cod_Prod = V.Cod_Prod " +
                "LEFT JOIN(select EA.IdAgen, EA.Cod_Prod, sum(EA.Entregado) AS EntregadoSA, S.Semana " +
                "FROM CatSemanas S " +
                "Left Join(select C.IdAgen, C.Cod_Prod, C.Cod_Campo, sum(C.Convertidas) as Entregado, C.Semana " +
                "FROM(select IdAgen, Cod_Prod, Cod_Campo, sum(Convertidas) as Convertidas, Semana FROM SEasonsun1213..UV_ProdRecepcion " +
                "where CodEstatus <> 'C' and Temporada = (select Temporada from CatSemanas where getdate() between Inicio and Fin) group by IdAgen, Cod_Prod, Cod_Campo, Semana " +
                ")C GROUP BY C.IdAgen, C.Cod_Prod, C.Cod_Campo, C.Semana)EA ON S.Semana = EA.Semana " +
                "WHERE getdate() between S.Inicio and S.Fin group by EA.IdAgen, EA.Cod_Prod, S.Semana)SA on V.IdAgen = SA.IdAgen and V.Cod_Prod = SA.Cod_Prod " +
                "WHERE V.IdAgen = " + (short)Session["IdAgen"] + " " +
                "GROUP BY F.SaldoFinal, F.Cod_Prod, F.Nombre, V.Entregado, SA.EntregadoSA, SA.Semana)V)V)V " +
                "GROUP BY V.Cod_Prod, V.Nombre, V.SaldoFinal, V.Entregado, V.EntregadoSA, V.Semana " +
                "ORDER BY V.Cod_Prod").ToList();
            return Json(datos, JsonRequestBehavior.AllowGet);
        }

        public JsonResult GraficaProduccion()
        {
            //List<ClassCurva> data = bd.Database.SqlQuery<ClassCurva>("SELECT *, ROUND((V.Entregado/V.Pronostico)*100,0) AS Diferencia, ROUND((V.Entregado/V.PronosticoAA)*100,0) AS DiferenciaAA, ROUND((V.EntregadoSA/V.PronosticoSA)*100,0) AS DiferenciaSA FROM(Select V.IdAgen, round(sum(V.Entregado),0) as Entregado, round(sum(V.Pronostico),0) as Pronostico, round(ISNULL(AA.PronosticoAA,0),0) AS PronosticoAA, round(SA.PronosticoSA,0) as PronosticoSA, round(SA.EntregadoSA,0) as EntregadoSA, SA.Semana from(select V.IdAgen, ROUND(isnull(E.Entregado,0),0) AS Entregado, ROUND(isnull(V.Pronostico,0),0) as Pronostico FROM ProdCamposCat C LEFT JOIN (select V.IdAgen, V.Cod_Prod, V.Cod_Campo, sum(V.Pronostico) as Pronostico from(select * from(select IdAgen, Cod_Prod, Cod_Campo, SUM(isnull([27], 0)) as [27], sum(isnull([28], 0)) as [28], sum(isnull([29], 0)) as[29], sum(isnull([30], 0)) as [30], sum(isnull([31], 0)) as [31], sum(isnull([32], 0)) as [32], sum(isnull([33], 0)) as [33], sum(isnull([34], 0)) as [34], sum(isnull([35], 0)) as [35], sum(isnull([36], 0)) as [36], sum(isnull([37], 0)) as [37], sum(isnull([38], 0)) as [38], sum(isnull([39], 0)) as [39], sum(isnull([40], 0)) as [40], sum(isnull([41], 0)) as [41], sum(isnull([42], 0)) as [42], sum(isnull([43], 0)) as [43], sum(isnull([44], 0)) as [44], sum(isnull([45], 0)) as [45], sum(isnull([46], 0)) as [46], sum(isnull([47], 0)) as [47], sum(isnull([48], 0)) as [48], sum(isnull([49], 0)) as [49], sum(isnull([50], 0)) as [50], sum(isnull([51], 0)) as [51], sum(isnull([52], 0)) as [52], sum(isnull([1], 0)) as [1], sum(isnull([2], 0)) as [2], sum(isnull([3], 0)) as [3], sum(isnull([4], 0)) as [4], sum(isnull([5], 0)) as [5], sum(isnull([6], 0)) as [6], sum(isnull([7], 0)) as [7], sum(isnull([8], 0)) as [8], sum(isnull([9], 0)) as [9], sum(isnull([10], 0)) as [10], sum(isnull([11], 0)) as [11], sum(isnull([12], 0)) as [12], sum(isnull([13], 0)) as [13], sum(isnull([14], 0)) as [14], sum(isnull([15], 0)) as [15],sum(isnull([16], 0)) as [16], sum(isnull([17], 0)) as [17], sum(isnull([18], 0)) as [18], sum(isnull([19], 0)) as [19], sum(isnull([20], 0)) as [20], sum(isnull([21], 0)) as [21], sum(isnull([22], 0)) as [22], sum(isnull([23], 0)) as [23], sum(isnull([24], 0)) as [24], sum(isnull([25], 0)) as [25], sum(isnull([26], 0)) as [26] from SIPGProyeccion P where Temporada = (select Temporada from CatSemanas where getdate() between Inicio and Fin) group by IdAgen, Cod_Prod, Cod_Campo)V UNPIVOT(Pronostico FOR Semana in ([27],[28],[29],[30],[31],[32],[33],[34],[35],[36],[37],[38],[39],[40],[41],[42],[43],[44],[45],[46],[47],[48],[49],[50],[51],[52],[1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],[17],[18],[19],[20],[21],[22],[23],[24],[25],[26]))X)V where Pronostico <> 0 group by V.IdAgen, V.Cod_Prod, V.Cod_Campo)V ON C.IdAgen=V.IdAgen and C.Cod_Prod = V.Cod_Prod and C.Cod_Campo=V.Cod_Campo left join(select C.IdAgen, C.Cod_Prod, C.Cod_Campo, SUM(C.Convertidas) as Entregado FROM(select IdAgen, Cod_Prod, Cod_Campo, sum(ISNULL(Convertidas,0)) as Convertidas FROM SEasonsun1213..UV_ProdRecepcion where CodEstatus <> 'C' and Temporada = (select Temporada from CatSemanas where getdate() between Inicio and Fin) group by IdAgen, Cod_prod,Cod_Campo Union All SELECT IdAgen, Cod_Prod, Cod_Campo, sum(ISNULL(Convertidas,0)) as Convertidas FROM SeasonPlan..UV_ProdRecepcion where CodEstatus <> 'C' and Temporada = (select Temporada from CatSemanas where getdate() between Inicio and Fin) group by IdAgen, Cod_prod,Cod_Campo)C GROUP BY C.IdAgen, C.Cod_Prod,Cod_Campo)E on V.IdAgen=E.IdAgen and V.Cod_Prod = E.Cod_Prod and V.Cod_Campo=E.Cod_Campo WHERE C.Activo='S' group by V.IdAgen, E.Entregado, V.Pronostico)V left join(SELECT V.IdAgen, SUM(isnull(V.Pronostico,0)) as PronosticoAA FROM ProdCamposCat C left join (select * from(select * from(select IdAgen, Cod_Prod, Cod_Campo, SUM(isnull([27], 0)) as [27], sum(isnull([28], 0)) as [28], sum(isnull([29], 0)) as [29], sum(isnull([30], 0)) as [30], sum(isnull([31], 0)) as [31], sum(isnull([32], 0)) as [32], sum(isnull([33], 0)) as [33], sum(isnull([34], 0)) as [34], sum(isnull([35], 0)) as [35], sum(isnull([36], 0)) as [36], sum(isnull([37], 0)) as [37], sum(isnull([38], 0)) as [38], sum(isnull([39], 0)) as [39], sum(isnull([40], 0)) as [40], sum(isnull([41], 0)) as [41], sum(isnull([42], 0)) as [42], sum(isnull([43], 0)) as [43], sum(isnull([44], 0)) as [44], sum(isnull([45], 0)) as [45], sum(isnull([46], 0)) as [46], sum(isnull([47], 0)) as [47], sum(isnull([48], 0)) as [48], sum(isnull([49], 0)) as [49], sum(isnull([50], 0)) as [50], sum(isnull([51], 0)) as [51], sum(isnull([52], 0)) as [52], sum(isnull([1], 0)) as [1], sum(isnull([2], 0)) as [2], sum(isnull([3], 0)) as [3], sum(isnull([4], 0)) as [4], sum(isnull([5], 0)) as [5], sum(isnull([6], 0)) as [6], sum(isnull([7], 0)) as [7], sum(isnull([8], 0)) as [8], sum(isnull([9], 0)) as [9], sum(isnull([10], 0)) as [10], sum(isnull([11], 0)) as [11], sum(isnull([12], 0)) as [12], sum(isnull([13], 0)) as [13], sum(isnull([14], 0)) as [14], sum(isnull([15], 0)) as [15],sum(isnull([16], 0)) as [16], sum(isnull([17], 0)) as [17], sum(isnull([18], 0)) as [18], sum(isnull([19], 0)) as [19], sum(isnull([20], 0)) as [20], sum(isnull([21], 0)) as [21], sum(isnull([22], 0)) as [22], sum(isnull([23], 0)) as [23], sum(isnull([24], 0)) as [24], sum(isnull([25], 0)) as [25], sum(isnull([26], 0)) as [26] from SIPGProyeccion P where Temporada = (select Temporada from CatSemanas where getdate() between Inicio and Fin) group by IdAgen, Cod_Prod, Cod_Campo)V UNPIVOT(Pronostico FOR Semana in ([27],[28],[29],[30],[31],[32],[33],[34],[35],[36],[37],[38],[39],[40],[41],[42],[43],[44],[45],[46],[47],[48],[49],[50],[51],[52],[1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],[17],[18],[19],[20],[21],[22],[23],[24],[25],[26]))X)V where Pronostico <> 0)V on C.IdAgen=V.IdAgen and C.Cod_Prod = V.Cod_Prod and C.Cod_Campo=V.Cod_Campo left join(select C.IdAgen, C.Cod_Prod, C.Cod_Campo, C.Semana,SUM(C.Convertidas) as Entregado FROM(select IdAgen, Cod_Prod, Cod_Campo, Semana,sum(ISNULL(Convertidas,0)) as Convertidas FROM SEasonsun1213..UV_ProdRecepcion where CodEstatus <> 'C' and Temporada = (select Temporada from CatSemanas where getdate() between Inicio and Fin) group by IdAgen, Cod_prod,Cod_Campo,Semana Union All SELECT IdAgen, Cod_Prod, Cod_Campo, Semana,sum(ISNULL(Convertidas,0)) as Convertidas FROM SeasonPlan..UV_ProdRecepcion where CodEstatus <> 'C' and Temporada = (select Temporada from CatSemanas where getdate() between Inicio and Fin) group by IdAgen, Cod_prod,Cod_Campo,Semana)C GROUP BY C.IdAgen, C.Cod_Prod,Cod_Campo, C.Semana)E on V.IdAgen=E.IdAgen and V.Cod_Prod = E.Cod_Prod and V.Cod_Campo=E.Cod_Campo AND V.Semana=E.Semana left join CatSemanas S on S.Semana=V.Semana " +
            //    "where S.Inicio between '20190701' and getdate() group by V.IdAgen)AA ON V.IdAgen = AA.IdAgen LEFT JOIN(SELECT PA.IdAgen, sum(round(PA.Pronostico,0)) AS PronosticoSA, sum(ISNULL(EA.Entregado, 0)) AS EntregadoSA, S.Semana FROM CatSemanas S left join(select* from(select* from(select IdAgen, Cod_Prod, Cod_Campo, SUM(isnull([27], 0)) as [27], sum(isnull([28], 0)) as [28], sum(isnull([29], 0)) as[29], sum(isnull([30], 0)) as [30], sum(isnull([31], 0)) as [31], sum(isnull([32], 0)) as [32], sum(isnull([33], 0)) as [33], sum(isnull([34], 0)) as [34], sum(isnull([35], 0)) as [35], sum(isnull([36], 0)) as [36], sum(isnull([37], 0)) as [37], sum(isnull([38], 0)) as [38], sum(isnull([39], 0)) as [39], sum(isnull([40], 0)) as [40], sum(isnull([41], 0)) as [41], sum(isnull([42], 0)) as [42], sum(isnull([43], 0)) as [43], sum(isnull([44], 0)) as [44], sum(isnull([45], 0)) as [45], sum(isnull([46], 0)) as [46], sum(isnull([47], 0)) as [47], sum(isnull([48], 0)) as [48], sum(isnull([49], 0)) as [49], sum(isnull([50], 0)) as [50], sum(isnull([51], 0)) as [51], sum(isnull([52], 0)) as [52], sum(isnull([1], 0)) as [1], sum(isnull([2], 0)) as [2], sum(isnull([3], 0)) as [3], sum(isnull([4], 0)) as [4], sum(isnull([5], 0)) as [5], sum(isnull([6], 0)) as [6], sum(isnull([7], 0)) as [7], sum(isnull([8], 0)) as [8], sum(isnull([9], 0)) as [9], sum(isnull([10], 0)) as [10], sum(isnull([11], 0)) as [11], sum(isnull([12], 0)) as [12], sum(isnull([13], 0)) as [13], sum(isnull([14], 0)) as [14], sum(isnull([15], 0)) as [15],sum(isnull([16], 0)) as [16], sum(isnull([17], 0)) as [17], sum(isnull([18], 0)) as [18], sum(isnull([19], 0)) as [19], sum(isnull([20], 0)) as [20], sum(isnull([21], 0)) as [21], sum(isnull([22], 0)) as [22], sum(isnull([23], 0)) as [23], sum(isnull([24], 0)) as [24], sum(isnull([25], 0)) as [25], sum(isnull([26], 0)) as [26] FROM SIPGProyeccion P where Temporada = (select Temporada from CatSemanas where getdate() between Inicio and Fin) group by IdAgen,Cod_Prod,Cod_Campo)V UNPIVOT(Pronostico FOR Semana in ([27],[28],[29],[30],[31],[32],[33],[34],[35],[36],[37],[38],[39],[40],[41],[42],[43],[44],[45],[46],[47],[48],[49],[50],[51],[52],[1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],[17],[18],[19],[20],[21],[22],[23],[24],[25],[26]))X)V where Pronostico <> 0)PA ON S.Semana = PA.Semana Left Join(select C.IdAgen, C.Cod_Prod, C.Cod_Campo, sum(C.Convertidas) as Entregado, C.Semana FROM(select IdAgen, Cod_Prod, Cod_Campo, sum(Convertidas) as Convertidas, Semana FROM SEasonsun1213..UV_ProdRecepcion where CodEstatus<> 'C' and Temporada = (select Temporada from CatSemanas where getdate() between Inicio and Fin) group by IdAgen, Cod_prod, Cod_Campo, Semana Union All SELECT IdAgen, Cod_Prod, Cod_Campo, sum(Convertidas) as Convertidas, Semana FROM SeasonPlan..UV_ProdRecepcion where CodEstatus<> 'C' and Temporada = (select Temporada from CatSemanas where getdate() between Inicio and Fin) group by IdAgen, Cod_prod, Cod_Campo, Semana)C GROUP BY C.IdAgen, C.Cod_Prod, C.Cod_Campo, C.Semana)EA ON PA.IdAgen = EA.IdAgen and PA.Cod_Prod = EA.Cod_Prod and PA.Cod_Campo = EA.Cod_Campo and S.Semana = EA.Semana WHERE getdate() between S.Inicio and S.Fin group by PA.IdAgen, S.Semana)SA on V.IdAgen = SA.IdAgen " +
            //    "where V.IdAgen = " + (short)Session["IdAgen"] + " group by V.IdAgen, AA.PronosticoAA, SA.PronosticoSA, SA.EntregadoSA, SA.Semana)V").ToList();
            //return Json(data, JsonRequestBehavior.AllowGet);

            List<ClassCurva> data = bd.Database.SqlQuery<ClassCurva>("Select V.IdAgen, round(V.Entregado,0) as EntregadoT, round(SA.EntregadoSA,0) as EntregadoSA, SA.Semana from(Select E.IdAgen, SUM(isnull(E.Entregado,0)) AS Entregado " +
                "FROM(Select C.IdAgen, C.Cod_Prod, C.Cod_Campo, SUM(C.Convertidas) as Entregado " +
                "FROM(Select IdAgen, Cod_Prod, Cod_Campo, sum(ISNULL(Convertidas, 0)) as Convertidas FROM SEasonsun1213..UV_ProdRecepcion " +
                "where CodEstatus <> 'C' and Temporada = (Select Temporada from CatSemanas where getdate() between Inicio and Fin) group by IdAgen, Cod_prod, Cod_Campo " +
                ")C GROUP BY C.IdAgen, C.Cod_Prod, Cod_Campo)E GROUP BY E.IdAgen)V " +
                "Left Join(Select EA.IdAgen, sum(ISNULL(EA.Entregado, 0)) AS EntregadoSA, S.Semana FROM CatSemanas S " +
                " Left Join(Select C.IdAgen, C.Cod_Prod, C.Cod_Campo, sum(C.Convertidas) as Entregado, C.Semana FROM( " +
                " select IdAgen, Cod_Prod, Cod_Campo, sum(Convertidas) as Convertidas, Semana FROM SEasonsun1213..UV_ProdRecepcion where CodEstatus <> 'C' and " +
                "Temporada = (Select Temporada from CatSemanas where getdate() between Inicio and Fin) group by IdAgen, Cod_prod, Cod_Campo, Semana)C " +
                "GROUP BY C.IdAgen, C.Cod_Prod, C.Cod_Campo, C.Semana)EA ON S.Semana = EA.Semana WHERE getdate() between S.Inicio and S.Fin group by EA.IdAgen, S.Semana)SA on V.IdAgen = SA.IdAgen " +
                "where V.IdAgen = " + (short)Session["IdAgen"] + " group by V.IdAgen, V.Entregado,SA.EntregadoSA, SA.Semana").ToList();
            return Json(data, JsonRequestBehavior.AllowGet);
        }

        //COMENTARIOS
        public ActionResult Comentarios()
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

        //Nombre Productor
        public JsonResult GetProductor(string Cod_Prod)
        {
            bd.Configuration.ProxyCreationEnabled = false;
            ProdProductoresCat nom_prod = bd.ProdProductoresCat.Where(x => x.Cod_Prod == Cod_Prod).FirstOrDefault();
            return Json(nom_prod, JsonRequestBehavior.AllowGet);
        }

        //Enviar comentario
        [HttpPost]
        public ActionResult Comentarios(SIPGComentarios model)
        {
            try
            {
                SIPGComentarios.IdAgen = (short)Session["IdAgen"];
                SIPGComentarios.Fecha = DateTime.Now;
                SIPGComentarios.Cod_prod = model.Cod_prod;
                SIPGComentarios.Comentario = model.Comentario;
                bd.SIPGComentarios.Add(SIPGComentarios);
                bd.SaveChanges();
                //correo
                try
                {
                    MailMessage correo = new MailMessage();
                    correo.From = new MailAddress("indicadores.giddingsfruit@gmail.com", "Indicadores GiddingsFruit");
                    correo.To.Add("marholy.martinez@giddingsfruit.mx");
                    correo.Subject = "Comentario";
                    correo.Body = "Enviado por: " + Session["Nombre"].ToString() + " <br/>";
                    correo.Body += " <br/>";
                    correo.Body += "Productor: " + model.Cod_prod + " <br/>";
                    correo.Body += " <br/>";
                    correo.Body += "Comentario: " + model.Comentario + " <br/>";

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

                TempData["sms"] = "Comentario enviado";
            }
            catch (Exception e)
            {
                TempData["error"] = e.ToString();
                ViewBag.error = TempData["error"].ToString();
            }
            return View();
        }

        //CURVA
        public ActionResult Curva()
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

        public JsonResult CurvaList(int agente = 0, string filtro = "")
        {
            List<ClassCurva> curva = new List<ClassCurva>();

            if ((short)Session["IdAgen"] != 0)
            {
                if (filtro == "cinco")
                {
                    curva = bd.Database.SqlQuery<ClassCurva>("select C.Cod_Prod, C.Nombre as Productor, C.PronosticoT, C.PronosticoA, C.EntregadoT, (CASE WHEN C.EntregadoT ='0' AND C.PronosticoA ='0' THEN '' ELSE C.DiferenciaA END) AS DiferenciaA, C.Semana, C.PronosticoSA, C.EntregadoSA, (CASE WHEN C.EntregadoSA ='0' AND  C.PronosticoSA ='0' THEN '' ELSE C.DiferenciaSA END) AS DiferenciaSA from(SELECT distinct V.IdAgen, V.Cod_Prod, V.Nombre, (case when round(V.PronosticoT,0) is null then '0' else round(V.PronosticoT,0) end) as PronosticoT, (case when round(V.PronosticoA,0) is null then '0' else round(V.PronosticoA,0) end) as PronosticoA, (case when round(V.EntregadoT,0) is null then '0' else round(V.EntregadoT,0) end)as EntregadoT, (case when round(((V.DiferenciaA*100)-100),0) is null then '100' else round(((V.DiferenciaA*100)-100),0) end)as DiferenciaA, V.Semana, (case when round(V.PronosticoSA,0) is null then '0' else round(V.PronosticoSA,0) end) as PronosticoSA, (case when round(V.EntregadoSA,0) is null then '0' else round(V.EntregadoSA,0) end) as EntregadoSA, (case when round(((V.DiferenciaSA*100)-100),0) is null then '100' else round(((V.DiferenciaSA*100)-100),0) end) as DiferenciaSA FROM(select *,(V.EntregadoT/V.PronosticoA) AS DiferenciaA, (V.EntregadoSA / V.PronosticoSA) AS DiferenciaSA FROM(select C.IdAgen, C.Cod_Prod, P.Nombre, V.PronosticoT, V.EntregadoT, AA.PronosticoA, SA.Semana, SA.PronosticoSA, SA.EntregadoSA FROM ProdCamposCat C left join ProdProductoresCat P on C.Cod_Prod=P.Cod_Prod left join(select V.IdAgen, V.Cod_prod, sum(ISNULL(V.Pronostico,0)) as PronosticoT , sum(ISNULL(V.EntregadoT,0)) AS EntregadoT from(select V.IdAgen, V.Cod_prod, sum(V.Pronostico) as Pronostico, E.EntregadoT from(select * from(select V.IdAgen, V.Cod_Prod,max(V.Fecha)as fecha, V.[27],V.[28],V.[29],V.[30],V.[31],V.[32],V.[33],V.[34],V.[35],V.[36],V.[37],V.[38],V.[39],V.[40],V.[41],V.[42],V.[43],V.[44],V.[45],V.[46],V.[47],V.[48],V.[49],V.[50],V.[51],V.[52],V.[53],V.[1],V.[2],V.[3],V.[4],V.[5],V.[6],V.[7],V.[8],V.[9],V.[10],V.[11],V.[12],V.[13],V.[14],V.[15],V.[16],V.[17],V.[18],V.[19],V.[20],V.[21],V.[22],V.[23],V.[24],V.[25],V.[26] From(select distinct IdAgen, Cod_Prod, Fecha, SUM(isnull([27], 0)) as [27], sum(isnull([28], 0)) as [28], sum(isnull([29], 0)) as[29], sum(isnull([30], 0)) as [30], sum(isnull([31], 0)) as [31], sum(isnull([32], 0)) as [32], sum(isnull([33], 0)) as [33], sum(isnull([34], 0)) as [34], sum(isnull([35], 0)) as [35], sum(isnull([36], 0)) as [36], sum(isnull([37], 0)) as [37], sum(isnull([38], 0)) as [38], sum(isnull([39], 0)) as [39], sum(isnull([40], 0)) as [40], sum(isnull([41], 0)) as [41], sum(isnull([42], 0)) as [42], sum(isnull([43], 0)) as [43], sum(isnull([44], 0)) as [44], sum(isnull([45], 0)) as [45], sum(isnull([46], 0)) as [46], sum(isnull([47], 0)) as [47], sum(isnull([48], 0)) as [48], sum(isnull([49], 0)) as [49], sum(isnull([50], 0)) as [50], sum(isnull([51], 0)) as [51], sum(isnull([52], 0)) as [52], sum(isnull([53], 0)) as [53], sum(isnull([1], 0)) as [1], sum(isnull([2], 0)) as [2], sum(isnull([3], 0)) as [3], sum(isnull([4], 0)) as [4], sum(isnull([5], 0)) as [5], sum(isnull([6], 0)) as [6], sum(isnull([7], 0)) as [7], sum(isnull([8], 0)) as [8], sum(isnull([9], 0)) as [9], sum(isnull([10], 0)) as [10], sum(isnull([11], 0)) as [11], sum(isnull([12], 0)) as [12], sum(isnull([13], 0)) as [13], sum(isnull([14], 0)) as [14], sum(isnull([15], 0)) as [15],sum(isnull([16], 0)) as [16], sum(isnull([17], 0)) as [17], sum(isnull([18], 0)) as [18], sum(isnull([19], 0)) as [19], sum(isnull([20], 0)) as [20], sum(isnull([21], 0)) as [21], sum(isnull([22], 0)) as [22], sum(isnull([23], 0)) as [23], sum(isnull([24], 0)) as [24], sum(isnull([25], 0)) as [25], sum(isnull([26], 0)) as [26] from SIPGProyeccion P where Temporada = (select Temporada from CatSemanas where getdate() between Inicio and Fin) and  " +
                           "Fecha = (select max(Fecha) from SIPGProyeccion where IdAgen = " + (short)Session["IdAgen"] + ") group by IdAgen, Cod_Prod,Fecha)V GROUP BY V.IdAgen, V.Cod_Prod,V.[27],V.[28],V.[29],V.[30],V.[31],V.[32],V.[33],V.[34],V.[35],V.[36],V.[37],V.[38],V.[39],V.[40],V.[41],V.[42],V.[43],V.[44],V.[45],V.[46],V.[47],V.[48],V.[49],V.[50],V.[51],V.[52],V.[53],V.[1],V.[2],V.[3],V.[4],V.[5],V.[6],V.[7],V.[8],V.[9],V.[10],V.[11],V.[12],V.[13],V.[14],V.[15],V.[16],V.[17],V.[18],V.[19],V.[20],V.[21],V.[22],V.[23],V.[24],V.[25],V.[26])V UNPIVOT(Pronostico FOR Semana in ([27],[28],[29],[30],[31],[32],[33],[34],[35],[36],[37],[38],[39],[40],[41],[42],[43],[44],[45],[46],[47],[48],[49],[50],[51],[52],[53],[1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],[17],[18],[19],[20],[21],[22],[23],[24],[25],[26]))X Where Pronostico <> 0)V left join(select C.IdAgen, C.Cod_Prod, sum(C.Convertidas) as EntregadoT FROM(select IdAgen, Cod_Prod, Cod_Campo, sum(Convertidas) as Convertidas FROM UV_ProdRecepcion where CodEstatus <> 'C' and Temporada = (select Temporada from CatSemanas where getdate() between Inicio and Fin) group by IdAgen, Cod_prod, Cod_Campo)C GROUP BY C.IdAgen, C.Cod_Prod)E ON V.IdAgen = E.IdAgen AND V.Cod_Prod = E.Cod_Prod group by V.IdAgen, V.Cod_prod, E.EntregadoT)V GROUP BY V.IdAgen, V.Cod_prod)V ON C.IdAgen = V.IdAgen and C.Cod_Prod = V.Cod_Prod left join(select V.IdAgen, V.Cod_prod, sum(isnull(V.Pronostico,0)) as PronosticoA FROM CatSemanas S left join(select * from(select V.IdAgen, V.Cod_Prod, max(V.Fecha) as fecha, V.[27], V.[28], V.[29], V.[30], V.[31], V.[32], V.[33], V.[34], V.[35], V.[36], V.[37], V.[38], V.[39], V.[40], V.[41], V.[42], V.[43], V.[44], V.[45], V.[46], V.[47], V.[48], V.[49], V.[50], V.[51], V.[52], V.[53], V.[1], V.[2], V.[3], V.[4], V.[5], V.[6], V.[7], V.[8], V.[9], V.[10], V.[11], V.[12], V.[13], V.[14], V.[15], V.[16], V.[17], V.[18], V.[19], V.[20], V.[21], V.[22], V.[23], V.[24], V.[25], V.[26] From(select distinct IdAgen, Cod_Prod, Fecha, SUM(isnull([27], 0)) as [27], sum(isnull([28], 0)) as [28], sum(isnull([29], 0)) as[29], sum(isnull([30], 0)) as [30], sum(isnull([31], 0)) as [31], sum(isnull([32], 0)) as [32], sum(isnull([33], 0)) as [33], sum(isnull([34], 0)) as [34], sum(isnull([35], 0)) as [35], sum(isnull([36], 0)) as [36], sum(isnull([37], 0)) as [37], sum(isnull([38], 0)) as [38], sum(isnull([39], 0)) as [39], sum(isnull([40], 0)) as [40], sum(isnull([41], 0)) as [41], sum(isnull([42], 0)) as [42], sum(isnull([43], 0)) as [43], sum(isnull([44], 0)) as [44], sum(isnull([45], 0)) as [45], sum(isnull([46], 0)) as [46], sum(isnull([47], 0)) as [47], sum(isnull([48], 0)) as [48], sum(isnull([49], 0)) as [49], sum(isnull([50], 0)) as [50], sum(isnull([51], 0)) as [51], sum(isnull([52], 0)) as [52],sum(isnull([53], 0)) as [53], sum(isnull([1], 0)) as [1], sum(isnull([2], 0)) as [2], sum(isnull([3], 0)) as [3], sum(isnull([4], 0)) as [4], sum(isnull([5], 0)) as [5], sum(isnull([6], 0)) as [6], sum(isnull([7], 0)) as [7], sum(isnull([8], 0)) as [8], sum(isnull([9], 0)) as [9], sum(isnull([10], 0)) as [10], sum(isnull([11], 0)) as [11], sum(isnull([12], 0)) as [12], sum(isnull([13], 0)) as [13], sum(isnull([14], 0)) as [14], sum(isnull([15], 0)) as [15],sum(isnull([16], 0)) as [16], sum(isnull([17], 0)) as [17], sum(isnull([18], 0)) as [18], sum(isnull([19], 0)) as [19], sum(isnull([20], 0)) as [20], sum(isnull([21], 0)) as [21], sum(isnull([22], 0)) as [22], sum(isnull([23], 0)) as [23], sum(isnull([24], 0)) as [24], sum(isnull([25], 0)) as [25], sum(isnull([26], 0)) as [26] from SIPGProyeccion P where Temporada = (select Temporada from CatSemanas where getdate() between Inicio and Fin) " +
                           "and Fecha = (select max(Fecha) from SIPGProyeccion where IdAgen = " + (short)Session["IdAgen"] + ") group by IdAgen, Cod_Prod,Fecha)V GROUP BY V.IdAgen, V.Cod_Prod,V.[27],V.[28],V.[29],V.[30],V.[31],V.[32],V.[33],V.[34],V.[35],V.[36],V.[37],V.[38],V.[39],V.[40],V.[41],V.[42],V.[43],V.[44],V.[45],V.[46],V.[47],V.[48],V.[49],V.[50],V.[51],V.[52],V.[53],V.[1],V.[2],V.[3],V.[4],V.[5],V.[6],V.[7],V.[8],V.[9],V.[10],V.[11],V.[12],V.[13],V.[14],V.[15],V.[16],V.[17],V.[18],V.[19],V.[20],V.[21],V.[22],V.[23],V.[24],V.[25],V.[26])V UNPIVOT(Pronostico FOR Semana in ([27],[28],[29],[30],[31],[32],[33],[34],[35],[36],[37],[38],[39],[40],[41],[42],[43],[44],[45],[46],[47],[48],[49],[50],[51],[52],[53],[1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],[17],[18],[19],[20],[21],[22],[23],[24],[25],[26]))X where Pronostico <> 0)V on S.Semana = V.Semana where S.Inicio between(select Inicio from catsemanas where temporada = (select Temporada from CatSemanas where getdate() between Inicio and Fin) and semana = 27) and getdate() group by V.IdAgen,V.Cod_prod)AA ON V.IdAgen = AA.IdAgen and V.Cod_Prod = AA.Cod_Prod LEFT JOIN(SELECT PA.IdAgen, PA.Cod_Prod, SUM(isnull(PA.Pronostico,0)) AS PronosticoSA, sum(isnull(EA.Entregado, 0)) AS EntregadoSA, S.Semana FROM CatSemanas S left join(select* from(select V.IdAgen, V.Cod_Prod, max(V.Fecha)as fecha, V.[27],V.[28],V.[29],V.[30],V.[31],V.[32],V.[33],V.[34],V.[35],V.[36],V.[37],V.[38],V.[39],V.[40],V.[41],V.[42],V.[43],V.[44],V.[45],V.[46],V.[47],V.[48],V.[49],V.[50],V.[51],V.[52],V.[53],V.[1],V.[2],V.[3],V.[4],V.[5],V.[6],V.[7],V.[8],V.[9],V.[10],V.[11],V.[12],V.[13],V.[14],V.[15],V.[16],V.[17],V.[18],V.[19],V.[20],V.[21],V.[22],V.[23],V.[24],V.[25],V.[26] From(select distinct IdAgen, Cod_Prod, Fecha, SUM(isnull([27], 0)) as [27], sum(isnull([28], 0)) as [28], sum(isnull([29], 0)) as[29], sum(isnull([30], 0)) as [30], sum(isnull([31], 0)) as [31], sum(isnull([32], 0)) as [32], sum(isnull([33], 0)) as [33], sum(isnull([34], 0)) as [34], sum(isnull([35], 0)) as [35], sum(isnull([36], 0)) as [36], sum(isnull([37], 0)) as [37], sum(isnull([38], 0)) as [38], sum(isnull([39], 0)) as [39], sum(isnull([40], 0)) as [40], sum(isnull([41], 0)) as [41], sum(isnull([42], 0)) as [42], sum(isnull([43], 0)) as [43], sum(isnull([44], 0)) as [44], sum(isnull([45], 0)) as [45], sum(isnull([46], 0)) as [46], sum(isnull([47], 0)) as [47], sum(isnull([48], 0)) as [48], sum(isnull([49], 0)) as [49], sum(isnull([50], 0)) as [50], sum(isnull([51], 0)) as [51], sum(isnull([52], 0)) as [52],sum(isnull([53], 0)) as [53], sum(isnull([1], 0)) as [1], sum(isnull([2], 0)) as [2], sum(isnull([3], 0)) as [3], sum(isnull([4], 0)) as [4], sum(isnull([5], 0)) as [5], sum(isnull([6], 0)) as [6], sum(isnull([7], 0)) as [7], sum(isnull([8], 0)) as [8], sum(isnull([9], 0)) as [9], sum(isnull([10], 0)) as [10], sum(isnull([11], 0)) as [11], sum(isnull([12], 0)) as [12], sum(isnull([13], 0)) as [13], sum(isnull([14], 0)) as [14], sum(isnull([15], 0)) as [15],sum(isnull([16], 0)) as [16], sum(isnull([17], 0)) as [17], sum(isnull([18], 0)) as [18], sum(isnull([19], 0)) as [19], sum(isnull([20], 0)) as [20], sum(isnull([21], 0)) as [21], sum(isnull([22], 0)) as [22], sum(isnull([23], 0)) as [23], sum(isnull([24], 0)) as [24], sum(isnull([25], 0)) as [25], sum(isnull([26], 0)) as [26] from SIPGProyeccion P where Temporada = (select Temporada from CatSemanas where getdate() between Inicio and Fin) " +
                           "and Fecha = (select max(Fecha) from SIPGProyeccion where IdAgen = " + (short)Session["IdAgen"] + ") group by IdAgen, Cod_Prod,Fecha)V GROUP BY V.IdAgen, V.Cod_Prod,V.[27],V.[28],V.[29],V.[30],V.[31],V.[32],V.[33],V.[34],V.[35],V.[36],V.[37],V.[38],V.[39],V.[40],V.[41],V.[42],V.[43],V.[44],V.[45],V.[46],V.[47],V.[48],V.[49],V.[50],V.[51],V.[52],V.[53],V.[1],V.[2],V.[3],V.[4],V.[5],V.[6],V.[7],V.[8],V.[9],V.[10],V.[11],V.[12],V.[13],V.[14],V.[15],V.[16],V.[17],V.[18],V.[19],V.[20],V.[21],V.[22],V.[23],V.[24],V.[25],V.[26])V UNPIVOT(Pronostico FOR Semana in ([27],[28],[29],[30],[31],[32],[33],[34],[35],[36],[37],[38],[39],[40],[41],[42],[43],[44],[45],[46],[47],[48],[49],[50],[51],[52],[53],[1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],[17],[18],[19],[20],[21],[22],[23],[24],[25],[26]))X where Pronostico <> 0)PA ON S.Semana = PA.Semana Left Join(select C.IdAgen, C.Cod_Prod, sum(C.Convertidas) as Entregado, C.Semana FROM(select IdAgen, Cod_Prod, sum(Convertidas) as Convertidas, Semana FROM UV_ProdRecepcion where CodEstatus <> 'C' and Temporada = (select Temporada from CatSemanas where getdate() between Inicio and Fin) group by IdAgen, Cod_prod, Cod_Campo, Semana)C GROUP BY C.IdAgen, C.Cod_Prod, C.Semana)EA ON PA.IdAgen = EA.IdAgen and PA.Cod_Prod = EA.Cod_Prod and S.Semana = EA.Semana WHERE getdate() between S.Inicio and S.Fin GROUP BY PA.IdAgen, PA.Cod_Prod, S.Semana)SA on V.IdAgen = SA.IdAgen AND V.Cod_Prod = SA.Cod_Prod group by C.IdAgen, C.Cod_Prod, P.Nombre, V.PronosticoT, V.EntregadoT, AA.PronosticoA, SA.Semana, SA.PronosticoSA, SA.EntregadoSA)V)V)C " +
                           "WHERE C.IdAgen = " + (short)Session["IdAgen"] + "  and C.Cod_Prod not in(Select distinct V.Cod_Prod FROM(select* from CatSemanas where Getdate() between Inicio and Fin)S left join SIPGVisitas V on V.Fecha between S.Inicio and S.Fin " +
                           "where V.IdAgen = " + (short)Session["IdAgen"] + ") and C.DiferenciaA > 0 and C.DiferenciaA <= 5 or C.DiferenciaA < 0 and C.DiferenciaA > -5 ORDER BY C.DiferenciaA DESC").ToList();
                }
                else if (filtro == "vcinco")
                {
                    curva = bd.Database.SqlQuery<ClassCurva>("select C.Cod_Prod, C.Nombre as Productor, C.PronosticoT, C.PronosticoA, C.EntregadoT, (CASE WHEN C.EntregadoT ='0' AND C.PronosticoA ='0' THEN '' ELSE C.DiferenciaA END) AS DiferenciaA, C.Semana, C.PronosticoSA, C.EntregadoSA, (CASE WHEN C.EntregadoSA ='0' AND  C.PronosticoSA ='0' THEN '' ELSE C.DiferenciaSA END) AS DiferenciaSA from(SELECT distinct V.IdAgen, V.Cod_Prod, V.Nombre, (case when round(V.PronosticoT,0) is null then '0' else round(V.PronosticoT,0) end) as PronosticoT, (case when round(V.PronosticoA,0) is null then '0' else round(V.PronosticoA,0) end) as PronosticoA, (case when round(V.EntregadoT,0) is null then '0' else round(V.EntregadoT,0) end)as EntregadoT, (case when round(((V.DiferenciaA*100)-100),0) is null then '100' else round(((V.DiferenciaA*100)-100),0) end)as DiferenciaA, V.Semana, (case when round(V.PronosticoSA,0) is null then '0' else round(V.PronosticoSA,0) end) as PronosticoSA, (case when round(V.EntregadoSA,0) is null then '0' else round(V.EntregadoSA,0) end) as EntregadoSA, (case when round(((V.DiferenciaSA*100)-100),0) is null then '100' else round(((V.DiferenciaSA*100)-100),0) end) as DiferenciaSA FROM(select *,(V.EntregadoT/V.PronosticoA) AS DiferenciaA, (V.EntregadoSA / V.PronosticoSA) AS DiferenciaSA FROM(select C.IdAgen, C.Cod_Prod, P.Nombre, V.PronosticoT, V.EntregadoT, AA.PronosticoA, SA.Semana, SA.PronosticoSA, SA.EntregadoSA FROM ProdCamposCat C left join ProdProductoresCat P on C.Cod_Prod=P.Cod_Prod left join(select V.IdAgen, V.Cod_prod, sum(ISNULL(V.Pronostico,0)) as PronosticoT , sum(ISNULL(V.EntregadoT,0)) AS EntregadoT from(select V.IdAgen, V.Cod_prod, sum(V.Pronostico) as Pronostico, E.EntregadoT from(select * from(select V.IdAgen, V.Cod_Prod,max(V.Fecha)as fecha, V.[27],V.[28],V.[29],V.[30],V.[31],V.[32],V.[33],V.[34],V.[35],V.[36],V.[37],V.[38],V.[39],V.[40],V.[41],V.[42],V.[43],V.[44],V.[45],V.[46],V.[47],V.[48],V.[49],V.[50],V.[51],V.[52],V.[53],V.[1],V.[2],V.[3],V.[4],V.[5],V.[6],V.[7],V.[8],V.[9],V.[10],V.[11],V.[12],V.[13],V.[14],V.[15],V.[16],V.[17],V.[18],V.[19],V.[20],V.[21],V.[22],V.[23],V.[24],V.[25],V.[26] From(select distinct IdAgen, Cod_Prod, Fecha, SUM(isnull([27], 0)) as [27], sum(isnull([28], 0)) as [28], sum(isnull([29], 0)) as[29], sum(isnull([30], 0)) as [30], sum(isnull([31], 0)) as [31], sum(isnull([32], 0)) as [32], sum(isnull([33], 0)) as [33], sum(isnull([34], 0)) as [34], sum(isnull([35], 0)) as [35], sum(isnull([36], 0)) as [36], sum(isnull([37], 0)) as [37], sum(isnull([38], 0)) as [38], sum(isnull([39], 0)) as [39], sum(isnull([40], 0)) as [40], sum(isnull([41], 0)) as [41], sum(isnull([42], 0)) as [42], sum(isnull([43], 0)) as [43], sum(isnull([44], 0)) as [44], sum(isnull([45], 0)) as [45], sum(isnull([46], 0)) as [46], sum(isnull([47], 0)) as [47], sum(isnull([48], 0)) as [48], sum(isnull([49], 0)) as [49], sum(isnull([50], 0)) as [50], sum(isnull([51], 0)) as [51], sum(isnull([52], 0)) as [52], sum(isnull([53], 0)) as [53], sum(isnull([1], 0)) as [1], sum(isnull([2], 0)) as [2], sum(isnull([3], 0)) as [3], sum(isnull([4], 0)) as [4], sum(isnull([5], 0)) as [5], sum(isnull([6], 0)) as [6], sum(isnull([7], 0)) as [7], sum(isnull([8], 0)) as [8], sum(isnull([9], 0)) as [9], sum(isnull([10], 0)) as [10], sum(isnull([11], 0)) as [11], sum(isnull([12], 0)) as [12], sum(isnull([13], 0)) as [13], sum(isnull([14], 0)) as [14], sum(isnull([15], 0)) as [15],sum(isnull([16], 0)) as [16], sum(isnull([17], 0)) as [17], sum(isnull([18], 0)) as [18], sum(isnull([19], 0)) as [19], sum(isnull([20], 0)) as [20], sum(isnull([21], 0)) as [21], sum(isnull([22], 0)) as [22], sum(isnull([23], 0)) as [23], sum(isnull([24], 0)) as [24], sum(isnull([25], 0)) as [25], sum(isnull([26], 0)) as [26] from SIPGProyeccion P where Temporada = (select Temporada from CatSemanas where getdate() between Inicio and Fin) and  " +
                       "Fecha = (select max(Fecha) from SIPGProyeccion where IdAgen = " + (short)Session["IdAgen"] + ") group by IdAgen, Cod_Prod,Fecha)V GROUP BY V.IdAgen, V.Cod_Prod,V.[27],V.[28],V.[29],V.[30],V.[31],V.[32],V.[33],V.[34],V.[35],V.[36],V.[37],V.[38],V.[39],V.[40],V.[41],V.[42],V.[43],V.[44],V.[45],V.[46],V.[47],V.[48],V.[49],V.[50],V.[51],V.[52],V.[53],V.[1],V.[2],V.[3],V.[4],V.[5],V.[6],V.[7],V.[8],V.[9],V.[10],V.[11],V.[12],V.[13],V.[14],V.[15],V.[16],V.[17],V.[18],V.[19],V.[20],V.[21],V.[22],V.[23],V.[24],V.[25],V.[26])V UNPIVOT(Pronostico FOR Semana in ([27],[28],[29],[30],[31],[32],[33],[34],[35],[36],[37],[38],[39],[40],[41],[42],[43],[44],[45],[46],[47],[48],[49],[50],[51],[52],[53],[1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],[17],[18],[19],[20],[21],[22],[23],[24],[25],[26]))X Where Pronostico <> 0)V left join(select C.IdAgen, C.Cod_Prod, sum(C.Convertidas) as EntregadoT FROM(select IdAgen, Cod_Prod, Cod_Campo, sum(Convertidas) as Convertidas FROM UV_ProdRecepcion where CodEstatus <> 'C' and Temporada = (select Temporada from CatSemanas where getdate() between Inicio and Fin) group by IdAgen, Cod_prod, Cod_Campo)C GROUP BY C.IdAgen, C.Cod_Prod)E ON V.IdAgen = E.IdAgen AND V.Cod_Prod = E.Cod_Prod group by V.IdAgen, V.Cod_prod, E.EntregadoT)V GROUP BY V.IdAgen, V.Cod_prod)V ON C.IdAgen = V.IdAgen and C.Cod_Prod = V.Cod_Prod left join(select V.IdAgen, V.Cod_prod, sum(isnull(V.Pronostico,0)) as PronosticoA FROM CatSemanas S left join(select * from(select V.IdAgen, V.Cod_Prod, max(V.Fecha) as fecha, V.[27], V.[28], V.[29], V.[30], V.[31], V.[32], V.[33], V.[34], V.[35], V.[36], V.[37], V.[38], V.[39], V.[40], V.[41], V.[42], V.[43], V.[44], V.[45], V.[46], V.[47], V.[48], V.[49], V.[50], V.[51], V.[52], V.[53], V.[1], V.[2], V.[3], V.[4], V.[5], V.[6], V.[7], V.[8], V.[9], V.[10], V.[11], V.[12], V.[13], V.[14], V.[15], V.[16], V.[17], V.[18], V.[19], V.[20], V.[21], V.[22], V.[23], V.[24], V.[25], V.[26] From(select distinct IdAgen, Cod_Prod, Fecha, SUM(isnull([27], 0)) as [27], sum(isnull([28], 0)) as [28], sum(isnull([29], 0)) as[29], sum(isnull([30], 0)) as [30], sum(isnull([31], 0)) as [31], sum(isnull([32], 0)) as [32], sum(isnull([33], 0)) as [33], sum(isnull([34], 0)) as [34], sum(isnull([35], 0)) as [35], sum(isnull([36], 0)) as [36], sum(isnull([37], 0)) as [37], sum(isnull([38], 0)) as [38], sum(isnull([39], 0)) as [39], sum(isnull([40], 0)) as [40], sum(isnull([41], 0)) as [41], sum(isnull([42], 0)) as [42], sum(isnull([43], 0)) as [43], sum(isnull([44], 0)) as [44], sum(isnull([45], 0)) as [45], sum(isnull([46], 0)) as [46], sum(isnull([47], 0)) as [47], sum(isnull([48], 0)) as [48], sum(isnull([49], 0)) as [49], sum(isnull([50], 0)) as [50], sum(isnull([51], 0)) as [51], sum(isnull([52], 0)) as [52],sum(isnull([53], 0)) as [53], sum(isnull([1], 0)) as [1], sum(isnull([2], 0)) as [2], sum(isnull([3], 0)) as [3], sum(isnull([4], 0)) as [4], sum(isnull([5], 0)) as [5], sum(isnull([6], 0)) as [6], sum(isnull([7], 0)) as [7], sum(isnull([8], 0)) as [8], sum(isnull([9], 0)) as [9], sum(isnull([10], 0)) as [10], sum(isnull([11], 0)) as [11], sum(isnull([12], 0)) as [12], sum(isnull([13], 0)) as [13], sum(isnull([14], 0)) as [14], sum(isnull([15], 0)) as [15],sum(isnull([16], 0)) as [16], sum(isnull([17], 0)) as [17], sum(isnull([18], 0)) as [18], sum(isnull([19], 0)) as [19], sum(isnull([20], 0)) as [20], sum(isnull([21], 0)) as [21], sum(isnull([22], 0)) as [22], sum(isnull([23], 0)) as [23], sum(isnull([24], 0)) as [24], sum(isnull([25], 0)) as [25], sum(isnull([26], 0)) as [26] from SIPGProyeccion P where Temporada = (select Temporada from CatSemanas where getdate() between Inicio and Fin) " +
                       "and Fecha = (select max(Fecha) from SIPGProyeccion where IdAgen = " + (short)Session["IdAgen"] + ") group by IdAgen, Cod_Prod,Fecha)V GROUP BY V.IdAgen, V.Cod_Prod,V.[27],V.[28],V.[29],V.[30],V.[31],V.[32],V.[33],V.[34],V.[35],V.[36],V.[37],V.[38],V.[39],V.[40],V.[41],V.[42],V.[43],V.[44],V.[45],V.[46],V.[47],V.[48],V.[49],V.[50],V.[51],V.[52],V.[53],V.[1],V.[2],V.[3],V.[4],V.[5],V.[6],V.[7],V.[8],V.[9],V.[10],V.[11],V.[12],V.[13],V.[14],V.[15],V.[16],V.[17],V.[18],V.[19],V.[20],V.[21],V.[22],V.[23],V.[24],V.[25],V.[26])V UNPIVOT(Pronostico FOR Semana in ([27],[28],[29],[30],[31],[32],[33],[34],[35],[36],[37],[38],[39],[40],[41],[42],[43],[44],[45],[46],[47],[48],[49],[50],[51],[52],[53],[1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],[17],[18],[19],[20],[21],[22],[23],[24],[25],[26]))X where Pronostico <> 0)V on S.Semana = V.Semana where S.Inicio between(select Inicio from catsemanas where temporada = (select Temporada from CatSemanas where getdate() between Inicio and Fin) and semana = 27) and getdate() group by V.IdAgen,V.Cod_prod)AA ON V.IdAgen = AA.IdAgen and V.Cod_Prod = AA.Cod_Prod LEFT JOIN(SELECT PA.IdAgen, PA.Cod_Prod, SUM(isnull(PA.Pronostico,0)) AS PronosticoSA, sum(isnull(EA.Entregado, 0)) AS EntregadoSA, S.Semana FROM CatSemanas S left join(select* from(select V.IdAgen, V.Cod_Prod, max(V.Fecha)as fecha, V.[27],V.[28],V.[29],V.[30],V.[31],V.[32],V.[33],V.[34],V.[35],V.[36],V.[37],V.[38],V.[39],V.[40],V.[41],V.[42],V.[43],V.[44],V.[45],V.[46],V.[47],V.[48],V.[49],V.[50],V.[51],V.[52],V.[53],V.[1],V.[2],V.[3],V.[4],V.[5],V.[6],V.[7],V.[8],V.[9],V.[10],V.[11],V.[12],V.[13],V.[14],V.[15],V.[16],V.[17],V.[18],V.[19],V.[20],V.[21],V.[22],V.[23],V.[24],V.[25],V.[26] From(select distinct IdAgen, Cod_Prod, Fecha, SUM(isnull([27], 0)) as [27], sum(isnull([28], 0)) as [28], sum(isnull([29], 0)) as[29], sum(isnull([30], 0)) as [30], sum(isnull([31], 0)) as [31], sum(isnull([32], 0)) as [32], sum(isnull([33], 0)) as [33], sum(isnull([34], 0)) as [34], sum(isnull([35], 0)) as [35], sum(isnull([36], 0)) as [36], sum(isnull([37], 0)) as [37], sum(isnull([38], 0)) as [38], sum(isnull([39], 0)) as [39], sum(isnull([40], 0)) as [40], sum(isnull([41], 0)) as [41], sum(isnull([42], 0)) as [42], sum(isnull([43], 0)) as [43], sum(isnull([44], 0)) as [44], sum(isnull([45], 0)) as [45], sum(isnull([46], 0)) as [46], sum(isnull([47], 0)) as [47], sum(isnull([48], 0)) as [48], sum(isnull([49], 0)) as [49], sum(isnull([50], 0)) as [50], sum(isnull([51], 0)) as [51], sum(isnull([52], 0)) as [52],sum(isnull([53], 0)) as [53], sum(isnull([1], 0)) as [1], sum(isnull([2], 0)) as [2], sum(isnull([3], 0)) as [3], sum(isnull([4], 0)) as [4], sum(isnull([5], 0)) as [5], sum(isnull([6], 0)) as [6], sum(isnull([7], 0)) as [7], sum(isnull([8], 0)) as [8], sum(isnull([9], 0)) as [9], sum(isnull([10], 0)) as [10], sum(isnull([11], 0)) as [11], sum(isnull([12], 0)) as [12], sum(isnull([13], 0)) as [13], sum(isnull([14], 0)) as [14], sum(isnull([15], 0)) as [15],sum(isnull([16], 0)) as [16], sum(isnull([17], 0)) as [17], sum(isnull([18], 0)) as [18], sum(isnull([19], 0)) as [19], sum(isnull([20], 0)) as [20], sum(isnull([21], 0)) as [21], sum(isnull([22], 0)) as [22], sum(isnull([23], 0)) as [23], sum(isnull([24], 0)) as [24], sum(isnull([25], 0)) as [25], sum(isnull([26], 0)) as [26] from SIPGProyeccion P where Temporada = (select Temporada from CatSemanas where getdate() between Inicio and Fin) " +
                       "and Fecha = (select max(Fecha) from SIPGProyeccion where IdAgen = " + (short)Session["IdAgen"] + ") group by IdAgen, Cod_Prod,Fecha)V GROUP BY V.IdAgen, V.Cod_Prod,V.[27],V.[28],V.[29],V.[30],V.[31],V.[32],V.[33],V.[34],V.[35],V.[36],V.[37],V.[38],V.[39],V.[40],V.[41],V.[42],V.[43],V.[44],V.[45],V.[46],V.[47],V.[48],V.[49],V.[50],V.[51],V.[52],V.[53],V.[1],V.[2],V.[3],V.[4],V.[5],V.[6],V.[7],V.[8],V.[9],V.[10],V.[11],V.[12],V.[13],V.[14],V.[15],V.[16],V.[17],V.[18],V.[19],V.[20],V.[21],V.[22],V.[23],V.[24],V.[25],V.[26])V UNPIVOT(Pronostico FOR Semana in ([27],[28],[29],[30],[31],[32],[33],[34],[35],[36],[37],[38],[39],[40],[41],[42],[43],[44],[45],[46],[47],[48],[49],[50],[51],[52],[53],[1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],[17],[18],[19],[20],[21],[22],[23],[24],[25],[26]))X where Pronostico <> 0)PA ON S.Semana = PA.Semana Left Join(select C.IdAgen, C.Cod_Prod, sum(C.Convertidas) as Entregado, C.Semana FROM(select IdAgen, Cod_Prod, sum(Convertidas) as Convertidas, Semana FROM UV_ProdRecepcion where CodEstatus <> 'C' and Temporada = (select Temporada from CatSemanas where getdate() between Inicio and Fin) group by IdAgen, Cod_prod, Cod_Campo, Semana)C GROUP BY C.IdAgen, C.Cod_Prod, C.Semana)EA ON PA.IdAgen = EA.IdAgen and PA.Cod_Prod = EA.Cod_Prod and S.Semana = EA.Semana WHERE getdate() between S.Inicio and S.Fin GROUP BY PA.IdAgen, PA.Cod_Prod, S.Semana)SA on V.IdAgen = SA.IdAgen AND V.Cod_Prod = SA.Cod_Prod group by C.IdAgen, C.Cod_Prod, P.Nombre, V.PronosticoT, V.EntregadoT, AA.PronosticoA, SA.Semana, SA.PronosticoSA, SA.EntregadoSA)V)V)C " +
                       "WHERE C.IdAgen = " + (short)Session["IdAgen"] + "  and C.Cod_Prod not in(Select distinct V.Cod_Prod FROM(select* from CatSemanas where Getdate() between Inicio and Fin)S left join SIPGVisitas V on V.Fecha between S.Inicio and S.Fin " +
                       "where V.IdAgen = " + (short)Session["IdAgen"] + ") and C.DiferenciaA > 5 and C.DiferenciaA <=25 or C.DiferenciaA < -5 and C.DiferenciaA >= -25 ORDER BY C.DiferenciaA DESC").ToList();
                }
                else if (filtro == "cincuenta")
                {
                    curva = bd.Database.SqlQuery<ClassCurva>("select C.Cod_Prod, C.Nombre as Productor, C.PronosticoT, C.PronosticoA, C.EntregadoT, (CASE WHEN C.EntregadoT ='0' AND C.PronosticoA ='0' THEN '' ELSE C.DiferenciaA END) AS DiferenciaA, C.Semana, C.PronosticoSA, C.EntregadoSA, (CASE WHEN C.EntregadoSA ='0' AND  C.PronosticoSA ='0' THEN '' ELSE C.DiferenciaSA END) AS DiferenciaSA from(SELECT distinct V.IdAgen, V.Cod_Prod, V.Nombre, (case when round(V.PronosticoT,0) is null then '0' else round(V.PronosticoT,0) end) as PronosticoT, (case when round(V.PronosticoA,0) is null then '0' else round(V.PronosticoA,0) end) as PronosticoA, (case when round(V.EntregadoT,0) is null then '0' else round(V.EntregadoT,0) end)as EntregadoT, (case when round(((V.DiferenciaA*100)-100),0) is null then '100' else round(((V.DiferenciaA*100)-100),0) end)as DiferenciaA, V.Semana, (case when round(V.PronosticoSA,0) is null then '0' else round(V.PronosticoSA,0) end) as PronosticoSA, (case when round(V.EntregadoSA,0) is null then '0' else round(V.EntregadoSA,0) end) as EntregadoSA, (case when round(((V.DiferenciaSA*100)-100),0) is null then '100' else round(((V.DiferenciaSA*100)-100),0) end) as DiferenciaSA FROM(select *,(V.EntregadoT/V.PronosticoA) AS DiferenciaA, (V.EntregadoSA / V.PronosticoSA) AS DiferenciaSA FROM(select C.IdAgen, C.Cod_Prod, P.Nombre, V.PronosticoT, V.EntregadoT, AA.PronosticoA, SA.Semana, SA.PronosticoSA, SA.EntregadoSA FROM ProdCamposCat C left join ProdProductoresCat P on C.Cod_Prod=P.Cod_Prod left join(select V.IdAgen, V.Cod_prod, sum(ISNULL(V.Pronostico,0)) as PronosticoT , sum(ISNULL(V.EntregadoT,0)) AS EntregadoT from(select V.IdAgen, V.Cod_prod, sum(V.Pronostico) as Pronostico, E.EntregadoT from(select * from(select V.IdAgen, V.Cod_Prod,max(V.Fecha)as fecha, V.[27],V.[28],V.[29],V.[30],V.[31],V.[32],V.[33],V.[34],V.[35],V.[36],V.[37],V.[38],V.[39],V.[40],V.[41],V.[42],V.[43],V.[44],V.[45],V.[46],V.[47],V.[48],V.[49],V.[50],V.[51],V.[52],V.[53],V.[1],V.[2],V.[3],V.[4],V.[5],V.[6],V.[7],V.[8],V.[9],V.[10],V.[11],V.[12],V.[13],V.[14],V.[15],V.[16],V.[17],V.[18],V.[19],V.[20],V.[21],V.[22],V.[23],V.[24],V.[25],V.[26] From(select distinct IdAgen, Cod_Prod, Fecha, SUM(isnull([27], 0)) as [27], sum(isnull([28], 0)) as [28], sum(isnull([29], 0)) as[29], sum(isnull([30], 0)) as [30], sum(isnull([31], 0)) as [31], sum(isnull([32], 0)) as [32], sum(isnull([33], 0)) as [33], sum(isnull([34], 0)) as [34], sum(isnull([35], 0)) as [35], sum(isnull([36], 0)) as [36], sum(isnull([37], 0)) as [37], sum(isnull([38], 0)) as [38], sum(isnull([39], 0)) as [39], sum(isnull([40], 0)) as [40], sum(isnull([41], 0)) as [41], sum(isnull([42], 0)) as [42], sum(isnull([43], 0)) as [43], sum(isnull([44], 0)) as [44], sum(isnull([45], 0)) as [45], sum(isnull([46], 0)) as [46], sum(isnull([47], 0)) as [47], sum(isnull([48], 0)) as [48], sum(isnull([49], 0)) as [49], sum(isnull([50], 0)) as [50], sum(isnull([51], 0)) as [51], sum(isnull([52], 0)) as [52], sum(isnull([53], 0)) as [53], sum(isnull([1], 0)) as [1], sum(isnull([2], 0)) as [2], sum(isnull([3], 0)) as [3], sum(isnull([4], 0)) as [4], sum(isnull([5], 0)) as [5], sum(isnull([6], 0)) as [6], sum(isnull([7], 0)) as [7], sum(isnull([8], 0)) as [8], sum(isnull([9], 0)) as [9], sum(isnull([10], 0)) as [10], sum(isnull([11], 0)) as [11], sum(isnull([12], 0)) as [12], sum(isnull([13], 0)) as [13], sum(isnull([14], 0)) as [14], sum(isnull([15], 0)) as [15],sum(isnull([16], 0)) as [16], sum(isnull([17], 0)) as [17], sum(isnull([18], 0)) as [18], sum(isnull([19], 0)) as [19], sum(isnull([20], 0)) as [20], sum(isnull([21], 0)) as [21], sum(isnull([22], 0)) as [22], sum(isnull([23], 0)) as [23], sum(isnull([24], 0)) as [24], sum(isnull([25], 0)) as [25], sum(isnull([26], 0)) as [26] from SIPGProyeccion P where Temporada = (select Temporada from CatSemanas where getdate() between Inicio and Fin) and  " +
                        "Fecha = (select max(Fecha) from SIPGProyeccion where IdAgen = " + (short)Session["IdAgen"] + ") group by IdAgen, Cod_Prod,Fecha)V GROUP BY V.IdAgen, V.Cod_Prod,V.[27],V.[28],V.[29],V.[30],V.[31],V.[32],V.[33],V.[34],V.[35],V.[36],V.[37],V.[38],V.[39],V.[40],V.[41],V.[42],V.[43],V.[44],V.[45],V.[46],V.[47],V.[48],V.[49],V.[50],V.[51],V.[52],V.[53],V.[1],V.[2],V.[3],V.[4],V.[5],V.[6],V.[7],V.[8],V.[9],V.[10],V.[11],V.[12],V.[13],V.[14],V.[15],V.[16],V.[17],V.[18],V.[19],V.[20],V.[21],V.[22],V.[23],V.[24],V.[25],V.[26])V UNPIVOT(Pronostico FOR Semana in ([27],[28],[29],[30],[31],[32],[33],[34],[35],[36],[37],[38],[39],[40],[41],[42],[43],[44],[45],[46],[47],[48],[49],[50],[51],[52],[53],[1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],[17],[18],[19],[20],[21],[22],[23],[24],[25],[26]))X Where Pronostico <> 0)V left join(select C.IdAgen, C.Cod_Prod, sum(C.Convertidas) as EntregadoT FROM(select IdAgen, Cod_Prod, Cod_Campo, sum(Convertidas) as Convertidas FROM UV_ProdRecepcion where CodEstatus <> 'C' and Temporada = (select Temporada from CatSemanas where getdate() between Inicio and Fin) group by IdAgen, Cod_prod, Cod_Campo)C GROUP BY C.IdAgen, C.Cod_Prod)E ON V.IdAgen = E.IdAgen AND V.Cod_Prod = E.Cod_Prod group by V.IdAgen, V.Cod_prod, E.EntregadoT)V GROUP BY V.IdAgen, V.Cod_prod)V ON C.IdAgen = V.IdAgen and C.Cod_Prod = V.Cod_Prod left join(select V.IdAgen, V.Cod_prod, sum(isnull(V.Pronostico,0)) as PronosticoA FROM CatSemanas S left join(select * from(select V.IdAgen, V.Cod_Prod, max(V.Fecha) as fecha, V.[27], V.[28], V.[29], V.[30], V.[31], V.[32], V.[33], V.[34], V.[35], V.[36], V.[37], V.[38], V.[39], V.[40], V.[41], V.[42], V.[43], V.[44], V.[45], V.[46], V.[47], V.[48], V.[49], V.[50], V.[51], V.[52], V.[53], V.[1], V.[2], V.[3], V.[4], V.[5], V.[6], V.[7], V.[8], V.[9], V.[10], V.[11], V.[12], V.[13], V.[14], V.[15], V.[16], V.[17], V.[18], V.[19], V.[20], V.[21], V.[22], V.[23], V.[24], V.[25], V.[26] From(select distinct IdAgen, Cod_Prod, Fecha, SUM(isnull([27], 0)) as [27], sum(isnull([28], 0)) as [28], sum(isnull([29], 0)) as[29], sum(isnull([30], 0)) as [30], sum(isnull([31], 0)) as [31], sum(isnull([32], 0)) as [32], sum(isnull([33], 0)) as [33], sum(isnull([34], 0)) as [34], sum(isnull([35], 0)) as [35], sum(isnull([36], 0)) as [36], sum(isnull([37], 0)) as [37], sum(isnull([38], 0)) as [38], sum(isnull([39], 0)) as [39], sum(isnull([40], 0)) as [40], sum(isnull([41], 0)) as [41], sum(isnull([42], 0)) as [42], sum(isnull([43], 0)) as [43], sum(isnull([44], 0)) as [44], sum(isnull([45], 0)) as [45], sum(isnull([46], 0)) as [46], sum(isnull([47], 0)) as [47], sum(isnull([48], 0)) as [48], sum(isnull([49], 0)) as [49], sum(isnull([50], 0)) as [50], sum(isnull([51], 0)) as [51], sum(isnull([52], 0)) as [52],sum(isnull([53], 0)) as [53], sum(isnull([1], 0)) as [1], sum(isnull([2], 0)) as [2], sum(isnull([3], 0)) as [3], sum(isnull([4], 0)) as [4], sum(isnull([5], 0)) as [5], sum(isnull([6], 0)) as [6], sum(isnull([7], 0)) as [7], sum(isnull([8], 0)) as [8], sum(isnull([9], 0)) as [9], sum(isnull([10], 0)) as [10], sum(isnull([11], 0)) as [11], sum(isnull([12], 0)) as [12], sum(isnull([13], 0)) as [13], sum(isnull([14], 0)) as [14], sum(isnull([15], 0)) as [15],sum(isnull([16], 0)) as [16], sum(isnull([17], 0)) as [17], sum(isnull([18], 0)) as [18], sum(isnull([19], 0)) as [19], sum(isnull([20], 0)) as [20], sum(isnull([21], 0)) as [21], sum(isnull([22], 0)) as [22], sum(isnull([23], 0)) as [23], sum(isnull([24], 0)) as [24], sum(isnull([25], 0)) as [25], sum(isnull([26], 0)) as [26] from SIPGProyeccion P where Temporada = (select Temporada from CatSemanas where getdate() between Inicio and Fin) " +
                        "and Fecha = (select max(Fecha) from SIPGProyeccion where IdAgen = " + (short)Session["IdAgen"] + ") group by IdAgen, Cod_Prod,Fecha)V GROUP BY V.IdAgen, V.Cod_Prod,V.[27],V.[28],V.[29],V.[30],V.[31],V.[32],V.[33],V.[34],V.[35],V.[36],V.[37],V.[38],V.[39],V.[40],V.[41],V.[42],V.[43],V.[44],V.[45],V.[46],V.[47],V.[48],V.[49],V.[50],V.[51],V.[52],V.[53],V.[1],V.[2],V.[3],V.[4],V.[5],V.[6],V.[7],V.[8],V.[9],V.[10],V.[11],V.[12],V.[13],V.[14],V.[15],V.[16],V.[17],V.[18],V.[19],V.[20],V.[21],V.[22],V.[23],V.[24],V.[25],V.[26])V UNPIVOT(Pronostico FOR Semana in ([27],[28],[29],[30],[31],[32],[33],[34],[35],[36],[37],[38],[39],[40],[41],[42],[43],[44],[45],[46],[47],[48],[49],[50],[51],[52],[53],[1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],[17],[18],[19],[20],[21],[22],[23],[24],[25],[26]))X where Pronostico <> 0)V on S.Semana = V.Semana where S.Inicio between(select Inicio from catsemanas where temporada = (select Temporada from CatSemanas where getdate() between Inicio and Fin) and semana = 27) and getdate() group by V.IdAgen,V.Cod_prod)AA ON V.IdAgen = AA.IdAgen and V.Cod_Prod = AA.Cod_Prod LEFT JOIN(SELECT PA.IdAgen, PA.Cod_Prod, SUM(isnull(PA.Pronostico,0)) AS PronosticoSA, sum(isnull(EA.Entregado, 0)) AS EntregadoSA, S.Semana FROM CatSemanas S left join(select* from(select V.IdAgen, V.Cod_Prod, max(V.Fecha)as fecha, V.[27],V.[28],V.[29],V.[30],V.[31],V.[32],V.[33],V.[34],V.[35],V.[36],V.[37],V.[38],V.[39],V.[40],V.[41],V.[42],V.[43],V.[44],V.[45],V.[46],V.[47],V.[48],V.[49],V.[50],V.[51],V.[52],V.[53],V.[1],V.[2],V.[3],V.[4],V.[5],V.[6],V.[7],V.[8],V.[9],V.[10],V.[11],V.[12],V.[13],V.[14],V.[15],V.[16],V.[17],V.[18],V.[19],V.[20],V.[21],V.[22],V.[23],V.[24],V.[25],V.[26] From(select distinct IdAgen, Cod_Prod, Fecha, SUM(isnull([27], 0)) as [27], sum(isnull([28], 0)) as [28], sum(isnull([29], 0)) as[29], sum(isnull([30], 0)) as [30], sum(isnull([31], 0)) as [31], sum(isnull([32], 0)) as [32], sum(isnull([33], 0)) as [33], sum(isnull([34], 0)) as [34], sum(isnull([35], 0)) as [35], sum(isnull([36], 0)) as [36], sum(isnull([37], 0)) as [37], sum(isnull([38], 0)) as [38], sum(isnull([39], 0)) as [39], sum(isnull([40], 0)) as [40], sum(isnull([41], 0)) as [41], sum(isnull([42], 0)) as [42], sum(isnull([43], 0)) as [43], sum(isnull([44], 0)) as [44], sum(isnull([45], 0)) as [45], sum(isnull([46], 0)) as [46], sum(isnull([47], 0)) as [47], sum(isnull([48], 0)) as [48], sum(isnull([49], 0)) as [49], sum(isnull([50], 0)) as [50], sum(isnull([51], 0)) as [51], sum(isnull([52], 0)) as [52],sum(isnull([53], 0)) as [53], sum(isnull([1], 0)) as [1], sum(isnull([2], 0)) as [2], sum(isnull([3], 0)) as [3], sum(isnull([4], 0)) as [4], sum(isnull([5], 0)) as [5], sum(isnull([6], 0)) as [6], sum(isnull([7], 0)) as [7], sum(isnull([8], 0)) as [8], sum(isnull([9], 0)) as [9], sum(isnull([10], 0)) as [10], sum(isnull([11], 0)) as [11], sum(isnull([12], 0)) as [12], sum(isnull([13], 0)) as [13], sum(isnull([14], 0)) as [14], sum(isnull([15], 0)) as [15],sum(isnull([16], 0)) as [16], sum(isnull([17], 0)) as [17], sum(isnull([18], 0)) as [18], sum(isnull([19], 0)) as [19], sum(isnull([20], 0)) as [20], sum(isnull([21], 0)) as [21], sum(isnull([22], 0)) as [22], sum(isnull([23], 0)) as [23], sum(isnull([24], 0)) as [24], sum(isnull([25], 0)) as [25], sum(isnull([26], 0)) as [26] from SIPGProyeccion P where Temporada = (select Temporada from CatSemanas where getdate() between Inicio and Fin) " +
                        "and Fecha = (select max(Fecha) from SIPGProyeccion where IdAgen = " + (short)Session["IdAgen"] + ") group by IdAgen, Cod_Prod,Fecha)V GROUP BY V.IdAgen, V.Cod_Prod,V.[27],V.[28],V.[29],V.[30],V.[31],V.[32],V.[33],V.[34],V.[35],V.[36],V.[37],V.[38],V.[39],V.[40],V.[41],V.[42],V.[43],V.[44],V.[45],V.[46],V.[47],V.[48],V.[49],V.[50],V.[51],V.[52],V.[53],V.[1],V.[2],V.[3],V.[4],V.[5],V.[6],V.[7],V.[8],V.[9],V.[10],V.[11],V.[12],V.[13],V.[14],V.[15],V.[16],V.[17],V.[18],V.[19],V.[20],V.[21],V.[22],V.[23],V.[24],V.[25],V.[26])V UNPIVOT(Pronostico FOR Semana in ([27],[28],[29],[30],[31],[32],[33],[34],[35],[36],[37],[38],[39],[40],[41],[42],[43],[44],[45],[46],[47],[48],[49],[50],[51],[52],[53],[1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],[17],[18],[19],[20],[21],[22],[23],[24],[25],[26]))X where Pronostico <> 0)PA ON S.Semana = PA.Semana Left Join(select C.IdAgen, C.Cod_Prod, sum(C.Convertidas) as Entregado, C.Semana FROM(select IdAgen, Cod_Prod, sum(Convertidas) as Convertidas, Semana FROM UV_ProdRecepcion where CodEstatus <> 'C' and Temporada = (select Temporada from CatSemanas where getdate() between Inicio and Fin) group by IdAgen, Cod_prod, Cod_Campo, Semana)C GROUP BY C.IdAgen, C.Cod_Prod, C.Semana)EA ON PA.IdAgen = EA.IdAgen and PA.Cod_Prod = EA.Cod_Prod and S.Semana = EA.Semana WHERE getdate() between S.Inicio and S.Fin GROUP BY PA.IdAgen, PA.Cod_Prod, S.Semana)SA on V.IdAgen = SA.IdAgen AND V.Cod_Prod = SA.Cod_Prod group by C.IdAgen, C.Cod_Prod, P.Nombre, V.PronosticoT, V.EntregadoT, AA.PronosticoA, SA.Semana, SA.PronosticoSA, SA.EntregadoSA)V)V)C " +
                        "WHERE C.IdAgen = " + (short)Session["IdAgen"] + "  and C.Cod_Prod not in(Select distinct V.Cod_Prod FROM(select* from CatSemanas where Getdate() between Inicio and Fin)S left join SIPGVisitas V on V.Fecha between S.Inicio and S.Fin " +
                        "where V.IdAgen = " + (short)Session["IdAgen"] + ") and C.DiferenciaA > 25 and C.DiferenciaA <=50 or C.DiferenciaA < -25 and C.DiferenciaA >= -50 ORDER BY C.DiferenciaA DESC").ToList();
                }
                else if (filtro == "todo")
                {
                    curva = bd.Database.SqlQuery<ClassCurva>("select C.Cod_Prod, C.Nombre as Productor, C.PronosticoT, C.PronosticoA, C.EntregadoT, (CASE WHEN C.EntregadoT ='0' AND C.PronosticoA ='0' THEN '' ELSE C.DiferenciaA END) AS DiferenciaA, C.Semana, C.PronosticoSA, C.EntregadoSA, (CASE WHEN C.EntregadoSA ='0' AND  C.PronosticoSA ='0' THEN '' ELSE C.DiferenciaSA END) AS DiferenciaSA from(SELECT distinct V.IdAgen, V.Cod_Prod, V.Nombre, (case when round(V.PronosticoT,0) is null then '0' else round(V.PronosticoT,0) end) as PronosticoT, (case when round(V.PronosticoA,0) is null then '0' else round(V.PronosticoA,0) end) as PronosticoA, (case when round(V.EntregadoT,0) is null then '0' else round(V.EntregadoT,0) end)as EntregadoT, (case when round(((V.DiferenciaA*100)-100),0) is null then '100' else round(((V.DiferenciaA*100)-100),0) end)as DiferenciaA, V.Semana, (case when round(V.PronosticoSA,0) is null then '0' else round(V.PronosticoSA,0) end) as PronosticoSA, (case when round(V.EntregadoSA,0) is null then '0' else round(V.EntregadoSA,0) end) as EntregadoSA, (case when round(((V.DiferenciaSA*100)-100),0) is null then '100' else round(((V.DiferenciaSA*100)-100),0) end) as DiferenciaSA FROM(select *,(V.EntregadoT/V.PronosticoA) AS DiferenciaA, (V.EntregadoSA / V.PronosticoSA) AS DiferenciaSA FROM(select C.IdAgen, C.Cod_Prod, P.Nombre, V.PronosticoT, V.EntregadoT, AA.PronosticoA, SA.Semana, SA.PronosticoSA, SA.EntregadoSA FROM ProdCamposCat C left join ProdProductoresCat P on C.Cod_Prod=P.Cod_Prod left join(select V.IdAgen, V.Cod_prod, sum(ISNULL(V.Pronostico,0)) as PronosticoT , sum(ISNULL(V.EntregadoT,0)) AS EntregadoT from(select V.IdAgen, V.Cod_prod, sum(V.Pronostico) as Pronostico, E.EntregadoT from(select * from(select V.IdAgen, V.Cod_Prod,max(V.Fecha)as fecha, V.[27],V.[28],V.[29],V.[30],V.[31],V.[32],V.[33],V.[34],V.[35],V.[36],V.[37],V.[38],V.[39],V.[40],V.[41],V.[42],V.[43],V.[44],V.[45],V.[46],V.[47],V.[48],V.[49],V.[50],V.[51],V.[52],V.[53],V.[1],V.[2],V.[3],V.[4],V.[5],V.[6],V.[7],V.[8],V.[9],V.[10],V.[11],V.[12],V.[13],V.[14],V.[15],V.[16],V.[17],V.[18],V.[19],V.[20],V.[21],V.[22],V.[23],V.[24],V.[25],V.[26] From(select distinct IdAgen, Cod_Prod, Fecha, SUM(isnull([27], 0)) as [27], sum(isnull([28], 0)) as [28], sum(isnull([29], 0)) as[29], sum(isnull([30], 0)) as [30], sum(isnull([31], 0)) as [31], sum(isnull([32], 0)) as [32], sum(isnull([33], 0)) as [33], sum(isnull([34], 0)) as [34], sum(isnull([35], 0)) as [35], sum(isnull([36], 0)) as [36], sum(isnull([37], 0)) as [37], sum(isnull([38], 0)) as [38], sum(isnull([39], 0)) as [39], sum(isnull([40], 0)) as [40], sum(isnull([41], 0)) as [41], sum(isnull([42], 0)) as [42], sum(isnull([43], 0)) as [43], sum(isnull([44], 0)) as [44], sum(isnull([45], 0)) as [45], sum(isnull([46], 0)) as [46], sum(isnull([47], 0)) as [47], sum(isnull([48], 0)) as [48], sum(isnull([49], 0)) as [49], sum(isnull([50], 0)) as [50], sum(isnull([51], 0)) as [51], sum(isnull([52], 0)) as [52], sum(isnull([53], 0)) as [53], sum(isnull([1], 0)) as [1], sum(isnull([2], 0)) as [2], sum(isnull([3], 0)) as [3], sum(isnull([4], 0)) as [4], sum(isnull([5], 0)) as [5], sum(isnull([6], 0)) as [6], sum(isnull([7], 0)) as [7], sum(isnull([8], 0)) as [8], sum(isnull([9], 0)) as [9], sum(isnull([10], 0)) as [10], sum(isnull([11], 0)) as [11], sum(isnull([12], 0)) as [12], sum(isnull([13], 0)) as [13], sum(isnull([14], 0)) as [14], sum(isnull([15], 0)) as [15],sum(isnull([16], 0)) as [16], sum(isnull([17], 0)) as [17], sum(isnull([18], 0)) as [18], sum(isnull([19], 0)) as [19], sum(isnull([20], 0)) as [20], sum(isnull([21], 0)) as [21], sum(isnull([22], 0)) as [22], sum(isnull([23], 0)) as [23], sum(isnull([24], 0)) as [24], sum(isnull([25], 0)) as [25], sum(isnull([26], 0)) as [26] from SIPGProyeccion P where Temporada = (select Temporada from CatSemanas where getdate() between Inicio and Fin) and  " +
                        "Fecha = (select max(Fecha) from SIPGProyeccion where IdAgen = " + (short)Session["IdAgen"] + ") group by IdAgen, Cod_Prod,Fecha)V GROUP BY V.IdAgen, V.Cod_Prod,V.[27],V.[28],V.[29],V.[30],V.[31],V.[32],V.[33],V.[34],V.[35],V.[36],V.[37],V.[38],V.[39],V.[40],V.[41],V.[42],V.[43],V.[44],V.[45],V.[46],V.[47],V.[48],V.[49],V.[50],V.[51],V.[52],V.[53],V.[1],V.[2],V.[3],V.[4],V.[5],V.[6],V.[7],V.[8],V.[9],V.[10],V.[11],V.[12],V.[13],V.[14],V.[15],V.[16],V.[17],V.[18],V.[19],V.[20],V.[21],V.[22],V.[23],V.[24],V.[25],V.[26])V UNPIVOT(Pronostico FOR Semana in ([27],[28],[29],[30],[31],[32],[33],[34],[35],[36],[37],[38],[39],[40],[41],[42],[43],[44],[45],[46],[47],[48],[49],[50],[51],[52],[53],[1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],[17],[18],[19],[20],[21],[22],[23],[24],[25],[26]))X Where Pronostico <> 0)V left join(select C.IdAgen, C.Cod_Prod, sum(C.Convertidas) as EntregadoT FROM(select IdAgen, Cod_Prod, Cod_Campo, sum(Convertidas) as Convertidas FROM UV_ProdRecepcion where CodEstatus <> 'C' and Temporada = (select Temporada from CatSemanas where getdate() between Inicio and Fin) group by IdAgen, Cod_prod, Cod_Campo)C GROUP BY C.IdAgen, C.Cod_Prod)E ON V.IdAgen = E.IdAgen AND V.Cod_Prod = E.Cod_Prod group by V.IdAgen, V.Cod_prod, E.EntregadoT)V GROUP BY V.IdAgen, V.Cod_prod)V ON C.IdAgen = V.IdAgen and C.Cod_Prod = V.Cod_Prod left join(select V.IdAgen, V.Cod_prod, sum(isnull(V.Pronostico,0)) as PronosticoA FROM CatSemanas S left join(select * from(select V.IdAgen, V.Cod_Prod, max(V.Fecha) as fecha, V.[27], V.[28], V.[29], V.[30], V.[31], V.[32], V.[33], V.[34], V.[35], V.[36], V.[37], V.[38], V.[39], V.[40], V.[41], V.[42], V.[43], V.[44], V.[45], V.[46], V.[47], V.[48], V.[49], V.[50], V.[51], V.[52], V.[53], V.[1], V.[2], V.[3], V.[4], V.[5], V.[6], V.[7], V.[8], V.[9], V.[10], V.[11], V.[12], V.[13], V.[14], V.[15], V.[16], V.[17], V.[18], V.[19], V.[20], V.[21], V.[22], V.[23], V.[24], V.[25], V.[26] From(select distinct IdAgen, Cod_Prod, Fecha, SUM(isnull([27], 0)) as [27], sum(isnull([28], 0)) as [28], sum(isnull([29], 0)) as[29], sum(isnull([30], 0)) as [30], sum(isnull([31], 0)) as [31], sum(isnull([32], 0)) as [32], sum(isnull([33], 0)) as [33], sum(isnull([34], 0)) as [34], sum(isnull([35], 0)) as [35], sum(isnull([36], 0)) as [36], sum(isnull([37], 0)) as [37], sum(isnull([38], 0)) as [38], sum(isnull([39], 0)) as [39], sum(isnull([40], 0)) as [40], sum(isnull([41], 0)) as [41], sum(isnull([42], 0)) as [42], sum(isnull([43], 0)) as [43], sum(isnull([44], 0)) as [44], sum(isnull([45], 0)) as [45], sum(isnull([46], 0)) as [46], sum(isnull([47], 0)) as [47], sum(isnull([48], 0)) as [48], sum(isnull([49], 0)) as [49], sum(isnull([50], 0)) as [50], sum(isnull([51], 0)) as [51], sum(isnull([52], 0)) as [52],sum(isnull([53], 0)) as [53], sum(isnull([1], 0)) as [1], sum(isnull([2], 0)) as [2], sum(isnull([3], 0)) as [3], sum(isnull([4], 0)) as [4], sum(isnull([5], 0)) as [5], sum(isnull([6], 0)) as [6], sum(isnull([7], 0)) as [7], sum(isnull([8], 0)) as [8], sum(isnull([9], 0)) as [9], sum(isnull([10], 0)) as [10], sum(isnull([11], 0)) as [11], sum(isnull([12], 0)) as [12], sum(isnull([13], 0)) as [13], sum(isnull([14], 0)) as [14], sum(isnull([15], 0)) as [15],sum(isnull([16], 0)) as [16], sum(isnull([17], 0)) as [17], sum(isnull([18], 0)) as [18], sum(isnull([19], 0)) as [19], sum(isnull([20], 0)) as [20], sum(isnull([21], 0)) as [21], sum(isnull([22], 0)) as [22], sum(isnull([23], 0)) as [23], sum(isnull([24], 0)) as [24], sum(isnull([25], 0)) as [25], sum(isnull([26], 0)) as [26] from SIPGProyeccion P where Temporada = (select Temporada from CatSemanas where getdate() between Inicio and Fin) " +
                        "and Fecha = (select max(Fecha) from SIPGProyeccion where IdAgen = " + (short)Session["IdAgen"] + ") group by IdAgen, Cod_Prod,Fecha)V GROUP BY V.IdAgen, V.Cod_Prod,V.[27],V.[28],V.[29],V.[30],V.[31],V.[32],V.[33],V.[34],V.[35],V.[36],V.[37],V.[38],V.[39],V.[40],V.[41],V.[42],V.[43],V.[44],V.[45],V.[46],V.[47],V.[48],V.[49],V.[50],V.[51],V.[52],V.[53],V.[1],V.[2],V.[3],V.[4],V.[5],V.[6],V.[7],V.[8],V.[9],V.[10],V.[11],V.[12],V.[13],V.[14],V.[15],V.[16],V.[17],V.[18],V.[19],V.[20],V.[21],V.[22],V.[23],V.[24],V.[25],V.[26])V UNPIVOT(Pronostico FOR Semana in ([27],[28],[29],[30],[31],[32],[33],[34],[35],[36],[37],[38],[39],[40],[41],[42],[43],[44],[45],[46],[47],[48],[49],[50],[51],[52],[53],[1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],[17],[18],[19],[20],[21],[22],[23],[24],[25],[26]))X where Pronostico <> 0)V on S.Semana = V.Semana where S.Inicio between(select Inicio from catsemanas where temporada = (select Temporada from CatSemanas where getdate() between Inicio and Fin) and semana = 27) and getdate() group by V.IdAgen,V.Cod_prod)AA ON V.IdAgen = AA.IdAgen and V.Cod_Prod = AA.Cod_Prod LEFT JOIN(SELECT PA.IdAgen, PA.Cod_Prod, SUM(isnull(PA.Pronostico,0)) AS PronosticoSA, sum(isnull(EA.Entregado, 0)) AS EntregadoSA, S.Semana FROM CatSemanas S left join(select* from(select V.IdAgen, V.Cod_Prod, max(V.Fecha)as fecha, V.[27],V.[28],V.[29],V.[30],V.[31],V.[32],V.[33],V.[34],V.[35],V.[36],V.[37],V.[38],V.[39],V.[40],V.[41],V.[42],V.[43],V.[44],V.[45],V.[46],V.[47],V.[48],V.[49],V.[50],V.[51],V.[52],V.[53],V.[1],V.[2],V.[3],V.[4],V.[5],V.[6],V.[7],V.[8],V.[9],V.[10],V.[11],V.[12],V.[13],V.[14],V.[15],V.[16],V.[17],V.[18],V.[19],V.[20],V.[21],V.[22],V.[23],V.[24],V.[25],V.[26] From(select distinct IdAgen, Cod_Prod, Fecha, SUM(isnull([27], 0)) as [27], sum(isnull([28], 0)) as [28], sum(isnull([29], 0)) as[29], sum(isnull([30], 0)) as [30], sum(isnull([31], 0)) as [31], sum(isnull([32], 0)) as [32], sum(isnull([33], 0)) as [33], sum(isnull([34], 0)) as [34], sum(isnull([35], 0)) as [35], sum(isnull([36], 0)) as [36], sum(isnull([37], 0)) as [37], sum(isnull([38], 0)) as [38], sum(isnull([39], 0)) as [39], sum(isnull([40], 0)) as [40], sum(isnull([41], 0)) as [41], sum(isnull([42], 0)) as [42], sum(isnull([43], 0)) as [43], sum(isnull([44], 0)) as [44], sum(isnull([45], 0)) as [45], sum(isnull([46], 0)) as [46], sum(isnull([47], 0)) as [47], sum(isnull([48], 0)) as [48], sum(isnull([49], 0)) as [49], sum(isnull([50], 0)) as [50], sum(isnull([51], 0)) as [51], sum(isnull([52], 0)) as [52],sum(isnull([53], 0)) as [53], sum(isnull([1], 0)) as [1], sum(isnull([2], 0)) as [2], sum(isnull([3], 0)) as [3], sum(isnull([4], 0)) as [4], sum(isnull([5], 0)) as [5], sum(isnull([6], 0)) as [6], sum(isnull([7], 0)) as [7], sum(isnull([8], 0)) as [8], sum(isnull([9], 0)) as [9], sum(isnull([10], 0)) as [10], sum(isnull([11], 0)) as [11], sum(isnull([12], 0)) as [12], sum(isnull([13], 0)) as [13], sum(isnull([14], 0)) as [14], sum(isnull([15], 0)) as [15],sum(isnull([16], 0)) as [16], sum(isnull([17], 0)) as [17], sum(isnull([18], 0)) as [18], sum(isnull([19], 0)) as [19], sum(isnull([20], 0)) as [20], sum(isnull([21], 0)) as [21], sum(isnull([22], 0)) as [22], sum(isnull([23], 0)) as [23], sum(isnull([24], 0)) as [24], sum(isnull([25], 0)) as [25], sum(isnull([26], 0)) as [26] from SIPGProyeccion P where Temporada = (select Temporada from CatSemanas where getdate() between Inicio and Fin) " +
                        "and Fecha = (select max(Fecha) from SIPGProyeccion where IdAgen = " + (short)Session["IdAgen"] + ") group by IdAgen, Cod_Prod,Fecha)V GROUP BY V.IdAgen, V.Cod_Prod,V.[27],V.[28],V.[29],V.[30],V.[31],V.[32],V.[33],V.[34],V.[35],V.[36],V.[37],V.[38],V.[39],V.[40],V.[41],V.[42],V.[43],V.[44],V.[45],V.[46],V.[47],V.[48],V.[49],V.[50],V.[51],V.[52],V.[53],V.[1],V.[2],V.[3],V.[4],V.[5],V.[6],V.[7],V.[8],V.[9],V.[10],V.[11],V.[12],V.[13],V.[14],V.[15],V.[16],V.[17],V.[18],V.[19],V.[20],V.[21],V.[22],V.[23],V.[24],V.[25],V.[26])V UNPIVOT(Pronostico FOR Semana in ([27],[28],[29],[30],[31],[32],[33],[34],[35],[36],[37],[38],[39],[40],[41],[42],[43],[44],[45],[46],[47],[48],[49],[50],[51],[52],[53],[1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],[17],[18],[19],[20],[21],[22],[23],[24],[25],[26]))X where Pronostico <> 0)PA ON S.Semana = PA.Semana Left Join(select C.IdAgen, C.Cod_Prod, sum(C.Convertidas) as Entregado, C.Semana FROM(select IdAgen, Cod_Prod, sum(Convertidas) as Convertidas, Semana FROM UV_ProdRecepcion where CodEstatus <> 'C' and Temporada = (select Temporada from CatSemanas where getdate() between Inicio and Fin) group by IdAgen, Cod_prod, Cod_Campo, Semana)C GROUP BY C.IdAgen, C.Cod_Prod, C.Semana)EA ON PA.IdAgen = EA.IdAgen and PA.Cod_Prod = EA.Cod_Prod and S.Semana = EA.Semana WHERE getdate() between S.Inicio and S.Fin GROUP BY PA.IdAgen, PA.Cod_Prod, S.Semana)SA on V.IdAgen = SA.IdAgen AND V.Cod_Prod = SA.Cod_Prod group by C.IdAgen, C.Cod_Prod, P.Nombre, V.PronosticoT, V.EntregadoT, AA.PronosticoA, SA.Semana, SA.PronosticoSA, SA.EntregadoSA)V)V)C " +
                        "WHERE C.IdAgen = " + (short)Session["IdAgen"] + "  and C.Cod_Prod not in(Select distinct V.Cod_Prod FROM(select* from CatSemanas where Getdate() between Inicio and Fin)S left join SIPGVisitas V on V.Fecha between S.Inicio and S.Fin " +
                        "where V.IdAgen = " + (short)Session["IdAgen"] + ") ORDER BY C.DiferenciaA DESC").ToList();
                }
                else if (filtro == "urgentes")
                {
                    curva = bd.Database.SqlQuery<ClassCurva>("select C.Cod_Prod, C.Nombre as Productor, C.PronosticoT, C.PronosticoA, C.EntregadoT, (CASE WHEN C.EntregadoT ='0' AND C.PronosticoA ='0' THEN '' ELSE C.DiferenciaA END) AS DiferenciaA, C.Semana, C.PronosticoSA, C.EntregadoSA, (CASE WHEN C.EntregadoSA ='0' AND  C.PronosticoSA ='0' THEN '' ELSE C.DiferenciaSA END) AS DiferenciaSA from(SELECT distinct V.IdAgen, V.Cod_Prod, V.Nombre, (case when round(V.PronosticoT,0) is null then '0' else round(V.PronosticoT,0) end) as PronosticoT, (case when round(V.PronosticoA,0) is null then '0' else round(V.PronosticoA,0) end) as PronosticoA, (case when round(V.EntregadoT,0) is null then '0' else round(V.EntregadoT,0) end)as EntregadoT, (case when round(((V.DiferenciaA*100)-100),0) is null then '100' else round(((V.DiferenciaA*100)-100),0) end)as DiferenciaA, V.Semana, (case when round(V.PronosticoSA,0) is null then '0' else round(V.PronosticoSA,0) end) as PronosticoSA, (case when round(V.EntregadoSA,0) is null then '0' else round(V.EntregadoSA,0) end) as EntregadoSA, (case when round(((V.DiferenciaSA*100)-100),0) is null then '100' else round(((V.DiferenciaSA*100)-100),0) end) as DiferenciaSA FROM(select *,(V.EntregadoT/V.PronosticoA) AS DiferenciaA, (V.EntregadoSA / V.PronosticoSA) AS DiferenciaSA FROM(select C.IdAgen, C.Cod_Prod, P.Nombre, V.PronosticoT, V.EntregadoT, AA.PronosticoA, SA.Semana, SA.PronosticoSA, SA.EntregadoSA FROM ProdCamposCat C left join ProdProductoresCat P on C.Cod_Prod=P.Cod_Prod left join(select V.IdAgen, V.Cod_prod, sum(ISNULL(V.Pronostico,0)) as PronosticoT , sum(ISNULL(V.EntregadoT,0)) AS EntregadoT from(select V.IdAgen, V.Cod_prod, sum(V.Pronostico) as Pronostico, E.EntregadoT from(select * from(select V.IdAgen, V.Cod_Prod,max(V.Fecha)as fecha, V.[27],V.[28],V.[29],V.[30],V.[31],V.[32],V.[33],V.[34],V.[35],V.[36],V.[37],V.[38],V.[39],V.[40],V.[41],V.[42],V.[43],V.[44],V.[45],V.[46],V.[47],V.[48],V.[49],V.[50],V.[51],V.[52],V.[53],V.[1],V.[2],V.[3],V.[4],V.[5],V.[6],V.[7],V.[8],V.[9],V.[10],V.[11],V.[12],V.[13],V.[14],V.[15],V.[16],V.[17],V.[18],V.[19],V.[20],V.[21],V.[22],V.[23],V.[24],V.[25],V.[26] From(select distinct IdAgen, Cod_Prod, Fecha, SUM(isnull([27], 0)) as [27], sum(isnull([28], 0)) as [28], sum(isnull([29], 0)) as[29], sum(isnull([30], 0)) as [30], sum(isnull([31], 0)) as [31], sum(isnull([32], 0)) as [32], sum(isnull([33], 0)) as [33], sum(isnull([34], 0)) as [34], sum(isnull([35], 0)) as [35], sum(isnull([36], 0)) as [36], sum(isnull([37], 0)) as [37], sum(isnull([38], 0)) as [38], sum(isnull([39], 0)) as [39], sum(isnull([40], 0)) as [40], sum(isnull([41], 0)) as [41], sum(isnull([42], 0)) as [42], sum(isnull([43], 0)) as [43], sum(isnull([44], 0)) as [44], sum(isnull([45], 0)) as [45], sum(isnull([46], 0)) as [46], sum(isnull([47], 0)) as [47], sum(isnull([48], 0)) as [48], sum(isnull([49], 0)) as [49], sum(isnull([50], 0)) as [50], sum(isnull([51], 0)) as [51], sum(isnull([52], 0)) as [52], sum(isnull([53], 0)) as [53], sum(isnull([1], 0)) as [1], sum(isnull([2], 0)) as [2], sum(isnull([3], 0)) as [3], sum(isnull([4], 0)) as [4], sum(isnull([5], 0)) as [5], sum(isnull([6], 0)) as [6], sum(isnull([7], 0)) as [7], sum(isnull([8], 0)) as [8], sum(isnull([9], 0)) as [9], sum(isnull([10], 0)) as [10], sum(isnull([11], 0)) as [11], sum(isnull([12], 0)) as [12], sum(isnull([13], 0)) as [13], sum(isnull([14], 0)) as [14], sum(isnull([15], 0)) as [15],sum(isnull([16], 0)) as [16], sum(isnull([17], 0)) as [17], sum(isnull([18], 0)) as [18], sum(isnull([19], 0)) as [19], sum(isnull([20], 0)) as [20], sum(isnull([21], 0)) as [21], sum(isnull([22], 0)) as [22], sum(isnull([23], 0)) as [23], sum(isnull([24], 0)) as [24], sum(isnull([25], 0)) as [25], sum(isnull([26], 0)) as [26] from SIPGProyeccion P where Temporada = (select Temporada from CatSemanas where getdate() between Inicio and Fin) and  " +
                        "Fecha = (select max(Fecha) from SIPGProyeccion where IdAgen = " + (short)Session["IdAgen"] + ") group by IdAgen, Cod_Prod,Fecha)V GROUP BY V.IdAgen, V.Cod_Prod,V.[27],V.[28],V.[29],V.[30],V.[31],V.[32],V.[33],V.[34],V.[35],V.[36],V.[37],V.[38],V.[39],V.[40],V.[41],V.[42],V.[43],V.[44],V.[45],V.[46],V.[47],V.[48],V.[49],V.[50],V.[51],V.[52],V.[53],V.[1],V.[2],V.[3],V.[4],V.[5],V.[6],V.[7],V.[8],V.[9],V.[10],V.[11],V.[12],V.[13],V.[14],V.[15],V.[16],V.[17],V.[18],V.[19],V.[20],V.[21],V.[22],V.[23],V.[24],V.[25],V.[26])V UNPIVOT(Pronostico FOR Semana in ([27],[28],[29],[30],[31],[32],[33],[34],[35],[36],[37],[38],[39],[40],[41],[42],[43],[44],[45],[46],[47],[48],[49],[50],[51],[52],[53],[1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],[17],[18],[19],[20],[21],[22],[23],[24],[25],[26]))X Where Pronostico <> 0)V left join(select C.IdAgen, C.Cod_Prod, sum(C.Convertidas) as EntregadoT FROM(select IdAgen, Cod_Prod, Cod_Campo, sum(Convertidas) as Convertidas FROM UV_ProdRecepcion where CodEstatus <> 'C' and Temporada = (select Temporada from CatSemanas where getdate() between Inicio and Fin) group by IdAgen, Cod_prod, Cod_Campo)C GROUP BY C.IdAgen, C.Cod_Prod)E ON V.IdAgen = E.IdAgen AND V.Cod_Prod = E.Cod_Prod group by V.IdAgen, V.Cod_prod, E.EntregadoT)V GROUP BY V.IdAgen, V.Cod_prod)V ON C.IdAgen = V.IdAgen and C.Cod_Prod = V.Cod_Prod left join(select V.IdAgen, V.Cod_prod, sum(isnull(V.Pronostico,0)) as PronosticoA FROM CatSemanas S left join(select * from(select V.IdAgen, V.Cod_Prod, max(V.Fecha) as fecha, V.[27], V.[28], V.[29], V.[30], V.[31], V.[32], V.[33], V.[34], V.[35], V.[36], V.[37], V.[38], V.[39], V.[40], V.[41], V.[42], V.[43], V.[44], V.[45], V.[46], V.[47], V.[48], V.[49], V.[50], V.[51], V.[52], V.[53], V.[1], V.[2], V.[3], V.[4], V.[5], V.[6], V.[7], V.[8], V.[9], V.[10], V.[11], V.[12], V.[13], V.[14], V.[15], V.[16], V.[17], V.[18], V.[19], V.[20], V.[21], V.[22], V.[23], V.[24], V.[25], V.[26] From(select distinct IdAgen, Cod_Prod, Fecha, SUM(isnull([27], 0)) as [27], sum(isnull([28], 0)) as [28], sum(isnull([29], 0)) as[29], sum(isnull([30], 0)) as [30], sum(isnull([31], 0)) as [31], sum(isnull([32], 0)) as [32], sum(isnull([33], 0)) as [33], sum(isnull([34], 0)) as [34], sum(isnull([35], 0)) as [35], sum(isnull([36], 0)) as [36], sum(isnull([37], 0)) as [37], sum(isnull([38], 0)) as [38], sum(isnull([39], 0)) as [39], sum(isnull([40], 0)) as [40], sum(isnull([41], 0)) as [41], sum(isnull([42], 0)) as [42], sum(isnull([43], 0)) as [43], sum(isnull([44], 0)) as [44], sum(isnull([45], 0)) as [45], sum(isnull([46], 0)) as [46], sum(isnull([47], 0)) as [47], sum(isnull([48], 0)) as [48], sum(isnull([49], 0)) as [49], sum(isnull([50], 0)) as [50], sum(isnull([51], 0)) as [51], sum(isnull([52], 0)) as [52],sum(isnull([53], 0)) as [53], sum(isnull([1], 0)) as [1], sum(isnull([2], 0)) as [2], sum(isnull([3], 0)) as [3], sum(isnull([4], 0)) as [4], sum(isnull([5], 0)) as [5], sum(isnull([6], 0)) as [6], sum(isnull([7], 0)) as [7], sum(isnull([8], 0)) as [8], sum(isnull([9], 0)) as [9], sum(isnull([10], 0)) as [10], sum(isnull([11], 0)) as [11], sum(isnull([12], 0)) as [12], sum(isnull([13], 0)) as [13], sum(isnull([14], 0)) as [14], sum(isnull([15], 0)) as [15],sum(isnull([16], 0)) as [16], sum(isnull([17], 0)) as [17], sum(isnull([18], 0)) as [18], sum(isnull([19], 0)) as [19], sum(isnull([20], 0)) as [20], sum(isnull([21], 0)) as [21], sum(isnull([22], 0)) as [22], sum(isnull([23], 0)) as [23], sum(isnull([24], 0)) as [24], sum(isnull([25], 0)) as [25], sum(isnull([26], 0)) as [26] from SIPGProyeccion P where Temporada = (select Temporada from CatSemanas where getdate() between Inicio and Fin) " +
                        "and Fecha = (select max(Fecha) from SIPGProyeccion where IdAgen = " + (short)Session["IdAgen"] + ") group by IdAgen, Cod_Prod,Fecha)V GROUP BY V.IdAgen, V.Cod_Prod,V.[27],V.[28],V.[29],V.[30],V.[31],V.[32],V.[33],V.[34],V.[35],V.[36],V.[37],V.[38],V.[39],V.[40],V.[41],V.[42],V.[43],V.[44],V.[45],V.[46],V.[47],V.[48],V.[49],V.[50],V.[51],V.[52],V.[53],V.[1],V.[2],V.[3],V.[4],V.[5],V.[6],V.[7],V.[8],V.[9],V.[10],V.[11],V.[12],V.[13],V.[14],V.[15],V.[16],V.[17],V.[18],V.[19],V.[20],V.[21],V.[22],V.[23],V.[24],V.[25],V.[26])V UNPIVOT(Pronostico FOR Semana in ([27],[28],[29],[30],[31],[32],[33],[34],[35],[36],[37],[38],[39],[40],[41],[42],[43],[44],[45],[46],[47],[48],[49],[50],[51],[52],[53],[1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],[17],[18],[19],[20],[21],[22],[23],[24],[25],[26]))X where Pronostico <> 0)V on S.Semana = V.Semana where S.Inicio between(select Inicio from catsemanas where temporada = (select Temporada from CatSemanas where getdate() between Inicio and Fin) and semana = 27) and getdate() group by V.IdAgen,V.Cod_prod)AA ON V.IdAgen = AA.IdAgen and V.Cod_Prod = AA.Cod_Prod LEFT JOIN(SELECT PA.IdAgen, PA.Cod_Prod, SUM(isnull(PA.Pronostico,0)) AS PronosticoSA, sum(isnull(EA.Entregado, 0)) AS EntregadoSA, S.Semana FROM CatSemanas S left join(select* from(select V.IdAgen, V.Cod_Prod, max(V.Fecha)as fecha, V.[27],V.[28],V.[29],V.[30],V.[31],V.[32],V.[33],V.[34],V.[35],V.[36],V.[37],V.[38],V.[39],V.[40],V.[41],V.[42],V.[43],V.[44],V.[45],V.[46],V.[47],V.[48],V.[49],V.[50],V.[51],V.[52],V.[53],V.[1],V.[2],V.[3],V.[4],V.[5],V.[6],V.[7],V.[8],V.[9],V.[10],V.[11],V.[12],V.[13],V.[14],V.[15],V.[16],V.[17],V.[18],V.[19],V.[20],V.[21],V.[22],V.[23],V.[24],V.[25],V.[26] From(select distinct IdAgen, Cod_Prod, Fecha, SUM(isnull([27], 0)) as [27], sum(isnull([28], 0)) as [28], sum(isnull([29], 0)) as[29], sum(isnull([30], 0)) as [30], sum(isnull([31], 0)) as [31], sum(isnull([32], 0)) as [32], sum(isnull([33], 0)) as [33], sum(isnull([34], 0)) as [34], sum(isnull([35], 0)) as [35], sum(isnull([36], 0)) as [36], sum(isnull([37], 0)) as [37], sum(isnull([38], 0)) as [38], sum(isnull([39], 0)) as [39], sum(isnull([40], 0)) as [40], sum(isnull([41], 0)) as [41], sum(isnull([42], 0)) as [42], sum(isnull([43], 0)) as [43], sum(isnull([44], 0)) as [44], sum(isnull([45], 0)) as [45], sum(isnull([46], 0)) as [46], sum(isnull([47], 0)) as [47], sum(isnull([48], 0)) as [48], sum(isnull([49], 0)) as [49], sum(isnull([50], 0)) as [50], sum(isnull([51], 0)) as [51], sum(isnull([52], 0)) as [52],sum(isnull([53], 0)) as [53], sum(isnull([1], 0)) as [1], sum(isnull([2], 0)) as [2], sum(isnull([3], 0)) as [3], sum(isnull([4], 0)) as [4], sum(isnull([5], 0)) as [5], sum(isnull([6], 0)) as [6], sum(isnull([7], 0)) as [7], sum(isnull([8], 0)) as [8], sum(isnull([9], 0)) as [9], sum(isnull([10], 0)) as [10], sum(isnull([11], 0)) as [11], sum(isnull([12], 0)) as [12], sum(isnull([13], 0)) as [13], sum(isnull([14], 0)) as [14], sum(isnull([15], 0)) as [15],sum(isnull([16], 0)) as [16], sum(isnull([17], 0)) as [17], sum(isnull([18], 0)) as [18], sum(isnull([19], 0)) as [19], sum(isnull([20], 0)) as [20], sum(isnull([21], 0)) as [21], sum(isnull([22], 0)) as [22], sum(isnull([23], 0)) as [23], sum(isnull([24], 0)) as [24], sum(isnull([25], 0)) as [25], sum(isnull([26], 0)) as [26] from SIPGProyeccion P where Temporada = (select Temporada from CatSemanas where getdate() between Inicio and Fin) " +
                        "and Fecha = (select max(Fecha) from SIPGProyeccion where IdAgen = " + (short)Session["IdAgen"] + ") group by IdAgen, Cod_Prod,Fecha)V GROUP BY V.IdAgen, V.Cod_Prod,V.[27],V.[28],V.[29],V.[30],V.[31],V.[32],V.[33],V.[34],V.[35],V.[36],V.[37],V.[38],V.[39],V.[40],V.[41],V.[42],V.[43],V.[44],V.[45],V.[46],V.[47],V.[48],V.[49],V.[50],V.[51],V.[52],V.[53],V.[1],V.[2],V.[3],V.[4],V.[5],V.[6],V.[7],V.[8],V.[9],V.[10],V.[11],V.[12],V.[13],V.[14],V.[15],V.[16],V.[17],V.[18],V.[19],V.[20],V.[21],V.[22],V.[23],V.[24],V.[25],V.[26])V UNPIVOT(Pronostico FOR Semana in ([27],[28],[29],[30],[31],[32],[33],[34],[35],[36],[37],[38],[39],[40],[41],[42],[43],[44],[45],[46],[47],[48],[49],[50],[51],[52],[53],[1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],[17],[18],[19],[20],[21],[22],[23],[24],[25],[26]))X where Pronostico <> 0)PA ON S.Semana = PA.Semana Left Join(select C.IdAgen, C.Cod_Prod, sum(C.Convertidas) as Entregado, C.Semana FROM(select IdAgen, Cod_Prod, sum(Convertidas) as Convertidas, Semana FROM UV_ProdRecepcion where CodEstatus <> 'C' and Temporada = (select Temporada from CatSemanas where getdate() between Inicio and Fin) group by IdAgen, Cod_prod, Cod_Campo, Semana)C GROUP BY C.IdAgen, C.Cod_Prod, C.Semana)EA ON PA.IdAgen = EA.IdAgen and PA.Cod_Prod = EA.Cod_Prod and S.Semana = EA.Semana WHERE getdate() between S.Inicio and S.Fin GROUP BY PA.IdAgen, PA.Cod_Prod, S.Semana)SA on V.IdAgen = SA.IdAgen AND V.Cod_Prod = SA.Cod_Prod group by C.IdAgen, C.Cod_Prod, P.Nombre, V.PronosticoT, V.EntregadoT, AA.PronosticoA, SA.Semana, SA.PronosticoSA, SA.EntregadoSA)V)V)C " +
                        "WHERE C.IdAgen = " + (short)Session["IdAgen"] + "  and C.Cod_Prod not in(Select distinct V.Cod_Prod FROM(select* from CatSemanas where Getdate() between Inicio and Fin)S left join SIPGVisitas V on V.Fecha between S.Inicio and S.Fin " +
                        "where V.IdAgen = " + (short)Session["IdAgen"] + ") and C.Semana is not NULL ORDER BY C.DiferenciaA DESC").ToList();
                }

            }
            else
            {
                if (agente != 0)
                {
                    curva = bd.Database.SqlQuery<ClassCurva>("select V.Cod_Prod, V.Nombre as Productor, V.PronosticoT, V.PronosticoA, V.EntregadoT, (CASE WHEN V.EntregadoT ='0' AND  V.PronosticoA ='0' THEN '' ELSE V.DiferenciaA END) AS DiferenciaA, V.Semana, V.PronosticoSA, V.EntregadoSA, (CASE WHEN V.EntregadoSA ='0' AND  V.PronosticoSA ='0' THEN '' ELSE V.DiferenciaSA END) AS DiferenciaSA  from(SELECT V.Cod_Prod, V.Nombre, (case when round(V.PronosticoT,0) is null then '0' else round(V.PronosticoT,0) end) as PronosticoT, (case when round(V.PronosticoA,0) is null then '0' else round(V.PronosticoA,0) end) as PronosticoA, (case when round(V.EntregadoT,0) is null then '0' else round(V.EntregadoT,0) end)as EntregadoT, (case when round(((V.DiferenciaA*100)-100),0) is null then '100' else round(((V.DiferenciaA*100)-100),0) end)as DiferenciaA, V.Semana, (case when round(V.PronosticoSA,0) is null then '0' else round(V.PronosticoSA,0) end) as PronosticoSA, (case when round(V.EntregadoSA,0) is null then '0' else round(V.EntregadoSA,0) end) as EntregadoSA, (case when round(((V.DiferenciaSA*100)-100),0) is null then '100' else round(((V.DiferenciaSA*100)-100),0) end) as DiferenciaSA FROM(select *,(V.EntregadoT/V.PronosticoA) AS DiferenciaA, (V.EntregadoSA / V.PronosticoSA) AS DiferenciaSA FROM(select C.Cod_Prod, P.Nombre, V.PronosticoT, V.EntregadoT, AA.PronosticoA, SA.Semana, SA.PronosticoSA, SA.EntregadoSA FROM ProdCamposCat C left join ProdProductoresCat P on C.Cod_Prod=P.Cod_Prod left join(select V.IdAgen, V.Cod_prod, sum(ISNULL(V.Pronostico,0)) as PronosticoT , sum(ISNULL(V.EntregadoT,0)) AS EntregadoT from(select V.IdAgen, V.Cod_prod, sum(V.Pronostico) as Pronostico, E.EntregadoT from(select * from(select V.IdAgen, V.Cod_Prod,max(V.Fecha)as fecha, V.[27],V.[28],V.[29],V.[30],V.[31],V.[32],V.[33],V.[34],V.[35],V.[36],V.[37],V.[38],V.[39],V.[40],V.[41],V.[42],V.[43],V.[44],V.[45],V.[46],V.[47],V.[48],V.[49],V.[50],V.[51],V.[52],V.[53],V.[1],V.[2],V.[3],V.[4],V.[5],V.[6],V.[7],V.[8],V.[9],V.[10],V.[11],V.[12],V.[13],V.[14],V.[15],V.[16],V.[17],V.[18],V.[19],V.[20],V.[21],V.[22],V.[23],V.[24],V.[25],V.[26] From(select distinct IdAgen, Cod_Prod, Fecha, SUM(isnull([27], 0)) as [27], sum(isnull([28], 0)) as [28], sum(isnull([29], 0)) as[29], sum(isnull([30], 0)) as [30], sum(isnull([31], 0)) as [31], sum(isnull([32], 0)) as [32], sum(isnull([33], 0)) as [33], sum(isnull([34], 0)) as [34], sum(isnull([35], 0)) as [35], sum(isnull([36], 0)) as [36], sum(isnull([37], 0)) as [37], sum(isnull([38], 0)) as [38], sum(isnull([39], 0)) as [39], sum(isnull([40], 0)) as [40], sum(isnull([41], 0)) as [41], sum(isnull([42], 0)) as [42], sum(isnull([43], 0)) as [43], sum(isnull([44], 0)) as [44], sum(isnull([45], 0)) as [45], sum(isnull([46], 0)) as [46], sum(isnull([47], 0)) as [47], sum(isnull([48], 0)) as [48], sum(isnull([49], 0)) as [49], sum(isnull([50], 0)) as [50], sum(isnull([51], 0)) as [51], sum(isnull([52], 0)) as [52], sum(isnull([53], 0)) as [53], sum(isnull([1], 0)) as [1], sum(isnull([2], 0)) as [2], sum(isnull([3], 0)) as [3], sum(isnull([4], 0)) as [4], sum(isnull([5], 0)) as [5], sum(isnull([6], 0)) as [6], sum(isnull([7], 0)) as [7], sum(isnull([8], 0)) as [8], sum(isnull([9], 0)) as [9], sum(isnull([10], 0)) as [10], sum(isnull([11], 0)) as [11], sum(isnull([12], 0)) as [12], sum(isnull([13], 0)) as [13], sum(isnull([14], 0)) as [14], sum(isnull([15], 0)) as [15],sum(isnull([16], 0)) as [16], sum(isnull([17], 0)) as [17], sum(isnull([18], 0)) as [18], sum(isnull([19], 0)) as [19], sum(isnull([20], 0)) as [20], sum(isnull([21], 0)) as [21], sum(isnull([22], 0)) as [22], sum(isnull([23], 0)) as [23], sum(isnull([24], 0)) as [24], sum(isnull([25], 0)) as [25], sum(isnull([26], 0)) as [26] from SIPGProyeccion P where Temporada = (select Temporada from CatSemanas where getdate() between Inicio and Fin) and Fecha=(select max(Fecha) from SIPGProyeccion) group by IdAgen, Cod_Prod,Fecha)V GROUP BY V.IdAgen, V.Cod_Prod,V.[27],V.[28],V.[29],V.[30],V.[31],V.[32],V.[33],V.[34],V.[35],V.[36],V.[37],V.[38],V.[39],V.[40],V.[41],V.[42],V.[43],V.[44],V.[45],V.[46],V.[47],V.[48],V.[49],V.[50],V.[51],V.[52],V.[53],V.[1],V.[2],V.[3],V.[4],V.[5],V.[6],V.[7],V.[8],V.[9],V.[10],V.[11],V.[12],V.[13],V.[14],V.[15],V.[16],V.[17],V.[18],V.[19],V.[20],V.[21],V.[22],V.[23],V.[24],V.[25],V.[26])V UNPIVOT(Pronostico FOR Semana in ([27],[28],[29],[30],[31],[32],[33],[34],[35],[36],[37],[38],[39],[40],[41],[42],[43],[44],[45],[46],[47],[48],[49],[50],[51],[52],[53],[1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],[17],[18],[19],[20],[21],[22],[23],[24],[25],[26]))X Where Pronostico <> 0)V left join(select C.IdAgen, C.Cod_Prod,sum(C.Convertidas) as EntregadoT FROM(select IdAgen, Cod_Prod, Cod_Campo,sum(Convertidas) as Convertidas FROM UV_ProdRecepcion where CodEstatus <> 'C' and Temporada = (select Temporada from CatSemanas where getdate() between Inicio and Fin) group by IdAgen, Cod_prod,Cod_Campo)C GROUP BY C.IdAgen, C.Cod_Prod)E ON V.IdAgen = E.IdAgen AND V.Cod_Prod = E.Cod_Prod group by V.IdAgen, V.Cod_prod, E.EntregadoT)V GROUP BY V.IdAgen, V.Cod_prod)V ON C.IdAgen = V.IdAgen and C.Cod_Prod = V.Cod_Prod left join(select V.IdAgen, V.Cod_prod, sum(isnull(V.Pronostico,0)) as PronosticoA FROM CatSemanas S left join(select * from(select V.IdAgen, V.Cod_Prod,max(V.Fecha)as fecha, V.[27],V.[28],V.[29],V.[30],V.[31],V.[32],V.[33],V.[34],V.[35],V.[36],V.[37],V.[38],V.[39],V.[40],V.[41],V.[42],V.[43],V.[44],V.[45],V.[46],V.[47],V.[48],V.[49],V.[50],V.[51],V.[52],V.[53],V.[1],V.[2],V.[3],V.[4],V.[5],V.[6],V.[7],V.[8],V.[9],V.[10],V.[11],V.[12],V.[13],V.[14],V.[15],V.[16],V.[17],V.[18],V.[19],V.[20],V.[21],V.[22],V.[23],V.[24],V.[25],V.[26] From(select distinct IdAgen, Cod_Prod, Fecha, SUM(isnull([27], 0)) as [27], sum(isnull([28], 0)) as [28], sum(isnull([29], 0)) as[29], sum(isnull([30], 0)) as [30], sum(isnull([31], 0)) as [31], sum(isnull([32], 0)) as [32], sum(isnull([33], 0)) as [33], sum(isnull([34], 0)) as [34], sum(isnull([35], 0)) as [35], sum(isnull([36], 0)) as [36], sum(isnull([37], 0)) as [37], sum(isnull([38], 0)) as [38], sum(isnull([39], 0)) as [39], sum(isnull([40], 0)) as [40], sum(isnull([41], 0)) as [41], sum(isnull([42], 0)) as [42], sum(isnull([43], 0)) as [43], sum(isnull([44], 0)) as [44], sum(isnull([45], 0)) as [45], sum(isnull([46], 0)) as [46], sum(isnull([47], 0)) as [47], sum(isnull([48], 0)) as [48], sum(isnull([49], 0)) as [49], sum(isnull([50], 0)) as [50], sum(isnull([51], 0)) as [51], sum(isnull([52], 0)) as [52],sum(isnull([53], 0)) as [53], sum(isnull([1], 0)) as [1], sum(isnull([2], 0)) as [2], sum(isnull([3], 0)) as [3], sum(isnull([4], 0)) as [4], sum(isnull([5], 0)) as [5], sum(isnull([6], 0)) as [6], sum(isnull([7], 0)) as [7], sum(isnull([8], 0)) as [8], sum(isnull([9], 0)) as [9], sum(isnull([10], 0)) as [10], sum(isnull([11], 0)) as [11], sum(isnull([12], 0)) as [12], sum(isnull([13], 0)) as [13], sum(isnull([14], 0)) as [14], sum(isnull([15], 0)) as [15],sum(isnull([16], 0)) as [16], sum(isnull([17], 0)) as [17], sum(isnull([18], 0)) as [18], sum(isnull([19], 0)) as [19], sum(isnull([20], 0)) as [20], sum(isnull([21], 0)) as [21], sum(isnull([22], 0)) as [22], sum(isnull([23], 0)) as [23], sum(isnull([24], 0)) as [24], sum(isnull([25], 0)) as [25], sum(isnull([26], 0)) as [26] from SIPGProyeccion P where Temporada = (select Temporada from CatSemanas where getdate() between Inicio and Fin) and Fecha=(select max(Fecha) from SIPGProyeccion) group by IdAgen, Cod_Prod,Fecha)V GROUP BY V.IdAgen, V.Cod_Prod,V.[27],V.[28],V.[29],V.[30],V.[31],V.[32],V.[33],V.[34],V.[35],V.[36],V.[37],V.[38],V.[39],V.[40],V.[41],V.[42],V.[43],V.[44],V.[45],V.[46],V.[47],V.[48],V.[49],V.[50],V.[51],V.[52],V.[53],V.[1],V.[2],V.[3],V.[4],V.[5],V.[6],V.[7],V.[8],V.[9],V.[10],V.[11],V.[12],V.[13],V.[14],V.[15],V.[16],V.[17],V.[18],V.[19],V.[20],V.[21],V.[22],V.[23],V.[24],V.[25],V.[26])V UNPIVOT(Pronostico FOR Semana in ([27],[28],[29],[30],[31],[32],[33],[34],[35],[36],[37],[38],[39],[40],[41],[42],[43],[44],[45],[46],[47],[48],[49],[50],[51],[52],[53],[1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],[17],[18],[19],[20],[21],[22],[23],[24],[25],[26]))X where Pronostico <> 0)V on S.Semana = V.Semana where S.Inicio between (select Inicio from catsemanas where temporada =(select Temporada from CatSemanas where getdate() between Inicio and Fin) and semana=27) and getdate() group by V.IdAgen,V.Cod_prod)AA ON V.IdAgen = AA.IdAgen and V.Cod_Prod = AA.Cod_Prod LEFT JOIN(SELECT PA.IdAgen, PA.Cod_Prod, SUM(isnull(PA.Pronostico,0)) AS PronosticoSA, sum(isnull(EA.Entregado, 0)) AS EntregadoSA, S.Semana FROM CatSemanas S left join(select * from(select V.IdAgen, V.Cod_Prod,max(V.Fecha)as fecha, V.[27],V.[28],V.[29],V.[30],V.[31],V.[32],V.[33],V.[34],V.[35],V.[36],V.[37],V.[38],V.[39],V.[40],V.[41],V.[42],V.[43],V.[44],V.[45],V.[46],V.[47],V.[48],V.[49],V.[50],V.[51],V.[52],V.[53],V.[1],V.[2],V.[3],V.[4],V.[5],V.[6],V.[7],V.[8],V.[9],V.[10],V.[11],V.[12],V.[13],V.[14],V.[15],V.[16],V.[17],V.[18],V.[19],V.[20],V.[21],V.[22],V.[23],V.[24],V.[25],V.[26] From(select distinct IdAgen, Cod_Prod, Fecha, SUM(isnull([27], 0)) as [27], sum(isnull([28], 0)) as [28], sum(isnull([29], 0)) as[29], sum(isnull([30], 0)) as [30], sum(isnull([31], 0)) as [31], sum(isnull([32], 0)) as [32], sum(isnull([33], 0)) as [33], sum(isnull([34], 0)) as [34], sum(isnull([35], 0)) as [35], sum(isnull([36], 0)) as [36], sum(isnull([37], 0)) as [37], sum(isnull([38], 0)) as [38], sum(isnull([39], 0)) as [39], sum(isnull([40], 0)) as [40], sum(isnull([41], 0)) as [41], sum(isnull([42], 0)) as [42], sum(isnull([43], 0)) as [43], sum(isnull([44], 0)) as [44], sum(isnull([45], 0)) as [45], sum(isnull([46], 0)) as [46], sum(isnull([47], 0)) as [47], sum(isnull([48], 0)) as [48], sum(isnull([49], 0)) as [49], sum(isnull([50], 0)) as [50], sum(isnull([51], 0)) as [51], sum(isnull([52], 0)) as [52],sum(isnull([53], 0)) as [53], sum(isnull([1], 0)) as [1], sum(isnull([2], 0)) as [2], sum(isnull([3], 0)) as [3], sum(isnull([4], 0)) as [4], sum(isnull([5], 0)) as [5], sum(isnull([6], 0)) as [6], sum(isnull([7], 0)) as [7], sum(isnull([8], 0)) as [8], sum(isnull([9], 0)) as [9], sum(isnull([10], 0)) as [10], sum(isnull([11], 0)) as [11], sum(isnull([12], 0)) as [12], sum(isnull([13], 0)) as [13], sum(isnull([14], 0)) as [14], sum(isnull([15], 0)) as [15],sum(isnull([16], 0)) as [16], sum(isnull([17], 0)) as [17], sum(isnull([18], 0)) as [18], sum(isnull([19], 0)) as [19], sum(isnull([20], 0)) as [20], sum(isnull([21], 0)) as [21], sum(isnull([22], 0)) as [22], sum(isnull([23], 0)) as [23], sum(isnull([24], 0)) as [24], sum(isnull([25], 0)) as [25], sum(isnull([26], 0)) as [26] from SIPGProyeccion P where Temporada = (select Temporada from CatSemanas where getdate() between Inicio and Fin)  and Fecha=(select max(Fecha) from SIPGProyeccion) group by IdAgen, Cod_Prod,Fecha)V GROUP BY V.IdAgen, V.Cod_Prod,V.[27],V.[28],V.[29],V.[30],V.[31],V.[32],V.[33],V.[34],V.[35],V.[36],V.[37],V.[38],V.[39],V.[40],V.[41],V.[42],V.[43],V.[44],V.[45],V.[46],V.[47],V.[48],V.[49],V.[50],V.[51],V.[52],V.[53],V.[1],V.[2],V.[3],V.[4],V.[5],V.[6],V.[7],V.[8],V.[9],V.[10],V.[11],V.[12],V.[13],V.[14],V.[15],V.[16],V.[17],V.[18],V.[19],V.[20],V.[21],V.[22],V.[23],V.[24],V.[25],V.[26] )V UNPIVOT(Pronostico FOR Semana in ([27],[28],[29],[30],[31],[32],[33],[34],[35],[36],[37],[38],[39],[40],[41],[42],[43],[44],[45],[46],[47],[48],[49],[50],[51],[52],[53],[1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],[17],[18],[19],[20],[21],[22],[23],[24],[25],[26]))X where Pronostico <> 0)PA ON S.Semana = PA.Semana Left Join(select C.IdAgen, C.Cod_Prod, sum(C.Convertidas) as Entregado, C.Semana FROM(select IdAgen, Cod_Prod, sum(Convertidas) as Convertidas, Semana FROM UV_ProdRecepcion where CodEstatus<> 'C' and Temporada = (select Temporada from CatSemanas where getdate() between Inicio and Fin) group by IdAgen, Cod_prod, Cod_Campo, Semana)C GROUP BY C.IdAgen, C.Cod_Prod, C.Semana)EA ON PA.IdAgen = EA.IdAgen and PA.Cod_Prod = EA.Cod_Prod and S.Semana = EA.Semana WHERE getdate() between S.Inicio and S.Fin GROUP BY PA.IdAgen, PA.Cod_Prod, S.Semana)SA on V.IdAgen = SA.IdAgen AND V.Cod_Prod = SA.Cod_Prod " +
                       "WHERE V.IdAgen = " + agente + " group by C.Cod_Prod, P.Nombre, V.PronosticoT, V.EntregadoT, AA.PronosticoA, SA.Semana, SA.PronosticoSA, SA.EntregadoSA)V)V)V " +
                       "ORDER BY DiferenciaA DESC").ToList(); 
                }
            }
            return Json(curva, JsonRequestBehavior.AllowGet);
        }

        public ActionResult CurvaNueva()
        {
            if (Session["Nombre"] != null)
            {
                ViewData["Nombre"] = Session["Nombre"].ToString();
                var List_Cultivos = new SelectList(new[]
                    {
                    new { Value = "0", Text = "--Cultivo--" },
                    new { Value = "1", Text = "ZARZAMORA" },
                    new { Value = "2", Text = "FRAMBUESA" },
                    new { Value = "3", Text = "ARANDANO" },
                    new { Value = "3", Text = "FRESA" },
                },
                    "Value", "Text", 0);

                var List_Manejo = new SelectList(new[]
              {
                    new { Value = "0", Text = "--Manejo--" },
                    new { Value = "1", Text = "CONVENCIONAL" },
                    new { Value = "2", Text = "ORGANICO" },
                },
                "Value", "Text", 0);

                var List_Num_corte = new SelectList(new[]
              {
                    new { Value = "0", Text = "--" },
                    new { Value = "1", Text = "1" },
                    new { Value = "2", Text = "2" },
                },
                "Value", "Text", 0);

                var List_Tipo_plantacion = new SelectList(new[]
             {
                    new { Value = "0", Text = "--" },
                    new { Value = "1", Text = "ENCAMADO" },
                    new { Value = "2", Text = "HIDROPONIA" },
                },
                "Value", "Text", 0);

                var List_Estructura = new SelectList(new[]
           {
                    new { Value = "0", Text = "--" },
                    new { Value = "1", Text = "TUNEL" },
                    new { Value = "2", Text = "MALLA SOMBRA" },
                    new { Value = "3", Text = "SIN TUNEL" },
                },
                "Value", "Text", 0);

                var List_Edad_planta = new SelectList(new[]
           {
                    new { Value = "0", Text = "--" },
                    new { Value = "1", Text = "MENOR A UN AÑO" },
                    new { Value = "2", Text = "1 AÑO" },
                    new { Value = "2", Text = "2 AÑOS" },
                    new { Value = "2", Text = "3 AÑOS" },
                    new { Value = "2", Text = "MAYOR A 3 AÑOS" },
                },
                "Value", "Text", 0);

                var List_Plantacion = new SelectList(new[]
           {
                    new { Value = "0", Text = "--" },
                    new { Value = "1", Text = "1ER AÑO" },
                    new { Value = "2", Text = "MAS DE UN AÑO" },
                },
                "Value", "Text", 0);

                var List_Tesco = new SelectList(new[]
                {
                    new { Value = "0", Text = "--" },
                    new { Value = "2", Text = "NO" },
                    new { Value = "1", Text = "SI" },
                },
                "Value", "Text", 0);

                ViewData["List_Cultivos"] = List_Cultivos;
                ViewData["List_Manejo"] = List_Manejo;
                ViewData["List_Num_corte"] = List_Num_corte;
                ViewData["List_Tipo_plantacion"] = List_Tipo_plantacion;
                ViewData["List_Estructura"] = List_Estructura;
                ViewData["List_Edad_planta"] = List_Edad_planta;
                ViewData["List_Plantacion"] = List_Plantacion;
                ViewData["List_Tesco"] = List_Tesco;
            }
            else
            {
                return RedirectToAction("Index", "Login");
            }
            return View();
        }

        [HttpPost]
        public ActionResult CurvaNueva(SIPGProyeccion model, string cultivo)
        {
            try
            {
                if (Session["Nombre"] != null)
                {
                    ViewData["Nombre"] = Session["Nombre"].ToString();

                    //     var List_Cultivos = new SelectList(new[]
                    //     {
                    //     new { Value = "0", Text = "--Cultivo--" },
                    //     new { Value = "1", Text = "ZARZAMORA" },
                    //     new { Value = "2", Text = "FRAMBUESA" },
                    //     new { Value = "3", Text = "ARANDANO" },
                    //     new { Value = "3", Text = "FRESA" },
                    // },
                    //     "Value", "Text", 0);

                    //     var List_Manejo = new SelectList(new[]
                    //   {
                    //     new { Value = "0", Text = "--Manejo--" },
                    //     new { Value = "1", Text = "CONVENCIONAL" },
                    //     new { Value = "2", Text = "ORGANICO" },
                    // },
                    //     "Value", "Text", 0);

                    //     var List_Num_corte = new SelectList(new[]
                    //   {
                    //     new { Value = "0", Text = "--" },
                    //     new { Value = "1", Text = "1" },
                    //     new { Value = "2", Text = "2" },
                    // },
                    //     "Value", "Text", 0);

                    //     var List_Tipo_plantacion = new SelectList(new[]
                    //  {
                    //     new { Value = "0", Text = "--" },
                    //     new { Value = "1", Text = "ENCAMADO" },
                    //     new { Value = "2", Text = "HIDROPONIA" },
                    // },
                    //     "Value", "Text", 0);

                    //     var List_Estructura = new SelectList(new[]
                    //{
                    //     new { Value = "0", Text = "--" },
                    //     new { Value = "1", Text = "TUNEL" },
                    //     new { Value = "2", Text = "MALLA SOMBRA" },
                    //     new { Value = "3", Text = "SIN TUNEL" },
                    // },
                    //     "Value", "Text", 0);

                    //     var List_Edad_planta = new SelectList(new[]
                    //{
                    //     new { Value = "0", Text = "--" },
                    //     new { Value = "1", Text = "MENOR A UN AÑO" },
                    //     new { Value = "2", Text = "1 AÑO" },
                    //     new { Value = "2", Text = "2 AÑOS" },
                    //     new { Value = "2", Text = "3 AÑOS" },
                    //     new { Value = "2", Text = "MAYOR A 3 AÑOS" },
                    // },
                    //     "Value", "Text", 0);

                    //     var List_Plantacion = new SelectList(new[]
                    //{
                    //     new { Value = "0", Text = "--" },
                    //     new { Value = "1", Text = "1ER AÑO" },
                    //     new { Value = "2", Text = "MAS DE UN AÑO" },
                    // },
                    //     "Value", "Text", 0);

                    //     var List_Tesco = new SelectList(new[]
                    //     {
                    //     new { Value = "0", Text = "--" },
                    //     new { Value = "2", Text = "NO" },
                    //     new { Value = "1", Text = "SI" },
                    // },
                    //     "Value", "Text", 0);

                    //     ViewData["List_Cultivos"] = List_Cultivos;
                    //     ViewData["List_Manejo"] = List_Manejo;
                    //     ViewData["List_Num_corte"] = List_Num_corte;
                    //     ViewData["List_Tipo_plantacion"] = List_Tipo_plantacion;
                    //     ViewData["List_Estructura"] = List_Estructura;
                    //     ViewData["List_Edad_planta"] = List_Edad_planta;
                    //     ViewData["List_Plantacion"] = List_Plantacion;
                    //     ViewData["List_Tesco"] = List_Tesco;

                    var modeloExistente = bd.SIPGProyeccion.FirstOrDefault(m => m.Cod_Prod == model.Cod_Prod && m.Cod_Campo == model.Cod_Campo && m.Sector == model.Sector);

                    if (modeloExistente == null)
                    {

                        int idagen = (short)Session["IdAgen"];
                        var res = bd.ProdAgenteCat.FirstOrDefault(m => m.IdAgen == idagen);
                        short IdZona = res.IdRegion;

                        var semanas1 = bd.CatSemanas.FirstOrDefault(m => model.Fecha_corte1 >= m.Inicio && model.Fecha_corte1 <= m.Fin);
                        int sem1 = semanas1.Semana;

                        var semanas2 = bd.CatSemanas.FirstOrDefault(m => model.Fecha_corte2 >= m.Inicio && model.Fecha_corte2 <= m.Fin);
                        int sem2 = semanas2.Semana;

                        SIPGProyeccion.Estado = "A";
                        SIPGProyeccion.IdAgen = (short)Session["IdAgen"];
                        SIPGProyeccion.Cod_Prod = model.Cod_Prod;
                        SIPGProyeccion.Cod_Campo = model.Cod_Campo;
                        SIPGProyeccion.Ubicacion = model.Ubicacion;
                        SIPGProyeccion.Num_corte = model.Num_corte;
                        SIPGProyeccion.Sector = model.Sector;
                        SIPGProyeccion.Ha = model.Ha;
                        SIPGProyeccion.Numplantas_xha = model.Numplantas_xha;
                        SIPGProyeccion.Manejo = model.Manejo;
                        SIPGProyeccion.Tipo_plantacion = model.Tipo_plantacion;
                        SIPGProyeccion.Fecha_plantacion = model.Fecha_plantacion;
                        SIPGProyeccion.Fecha_poda = model.Fecha_poda;
                        SIPGProyeccion.Fecha_defoliacion = model.Fecha_defoliacion;
                        SIPGProyeccion.Fecha_corte1 = model.Fecha_corte1;
                        SIPGProyeccion.Fecha_redefoliacion = model.Fecha_redefoliacion;
                        SIPGProyeccion.Fecha_corte2 = model.Fecha_corte2;
                        SIPGProyeccion.Sem1 = sem1;
                        SIPGProyeccion.Plantacion = model.Plantacion;
                        SIPGProyeccion.Sem2 = sem2;
                        SIPGProyeccion.Caja1 = model.Caja1;
                        SIPGProyeccion.Caja2 = model.Caja2;
                        SIPGProyeccion.Estructura = model.Estructura;
                        SIPGProyeccion.Tipo_certificacion = model.Tipo_certificacion;
                        SIPGProyeccion.Tesco = model.Tesco;
                        SIPGProyeccion.Edad_planta = model.Edad_planta;
                        SIPGProyeccion.Tipo_plantacion2 = model.Tipo_plantacion2;
                        SIPGProyeccion.Fecha_podamediacaña = model.Fecha_podamediacaña;
                        SIPGProyeccion.RendxKg = model.RendxKg;
                        SIPGProyeccion.RendxHa = model.RendxHa;
                        SIPGProyeccion.Temporada = "1920";
                        SIPGProyeccion.Zona = IdZona;
                        SIPGProyeccion.Acopio = model.Acopio;
                        SIPGProyeccion.Tipo = model.Tipo;
                        SIPGProyeccion.Producto = model.Producto;

                        bd.SIPGProyeccion.Add(model);

                        SIPGVisitas.IdAgen = (short)Session["IdAgen"];
                        SIPGVisitas.Fecha = DateTime.Now;
                        SIPGVisitas.Cod_Prod = SIPGVisitas.Cod_Prod;
                        SIPGVisitas.Cod_Campo = (short)SIPGVisitas.Cod_Campo;
                        SIPGVisitas.Sector = (short)SIPGVisitas.Sector;
                        bd.SIPGVisitas.Add(SIPGVisitas);

                        bd.SaveChanges();
                        TempData["sms"] = "Datos guardados con éxito";
                        ViewBag.sms = TempData["sms"].ToString();

                    }
                    else
                    {
                        TempData["sms"] = "La solicitud ya fue realizada anteriormente, verifique porfavor";
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

        //Lista de campos
        public JsonResult GetCamposList(string Cod_Prod)
        {
            bd.Configuration.ProxyCreationEnabled = false;
            List<ProdCamposCat> ListCampos = bd.ProdCamposCat.Where(x => x.Cod_Prod == Cod_Prod).ToList();
            return Json(ListCampos, JsonRequestBehavior.AllowGet);
        }

        //Ubicacion del campo
        public JsonResult GetUbicacion_campo(short Cod_Campo)
        {
            bd.Configuration.ProxyCreationEnabled = false;
            ProdCamposCat ubicacion = bd.ProdCamposCat.Where(x => x.Cod_Campo == Cod_Campo).FirstOrDefault();
            return Json(ubicacion, JsonRequestBehavior.AllowGet);
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

        //Acopios
        public JsonResult GetAcopios()
        {
            //int idagen = (short)Session["IdAgen"];
            //var res = bd.ProdAgenteCat.FirstOrDefault(m => m.IdAgen == idagen);
            //int IdZona = res.IdRegion;

            bd.Configuration.ProxyCreationEnabled = false;
            List<CatAcopios> Listacopios = bd.CatAcopios.OrderBy(x => x.Acopio).ToList(); //Where(x => x.IdZona == IdZona)
            return Json(Listacopios, JsonRequestBehavior.AllowGet);
        }

        //curva x zona
        public JsonResult CurvaListZ()
        {
            string x = "";
            if (Session["IdAgen"].ToString() == "1")
            {
                x = "1, 3, 4, 5";
            }
            else
            {
                x = "2";
            }
            List<ClassCurva> curvazona = bd.Database.SqlQuery<ClassCurva>("SELECT V.IdRegion,V.Zona,round(V.Pronostico,0)as Pronostico,round(V.PronosticoAA,0)as PronosticoAA,round(V.Entregado,0)as Entregado, round((((V.Entregado/V.PronosticoAA)*100)-100),0) AS DiferenciaAA,round(V.PronosticoSA,0)as PronosticoSA,round(V.EntregadoSA,0)as EntregadoSA, round((((V.EntregadoSA/V.PronosticoSA)*100)-100),0) AS DiferenciaSA FROM(select P.Zona as IdRegion, Z.Descripcion AS Zona, SUM(isnull(P.Pronostico,0)) AS Pronostico, SUM(isnull(MA.Pronostico,0)) AS PronosticoAA, SUM(isnull(E.Entregadas,0)) as Entregado, sum(isnull(SA.Pronostico,0)) as PronosticoSA, sum(isnull(SA.Entregadas,0)) as EntregadoSA from tbZonasAgricolas Z LEFT JOIN(select V.IdAgen, V.Pronostico, V.Zona FROM(select V.Zona, V.IdAgen, sum(V.Pronostico) as Pronostico from(select * from(select Zona, IdAgen, SUM(isnull([27], 0)) as [27], sum(isnull([28], 0)) as [28], sum(isnull([29], 0)) as[29], sum(isnull([30], 0)) as [30], sum(isnull([31], 0)) as [31], sum(isnull([32], 0)) as [32], sum(isnull([33], 0)) as [33], sum(isnull([34], 0)) as [34], sum(isnull([35], 0)) as [35], sum(isnull([36], 0)) as [36], sum(isnull([37], 0)) as [37], sum(isnull([38], 0)) as [38], sum(isnull([39], 0)) as [39], sum(isnull([40], 0)) as [40], sum(isnull([41], 0)) as [41], sum(isnull([42], 0)) as [42], sum(isnull([43], 0)) as [43], sum(isnull([44], 0)) as [44], sum(isnull([45], 0)) as [45], sum(isnull([46], 0)) as [46], sum(isnull([47], 0)) as [47], sum(isnull([48], 0)) as [48], sum(isnull([49], 0)) as [49], sum(isnull([50], 0)) as [50], sum(isnull([51], 0)) as [51], sum(isnull([52], 0)) as [52], sum(isnull([1], 0)) as [1], sum(isnull([2], 0)) as [2], sum(isnull([3], 0)) as [3], sum(isnull([4], 0)) as [4], sum(isnull([5], 0)) as [5], sum(isnull([6], 0)) as [6], sum(isnull([7], 0)) as [7], sum(isnull([8], 0)) as [8], sum(isnull([9], 0)) as [9], sum(isnull([10], 0)) as [10], sum(isnull([11], 0)) as [11], sum(isnull([12], 0)) as [12], sum(isnull([13], 0)) as [13], sum(isnull([14], 0)) as [14], sum(isnull([15], 0)) as [15],sum(isnull([16], 0)) as [16], sum(isnull([17], 0)) as [17], sum(isnull([18], 0)) as [18], sum(isnull([19], 0)) as [19], sum(isnull([20], 0)) as [20], sum(isnull([21], 0)) as [21], sum(isnull([22], 0)) as [22], sum(isnull([23], 0)) as [23], sum(isnull([24], 0)) as [24], sum(isnull([25], 0)) as [25], sum(isnull([26], 0)) as [26] " +
                "from SIPGProyeccion P where Temporada = (select Temporada from CatSemanas where getdate() between Inicio and Fin) and Estado = 'A' group by IdAgen, Zona)V UNPIVOT(Pronostico FOR Semana in ([27],[28],[29],[30],[31],[32],[33],[34],[35],[36],[37],[38],[39],[40],[41],[42],[43],[44],[45],[46],[47],[48],[49],[50],[51],[52],[1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],[17],[18],[19],[20],[21],[22],[23],[24],[25],[26]))X)V where Pronostico <> 0 group by V.Zona, V.IdAgen)V)P ON P.Zona = Z.CodZona left join(select C.IdAgen, sum(C.Convertidas) as Entregadas FROM(select IdAgen, sum(Convertidas) as Convertidas FROM SEasonsun1213..UV_ProdRecepcion where CodEstatus <> 'C' and Temporada = (select Temporada from CatSemanas where getdate() between Inicio and Fin) group by IdAgen Union All SELECT IdAgen, sum(Convertidas) as Convertidas FROM SeasonPlan..UV_ProdRecepcion where CodEstatus <> 'C' and Temporada = (select Temporada from CatSemanas where getdate() between Inicio and Fin) group by IdAgen)C GROUP BY C.IdAgen)E on P.IdAgen = E.IdAgen LEFT JOIN(select P.IdAgen, sum(P.Pronostico) as Pronostico FROM CatSemanas S LEFT JOIN(select * from(select * from(select IdAgen, SUM(isnull([27], 0)) as [27], sum(isnull([28], 0)) as [28], sum(isnull([29], 0)) as[29], sum(isnull([30], 0)) as [30], sum(isnull([31], 0)) as [31], sum(isnull([32], 0)) as [32], sum(isnull([33], 0)) as [33], sum(isnull([34], 0)) as [34], sum(isnull([35], 0)) as [35], sum(isnull([36], 0)) as [36], sum(isnull([37], 0)) as [37], sum(isnull([38], 0)) as [38], sum(isnull([39], 0)) as [39], sum(isnull([40], 0)) as [40], sum(isnull([41], 0)) as [41], sum(isnull([42], 0)) as [42], sum(isnull([43], 0)) as [43], sum(isnull([44], 0)) as [44], sum(isnull([45], 0)) as [45], sum(isnull([46], 0)) as [46], sum(isnull([47], 0)) as [47], sum(isnull([48], 0)) as [48], sum(isnull([49], 0)) as [49], sum(isnull([50], 0)) as [50], sum(isnull([51], 0)) as [51], sum(isnull([52], 0)) as [52], sum(isnull([1], 0)) as [1], sum(isnull([2], 0)) as [2], sum(isnull([3], 0)) as [3], sum(isnull([4], 0)) as [4], sum(isnull([5], 0)) as [5], sum(isnull([6], 0)) as [6], sum(isnull([7], 0)) as [7], sum(isnull([8], 0)) as [8], sum(isnull([9], 0)) as [9], sum(isnull([10], 0)) as [10], sum(isnull([11], 0)) as [11], sum(isnull([12], 0)) as [12], sum(isnull([13], 0)) as [13], sum(isnull([14], 0)) as [14], sum(isnull([15], 0)) as [15],sum(isnull([16], 0)) as [16], sum(isnull([17], 0)) as [17], sum(isnull([18], 0)) as [18], sum(isnull([19], 0)) as [19], sum(isnull([20], 0)) as [20], sum(isnull([21], 0)) as [21], sum(isnull([22], 0)) as [22], sum(isnull([23], 0)) as [23], sum(isnull([24], 0)) as [24], sum(isnull([25], 0)) as [25], sum(isnull([26], 0)) as [26] " +
                "from SIPGProyeccion P where Temporada = (select Temporada from CatSemanas where getdate() between Inicio and Fin) and Estado = 'A' group by IdAgen)V UNPIVOT(Pronostico FOR Semana in ([27],[28],[29],[30],[31],[32],[33],[34],[35],[36],[37],[38],[39],[40],[41],[42],[43],[44],[45],[46],[47],[48],[49],[50],[51],[52],[1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],[17],[18],[19],[20],[21],[22],[23],[24],[25],[26]))X)V where Pronostico <> 0)P on S.Semana = P.Semana " +
                "WHERE S.Inicio between '20190701' and getdate() AND S.Temporada = (select Temporada from CatSemanas where getdate() between Inicio and Fin) group by P.IdAgen)MA ON P.IdAgen = MA.IdAgen LEFT JOIN(SELECT PA.IdAgen, round(PA.Pronostico,0) AS Pronostico, EA.Entregadas FROM CatSemanas S left join(select* from(select* from(select IdAgen, SUM(isnull([27], 0)) as [27], sum(isnull([28], 0)) as [28], sum(isnull([29], 0)) as[29], sum(isnull([30], 0)) as [30], sum(isnull([31], 0)) as [31], sum(isnull([32], 0)) as [32], sum(isnull([33], 0)) as [33], sum(isnull([34], 0)) as [34], sum(isnull([35], 0)) as [35], sum(isnull([36], 0)) as [36], sum(isnull([37], 0)) as [37], sum(isnull([38], 0)) as [38], sum(isnull([39], 0)) as [39], sum(isnull([40], 0)) as [40], sum(isnull([41], 0)) as [41], sum(isnull([42], 0)) as [42], sum(isnull([43], 0)) as [43], sum(isnull([44], 0)) as [44], sum(isnull([45], 0)) as [45], sum(isnull([46], 0)) as [46], sum(isnull([47], 0)) as [47], sum(isnull([48], 0)) as [48], sum(isnull([49], 0)) as [49], sum(isnull([50], 0)) as [50], sum(isnull([51], 0)) as [51], sum(isnull([52], 0)) as [52], sum(isnull([1], 0)) as [1], sum(isnull([2], 0)) as [2], sum(isnull([3], 0)) as [3], sum(isnull([4], 0)) as [4], sum(isnull([5], 0)) as [5], sum(isnull([6], 0)) as [6], sum(isnull([7], 0)) as [7], sum(isnull([8], 0)) as [8], sum(isnull([9], 0)) as [9], sum(isnull([10], 0)) as [10], sum(isnull([11], 0)) as [11], sum(isnull([12], 0)) as [12], sum(isnull([13], 0)) as [13], sum(isnull([14], 0)) as [14], sum(isnull([15], 0)) as [15],sum(isnull([16], 0)) as [16], sum(isnull([17], 0)) as [17], sum(isnull([18], 0)) as [18], sum(isnull([19], 0)) as [19], sum(isnull([20], 0)) as [20], sum(isnull([21], 0)) as [21], sum(isnull([22], 0)) as [22], sum(isnull([23], 0)) as [23], sum(isnull([24], 0)) as [24], sum(isnull([25], 0)) as [25], sum(isnull([26], 0)) as [26] " +
                "from SIPGProyeccion P where Temporada = (select Temporada from CatSemanas where getdate() between Inicio and Fin) and Estado = 'A' group by IdAgen)V UNPIVOT(Pronostico FOR Semana in ([27],[28],[29],[30],[31],[32],[33],[34],[35],[36],[37],[38],[39],[40],[41],[42],[43],[44],[45],[46],[47],[48],[49],[50],[51],[52],[1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],[17],[18],[19],[20],[21],[22],[23],[24],[25],[26]))X)V where Pronostico <> 0)PA ON S.Semana = PA.Semana LEFT JOIN(select C.IdAgen, sum(C.Convertidas) as Entregadas, C.Semana FROM(select IdAgen, sum(Convertidas) as Convertidas, Semana FROM SEasonsun1213..UV_ProdRecepcion where CodEstatus<> 'C' and Temporada = (select Temporada from CatSemanas where getdate() between Inicio and Fin) group by IdAgen, Semana Union All select IdAgen, sum(Convertidas) as Convertidas, Semana FROM SeasonPlan..UV_ProdRecepcion where CodEstatus<> 'C' and Temporada = (select Temporada from CatSemanas where getdate() between Inicio and Fin) group by IdAgen, Semana)C GROUP BY C.IdAgen, C.Semana)EA ON PA.IdAgen = EA.IdAgen and S.Semana = EA.Semana " +
                "WHERE getdate() between S.Inicio and S.Fin)SA ON P.IdAgen = SA.IdAgen " +
                "WHERE P.Zona in (" + x + ") group by P.Zona, Z.Descripcion)V order by V.IdRegion").ToList();
            return Json(curvazona, JsonRequestBehavior.AllowGet);
        }

        public JsonResult CurvaListIng(int idregion = 0)
        {
            List<ClassCurva> curvaing = bd.Database.SqlQuery<ClassCurva>("select V.IdRegion,V.IdAgen, V.Ingeniero, V.Pronostico, V.PronosticoAA, V.Entregado, (CASE WHEN V.Entregado ='0' AND  V.PronosticoAA ='0' THEN '' ELSE V.DiferenciaAA END) AS DiferenciaAA, V.PronosticoSA, V.EntregadoSA, (CASE WHEN V.EntregadoSA ='0' AND  V.PronosticoSA ='0' THEN '' ELSE V.DiferenciaSA END) AS DiferenciaSA from(SELECT V.IdRegion,V.IdAgen, V.Ingeniero, (case when round(V.Pronostico,0) is null then '0' else round(V.Pronostico,0) end) as Pronostico, (case when round(V.PronosticoAA,0) is null then '0' else round(V.PronosticoAA,0) end) as PronosticoAA, (case when round(V.Entregado,0) is null then '0' else round(V.Entregado,0) end)as Entregado, (case when round(((V.DiferenciaAA*100)-100),0) is null then '100' else round(((V.DiferenciaAA*100)-100),0) end)as DiferenciaAA, (case when round(V.PronosticoSA,0) is null then '0' else round(V.PronosticoSA,0) end) as PronosticoSA, (case when round(V.EntregadoSA,0) is null then '0' else round(V.EntregadoSA,0) end) as EntregadoSA, (case when round(((V.DiferenciaSA*100)-100),0) is null then '100' else round(((V.DiferenciaSA*100)-100),0) end) as DiferenciaSA FROM(select * FROM(select *, (V.Entregado/ V.PronosticoAA) AS DiferenciaAA, (V.EntregadoSA/V.PronosticoSA) AS DiferenciaSA FROM(select A.IdAgen, A.Nombre AS Ingeniero, A.IdRegion, V.Pronostico, V.Entregado, AA.PronosticoAA, SA.PronosticoSA, SA.EntregadoSA FROM ProdCamposCat C left join ProdProductoresCat P on C.Cod_Prod=P.Cod_Prod left join ProdAgenteCat A on C.IdAgen=A.IdAgen left join(select V.IdAgen, E.Entregado, sum(V.Pronostico) as Pronostico FROM(select * from(select * from(select IdAgen, SUM(isnull([27], 0)) as [27], sum(isnull([28], 0)) as [28], sum(isnull([29], 0)) as[29], sum(isnull([30], 0)) as [30], sum(isnull([31], 0)) as [31], sum(isnull([32], 0)) as [32], sum(isnull([33], 0)) as [33], sum(isnull([34], 0)) as [34], sum(isnull([35], 0)) as [35], sum(isnull([36], 0)) as [36], sum(isnull([37], 0)) as [37], sum(isnull([38], 0)) as [38], sum(isnull([39], 0)) as [39], sum(isnull([40], 0)) as [40], sum(isnull([41], 0)) as [41], sum(isnull([42], 0)) as [42], sum(isnull([43], 0)) as [43], sum(isnull([44], 0)) as [44], sum(isnull([45], 0)) as [45], sum(isnull([46], 0)) as [46], sum(isnull([47], 0)) as [47], sum(isnull([48], 0)) as [48], sum(isnull([49], 0)) as [49], sum(isnull([50], 0)) as [50], sum(isnull([51], 0)) as [51], sum(isnull([52], 0)) as [52], sum(isnull([1], 0)) as [1], sum(isnull([2], 0)) as [2], sum(isnull([3], 0)) as [3], sum(isnull([4], 0)) as [4], sum(isnull([5], 0)) as [5], sum(isnull([6], 0)) as [6], sum(isnull([7], 0)) as [7], sum(isnull([8], 0)) as [8], sum(isnull([9], 0)) as [9], sum(isnull([10], 0)) as [10], sum(isnull([11], 0)) as [11], sum(isnull([12], 0)) as [12], sum(isnull([13], 0)) as [13], sum(isnull([14], 0)) as [14], sum(isnull([15], 0)) as [15],sum(isnull([16], 0)) as [16], sum(isnull([17], 0)) as [17], sum(isnull([18], 0)) as [18], sum(isnull([19], 0)) as [19], sum(isnull([20], 0)) as [20], sum(isnull([21], 0)) as [21], sum(isnull([22], 0)) as [22], sum(isnull([23], 0)) as [23], sum(isnull([24], 0)) as [24], sum(isnull([25], 0)) as [25], sum(isnull([26], 0)) as [26] from SIPGProyeccion P where Temporada = (select Temporada from CatSemanas where getdate() between Inicio and Fin) and Estado = 'A' group by IdAgen)V UNPIVOT(Pronostico FOR Semana in ([27],[28],[29],[30],[31],[32],[33],[34],[35],[36],[37],[38],[39],[40],[41],[42],[43],[44],[45],[46],[47],[48],[49],[50],[51],[52],[1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],[17],[18],[19],[20],[21],[22],[23],[24],[25],[26]))X )V where Pronostico <> 0)V left join(select C.IdAgen, sum(C.Convertidas) as Entregado FROM(select IdAgen, sum(Convertidas) as Convertidas FROM SEasonsun1213..UV_ProdRecepcion where CodEstatus <> 'C' and Temporada = (select Temporada from CatSemanas where getdate() between Inicio and Fin) group by IdAgen Union All SELECT IdAgen, sum(Convertidas) as Convertidas FROM SeasonPlan..UV_ProdRecepcion where CodEstatus <> 'C' and Temporada = (select Temporada from CatSemanas where getdate() between Inicio and Fin) group by IdAgen)C GROUP BY C.IdAgen)E on V.IdAgen = E.IdAgen group by V.IdAgen, E.Entregado)V ON C.IdAgen = V.IdAgen left join(select V.IdAgen, sum(V.Pronostico) as PronosticoAA FROM CatSemanas S left join(select * from(select * from(select IdAgen, SUM(isnull([27], 0)) as [27], sum(isnull([28], 0)) as [28], sum(isnull([29], 0)) as [29], sum(isnull([30], 0)) as [30], sum(isnull([31], 0)) as [31], sum(isnull([32], 0)) as [32], sum(isnull([33], 0)) as [33], sum(isnull([34], 0)) as [34], sum(isnull([35], 0)) as [35], sum(isnull([36], 0)) as [36], sum(isnull([37], 0)) as [37], sum(isnull([38], 0)) as [38], sum(isnull([39], 0)) as [39], sum(isnull([40], 0)) as [40], sum(isnull([41], 0)) as [41], sum(isnull([42], 0)) as [42], sum(isnull([43], 0)) as [43], sum(isnull([44], 0)) as [44], sum(isnull([45], 0)) as [45], sum(isnull([46], 0)) as [46], sum(isnull([47], 0)) as [47], sum(isnull([48], 0)) as [48], sum(isnull([49], 0)) as [49], sum(isnull([50], 0)) as [50], sum(isnull([51], 0)) as [51], sum(isnull([52], 0)) as [52], sum(isnull([1], 0)) as [1], sum(isnull([2], 0)) as [2], sum(isnull([3], 0)) as [3], sum(isnull([4], 0)) as [4], sum(isnull([5], 0)) as [5], sum(isnull([6], 0)) as [6], sum(isnull([7], 0)) as [7], sum(isnull([8], 0)) as [8], sum(isnull([9], 0)) as [9], sum(isnull([10], 0)) as [10], sum(isnull([11], 0)) as [11], sum(isnull([12], 0)) as [12], sum(isnull([13], 0)) as [13], sum(isnull([14], 0)) as [14], sum(isnull([15], 0)) as [15],sum(isnull([16], 0)) as [16], sum(isnull([17], 0)) as [17], sum(isnull([18], 0)) as [18], sum(isnull([19], 0)) as [19], sum(isnull([20], 0)) as [20], sum(isnull([21], 0)) as [21], sum(isnull([22], 0)) as [22], sum(isnull([23], 0)) as [23], sum(isnull([24], 0)) as [24], sum(isnull([25], 0)) as [25], sum(isnull([26], 0)) as [26] from  SIPGProyeccion P where Temporada = (select Temporada from CatSemanas where getdate() between Inicio and Fin) and Estado = 'A' group by IdAgen)V UNPIVOT(Pronostico FOR Semana in ([27],[28],[29],[30],[31],[32],[33],[34],[35],[36],[37],[38],[39],[40],[41],[42],[43],[44],[45],[46],[47],[48],[49],[50],[51],[52],[1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],[17],[18],[19],[20],[21],[22],[23],[24],[25],[26]))X)V where Pronostico <> 0)V on S.Semana = V.Semana " +
                "where S.Inicio between '20190701' and getdate() group by V.IdAgen)AA ON V.IdAgen = AA.IdAgen LEFT JOIN(SELECT PA.IdAgen, round(PA.Pronostico,0) AS PronosticoSA, EA.Entregado AS EntregadoSA, S.Semana FROM CatSemanas S left join(select* from(select* from(select IdAgen, SUM(isnull([27], 0)) as [27], sum(isnull([28], 0)) as [28], sum(isnull([29], 0)) as[29], sum(isnull([30], 0)) as [30], sum(isnull([31], 0)) as [31], sum(isnull([32], 0)) as [32], sum(isnull([33], 0)) as [33], sum(isnull([34], 0)) as [34], sum(isnull([35], 0)) as [35], sum(isnull([36], 0)) as [36], sum(isnull([37], 0)) as [37], sum(isnull([38], 0)) as [38], sum(isnull([39], 0)) as [39], sum(isnull([40], 0)) as [40], sum(isnull([41], 0)) as [41], sum(isnull([42], 0)) as [42], sum(isnull([43], 0)) as [43], sum(isnull([44], 0)) as [44], sum(isnull([45], 0)) as [45], sum(isnull([46], 0)) as [46], sum(isnull([47], 0)) as [47], sum(isnull([48], 0)) as [48], sum(isnull([49], 0)) as [49], sum(isnull([50], 0)) as [50], sum(isnull([51], 0)) as [51], sum(isnull([52], 0)) as [52], sum(isnull([1], 0)) as [1], sum(isnull([2], 0)) as [2], sum(isnull([3], 0)) as [3], sum(isnull([4], 0)) as [4], sum(isnull([5], 0)) as [5], sum(isnull([6], 0)) as [6], sum(isnull([7], 0)) as [7], sum(isnull([8], 0)) as [8], sum(isnull([9], 0)) as [9], sum(isnull([10], 0)) as [10], sum(isnull([11], 0)) as [11], sum(isnull([12], 0)) as [12], sum(isnull([13], 0)) as [13], sum(isnull([14], 0)) as [14], sum(isnull([15], 0)) as [15],sum(isnull([16], 0)) as [16], sum(isnull([17], 0)) as [17], sum(isnull([18], 0)) as [18], sum(isnull([19], 0)) as [19], sum(isnull([20], 0)) as [20], sum(isnull([21], 0)) as [21], sum(isnull([22], 0)) as [22], sum(isnull([23], 0)) as [23], sum(isnull([24], 0)) as [24], sum(isnull([25], 0)) as [25], sum(isnull([26], 0)) as [26] from SIPGProyeccion P where Temporada = (select Temporada from CatSemanas where getdate() between Inicio and Fin) and Estado = 'A' group by IdAgen)V UNPIVOT(Pronostico FOR Semana in ([27],[28],[29],[30],[31],[32],[33],[34],[35],[36],[37],[38],[39],[40],[41],[42],[43],[44],[45],[46],[47],[48],[49],[50],[51],[52],[1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],[17],[18],[19],[20],[21],[22],[23],[24],[25],[26]))X)V where Pronostico <> 0 )PA ON S.Semana = PA.Semana Left Join(select C.IdAgen, sum(C.Convertidas) as Entregado, C.Semana FROM(select IdAgen, sum(Convertidas) as Convertidas, Semana FROM SEasonsun1213..UV_ProdRecepcion where CodEstatus<> 'C' and Temporada = (select Temporada from CatSemanas where getdate() between Inicio and Fin) group by IdAgen, Semana Union All SELECT IdAgen, sum(Convertidas) as Convertidas, Semana FROM SeasonPlan..UV_ProdRecepcion where CodEstatus<> 'C' and Temporada = (select Temporada from CatSemanas where getdate() between Inicio and Fin) group by IdAgen, Semana)C GROUP BY C.IdAgen, C.Semana)EA ON PA.IdAgen = EA.IdAgen and S.Semana = EA.Semana WHERE getdate() between S.Inicio and S.Fin)SA on V.IdAgen = SA.IdAgen " +
                "where V.Pronostico <> 0 GROUP BY A.IdAgen, A.Nombre, A.IdRegion, V.Pronostico, V.Entregado, SA.PronosticoSA, SA.EntregadoSA, AA.PronosticoAA)V)V GROUP BY V.IdAgen, V.Ingeniero, V.IdRegion, V.Pronostico, V.Entregado, V.PronosticoAA, V.DiferenciaAA, V.PronosticoSA, V.EntregadoSA, DiferenciaSA)V)V " +
                "where V.IdRegion in (" + idregion + ") ORDER BY V.Ingeniero").ToList();
            return Json(curvaing, JsonRequestBehavior.AllowGet);
        }

        public JsonResult CurvaListCampos(string Cod_Prod = "")
        {
            List<ClassCurva> curvaxcod = bd.Database.SqlQuery<ClassCurva>("select V.Id as Id_Proyeccion,V.Cod_Prod, V.Cod_Campo, V.Campo, V.Sector,V.Ha,V.Tipo,V.Producto,V.PronosticoT, V.PronosticoA,(case when (LAG(V.EntregadoT) OVER (ORDER BY V.EntregadoT))=V.EntregadoT then 0 else V.EntregadoT end) AS  EntregadoT, (CASE WHEN V.EntregadoT ='0' AND  V.PronosticoA ='0' THEN '' ELSE V.DiferenciaA END) AS DiferenciaA, V.Semana, V.PronosticoSA, (case when (LAG(V.EntregadoSA) OVER (ORDER BY V.EntregadoSA))=V.EntregadoSA then 0 else V.EntregadoSA end) AS EntregadoSA, (CASE WHEN V.EntregadoSA ='0' AND  V.PronosticoSA ='0' THEN '' ELSE V.DiferenciaSA END) AS DiferenciaSA from(SELECT V.Id, V.Cod_Prod, V.Cod_Campo, V.Campo, V.Sector,V.Ha,V.Tipo,V.Producto,(case when round(V.PronosticoT,0) is null then '0' else round(V.PronosticoT,0) end) as PronosticoT, (case when round(V.PronosticoA,0) is null then '0' else round(V.PronosticoA,0) end) as PronosticoA, (case when round(V.EntregadoT,0) is null then '0' else round(V.EntregadoT,0) end)as EntregadoT, (case when round(((V.DiferenciaA*100)-100),0) is null then '100' else round(((V.DiferenciaA*100)-100),0) end)as DiferenciaA, V.Semana, (case when round(V.PronosticoSA,0) is null then '0' else round(V.PronosticoSA,0) end) as PronosticoSA, (case when round(V.EntregadoSA,0) is null then '0' else round(V.EntregadoSA,0) end) as EntregadoSA, (case when round(((V.DiferenciaSA*100)-100),0) is null then '100' else round(((V.DiferenciaSA*100)-100),0) end) as DiferenciaSA FROM(select *,(V.EntregadoT/V.PronosticoA) AS DiferenciaA, (V.EntregadoSA / V.PronosticoSA) AS DiferenciaSA FROM(select V.Id, C.Cod_Prod, C.Cod_Campo, C.Descripcion as Campo, V.Sector, V.Ha, T.Descripcion as Tipo, Pr.Descripcion as Producto, V.PronosticoT, V.EntregadoT, AA.PronosticoA, SA.Semana, SA.PronosticoSA, SA.EntregadoSA FROM ProdCamposCat C left join CatTiposProd T on C.Tipo=T.Tipo left join CatProductos Pr on C.Tipo=Pr.Tipo and C.Producto=Pr.Producto left join(select V.Id,V.IdAgen, V.Cod_prod, V.Cod_Campo,V.Sector,V.Ha,sum(ISNULL(V.Pronostico,0)) as PronosticoT , sum(ISNULL(V.EntregadoT,0)) AS EntregadoT from(select V.Id,V.IdAgen, V.Cod_prod, V.Cod_Campo,V.Sector,V.Ha,sum(V.Pronostico) as Pronostico, E.EntregadoT from(select * from(select V.Id, V.IdAgen, V.Cod_Prod,V.Cod_Campo,V.Sector,V.Ha,max(V.Fecha)as fecha, V.[27],V.[28],V.[29],V.[30],V.[31],V.[32],V.[33],V.[34],V.[35],V.[36],V.[37],V.[38],V.[39],V.[40],V.[41],V.[42],V.[43],V.[44],V.[45],V.[46],V.[47],V.[48],V.[49],V.[50],V.[51],V.[52],V.[53],V.[1],V.[2],V.[3],V.[4],V.[5],V.[6],V.[7],V.[8],V.[9],V.[10],V.[11],V.[12],V.[13],V.[14],V.[15],V.[16],V.[17],V.[18],V.[19],V.[20],V.[21],V.[22],V.[23],V.[24],V.[25],V.[26] From(select Id, IdAgen, Cod_Prod, Cod_Campo, Sector, Ha, Fecha, SUM(isnull([27], 0)) as [27], sum(isnull([28], 0)) as [28], sum(isnull([29], 0)) as[29], sum(isnull([30], 0)) as [30], sum(isnull([31], 0)) as [31], sum(isnull([32], 0)) as [32], sum(isnull([33], 0)) as [33], sum(isnull([34], 0)) as [34], sum(isnull([35], 0)) as [35], sum(isnull([36], 0)) as [36], sum(isnull([37], 0)) as [37], sum(isnull([38], 0)) as [38], sum(isnull([39], 0)) as [39], sum(isnull([40], 0)) as [40], sum(isnull([41], 0)) as [41], sum(isnull([42], 0)) as [42], sum(isnull([43], 0)) as [43], sum(isnull([44], 0)) as [44], sum(isnull([45], 0)) as [45], sum(isnull([46], 0)) as [46], sum(isnull([47], 0)) as [47], sum(isnull([48], 0)) as [48], sum(isnull([49], 0)) as [49], sum(isnull([50], 0)) as [50], sum(isnull([51], 0)) as [51], sum(isnull([52], 0)) as [52], sum(isnull([53], 0)) as [53], sum(isnull([1], 0)) as [1], sum(isnull([2], 0)) as [2], sum(isnull([3], 0)) as [3], sum(isnull([4], 0)) as [4], sum(isnull([5], 0)) as [5], sum(isnull([6], 0)) as [6], sum(isnull([7], 0)) as [7], sum(isnull([8], 0)) as [8], sum(isnull([9], 0)) as [9], sum(isnull([10], 0)) as [10], sum(isnull([11], 0)) as [11], sum(isnull([12], 0)) as [12], sum(isnull([13], 0)) as [13], sum(isnull([14], 0)) as [14], sum(isnull([15], 0)) as [15],sum(isnull([16], 0)) as [16], sum(isnull([17], 0)) as [17], sum(isnull([18], 0)) as [18], sum(isnull([19], 0)) as [19], sum(isnull([20], 0)) as [20], sum(isnull([21], 0)) as [21], sum(isnull([22], 0)) as [22], sum(isnull([23], 0)) as [23], sum(isnull([24], 0)) as [24], sum(isnull([25], 0)) as [25], sum(isnull([26], 0)) as [26] from SIPGProyeccion P where Temporada = (select Temporada from CatSemanas where getdate() between Inicio and Fin) " +
                "and Fecha = (select max(Fecha) from SIPGProyeccion WHERE IdAgen = " + (short)Session["IdAgen"] + " ) group by Id, IdAgen, Cod_Prod,Cod_Campo,Sector, Ha,Fecha)V GROUP BY V.Id, V.IdAgen, V.Cod_Prod,V.Cod_Campo,V.Sector,V.Ha,V.[27],V.[28],V.[29],V.[30],V.[31],V.[32],V.[33],V.[34],V.[35],V.[36],V.[37],V.[38],V.[39],V.[40],V.[41],V.[42],V.[43],V.[44],V.[45],V.[46],V.[47],V.[48],V.[49],V.[50],V.[51],V.[52],V.[53],V.[1],V.[2],V.[3],V.[4],V.[5],V.[6],V.[7],V.[8],V.[9],V.[10],V.[11],V.[12],V.[13],V.[14],V.[15],V.[16],V.[17],V.[18],V.[19],V.[20],V.[21],V.[22],V.[23],V.[24],V.[25],V.[26])V UNPIVOT(Pronostico FOR Semana in ([27],[28],[29],[30],[31],[32],[33],[34],[35],[36],[37],[38],[39],[40],[41],[42],[43],[44],[45],[46],[47],[48],[49],[50],[51],[52],[53],[1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],[17],[18],[19],[20],[21],[22],[23],[24],[25],[26]))X Where Pronostico <> 0)V left join(select C.IdAgen, C.Cod_Prod, C.Cod_Campo, sum(C.Convertidas) as EntregadoT FROM(select IdAgen, Cod_Prod, Cod_Campo, sum(Convertidas) as Convertidas FROM UV_ProdRecepcion where CodEstatus <> 'C' and Temporada = (select Temporada from CatSemanas where getdate() between Inicio and Fin) group by IdAgen, Cod_prod, Cod_Campo)C GROUP BY C.IdAgen, C.Cod_Prod, C.Cod_Campo)E ON V.IdAgen = E.IdAgen AND V.Cod_Prod = E.Cod_Prod AND V.Cod_Campo = E.Cod_Campo group by V.Id, V.IdAgen, V.Cod_prod, V.Cod_Campo, V.Sector,V.Ha,E.EntregadoT)V GROUP BY V.Id, V.IdAgen, V.Cod_prod, V.Cod_Campo,V.Sector,V.Ha)V ON C.IdAgen = V.IdAgen and C.Cod_Prod = V.Cod_Prod and C.Cod_Campo = V.Cod_Campo left join(select V.Id, sum(isnull(V.Pronostico,0)) as PronosticoA FROM CatSemanas S left join(select * from(select V.Id, V.[27], V.[28], V.[29], V.[30], V.[31], V.[32], V.[33], V.[34], V.[35], V.[36], V.[37], V.[38], V.[39], V.[40], V.[41], V.[42], V.[43], V.[44], V.[45], V.[46], V.[47], V.[48], V.[49], V.[50], V.[51], V.[52], V.[53], V.[1], V.[2], V.[3], V.[4], V.[5], V.[6], V.[7], V.[8], V.[9], V.[10], V.[11], V.[12], V.[13], V.[14], V.[15], V.[16], V.[17], V.[18], V.[19], V.[20], V.[21], V.[22], V.[23], V.[24], V.[25], V.[26] From(select Id, SUM(isnull([27], 0)) as [27], sum(isnull([28], 0)) as [28], sum(isnull([29], 0)) as[29], sum(isnull([30], 0)) as [30], sum(isnull([31], 0)) as [31], sum(isnull([32], 0)) as [32], sum(isnull([33], 0)) as [33], sum(isnull([34], 0)) as [34], sum(isnull([35], 0)) as [35], sum(isnull([36], 0)) as [36], sum(isnull([37], 0)) as [37], sum(isnull([38], 0)) as [38], sum(isnull([39], 0)) as [39], sum(isnull([40], 0)) as [40], sum(isnull([41], 0)) as [41], sum(isnull([42], 0)) as [42], sum(isnull([43], 0)) as [43], sum(isnull([44], 0)) as [44], sum(isnull([45], 0)) as [45], sum(isnull([46], 0)) as [46], sum(isnull([47], 0)) as [47], sum(isnull([48], 0)) as [48], sum(isnull([49], 0)) as [49], sum(isnull([50], 0)) as [50], sum(isnull([51], 0)) as [51], sum(isnull([52], 0)) as [52],sum(isnull([53], 0)) as [53],sum(isnull([1], 0)) as [1], sum(isnull([2], 0)) as [2], sum(isnull([3], 0)) as [3], sum(isnull([4], 0)) as [4], sum(isnull([5], 0)) as [5], sum(isnull([6], 0)) as [6], sum(isnull([7], 0)) as [7], sum(isnull([8], 0)) as [8], sum(isnull([9], 0)) as [9], sum(isnull([10], 0)) as [10], sum(isnull([11], 0)) as [11], sum(isnull([12], 0)) as [12], sum(isnull([13], 0)) as [13], sum(isnull([14], 0)) as [14], sum(isnull([15], 0)) as [15],sum(isnull([16], 0)) as [16], sum(isnull([17], 0)) as [17], sum(isnull([18], 0)) as [18], sum(isnull([19], 0)) as [19], sum(isnull([20], 0)) as [20], sum(isnull([21], 0)) as [21], sum(isnull([22], 0)) as [22], sum(isnull([23], 0)) as [23], sum(isnull([24], 0)) as [24], sum(isnull([25], 0)) as [25], sum(isnull([26], 0)) as [26] from SIPGProyeccion P where Temporada = (select Temporada from CatSemanas where getdate() between Inicio and Fin) " +
                "and Fecha = (select max(Fecha) from SIPGProyeccion WHERE IdAgen = " + (short)Session["IdAgen"] + " ) group by Id)V GROUP BY V.Id, V.[27],V.[28],V.[29],V.[30],V.[31],V.[32],V.[33],V.[34],V.[35],V.[36],V.[37],V.[38],V.[39],V.[40],V.[41],V.[42],V.[43],V.[44],V.[45],V.[46],V.[47],V.[48],V.[49],V.[50],V.[51],V.[52],V.[53],V.[1],V.[2],V.[3],V.[4],V.[5],V.[6],V.[7],V.[8],V.[9],V.[10],V.[11],V.[12],V.[13],V.[14],V.[15],V.[16],V.[17],V.[18],V.[19],V.[20],V.[21],V.[22],V.[23],V.[24],V.[25],V.[26])V UNPIVOT(Pronostico FOR Semana in ([27],[28],[29],[30],[31],[32],[33],[34],[35],[36],[37],[38],[39],[40],[41],[42],[43],[44],[45],[46],[47],[48],[49],[50],[51],[52],[53],[1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],[17],[18],[19],[20],[21],[22],[23],[24],[25],[26]))X where Pronostico <> 0)V on S.Semana = V.Semana where S.Inicio between(select Inicio from catsemanas where temporada = (select Temporada from CatSemanas where getdate() between Inicio and Fin) and semana = 27) and getdate() group by V.Id)AA ON V.Id = AA.Id LEFT JOIN(SELECT PA.Id, SUM(isnull(PA.Pronostico,0)) AS PronosticoSA, sum(isnull(EA.Entregado, 0)) AS EntregadoSA, S.Semana FROM CatSemanas S left join(select* from(select V.Id, V.IdAgen, V.Cod_Prod, V.Cod_Campo, V.Sector, max(V.Fecha)as fecha, V.[27],V.[28],V.[29],V.[30],V.[31],V.[32],V.[33],V.[34],V.[35],V.[36],V.[37],V.[38],V.[39],V.[40],V.[41],V.[42],V.[43],V.[44],V.[45],V.[46],V.[47],V.[48],V.[49],V.[50],V.[51],V.[52],V.[53],V.[1],V.[2],V.[3],V.[4],V.[5],V.[6],V.[7],V.[8],V.[9],V.[10],V.[11],V.[12],V.[13],V.[14],V.[15],V.[16],V.[17],V.[18],V.[19],V.[20],V.[21],V.[22],V.[23],V.[24],V.[25],V.[26] From(select Id, IdAgen, Cod_Prod, Cod_Campo, Sector, Fecha, SUM(isnull([27], 0)) as [27], sum(isnull([28], 0)) as [28], sum(isnull([29], 0)) as[29], sum(isnull([30], 0)) as [30], sum(isnull([31], 0)) as [31], sum(isnull([32], 0)) as [32], sum(isnull([33], 0)) as [33], sum(isnull([34], 0)) as [34], sum(isnull([35], 0)) as [35], sum(isnull([36], 0)) as [36], sum(isnull([37], 0)) as [37], sum(isnull([38], 0)) as [38], sum(isnull([39], 0)) as [39], sum(isnull([40], 0)) as [40], sum(isnull([41], 0)) as [41], sum(isnull([42], 0)) as [42], sum(isnull([43], 0)) as [43], sum(isnull([44], 0)) as [44], sum(isnull([45], 0)) as [45], sum(isnull([46], 0)) as [46], sum(isnull([47], 0)) as [47], sum(isnull([48], 0)) as [48], sum(isnull([49], 0)) as [49], sum(isnull([50], 0)) as [50], sum(isnull([51], 0)) as [51], sum(isnull([52], 0)) as [52], sum(isnull([53], 0)) as [53], sum(isnull([1], 0)) as [1], sum(isnull([2], 0)) as [2], sum(isnull([3], 0)) as [3], sum(isnull([4], 0)) as [4], sum(isnull([5], 0)) as [5], sum(isnull([6], 0)) as [6], sum(isnull([7], 0)) as [7], sum(isnull([8], 0)) as [8], sum(isnull([9], 0)) as [9], sum(isnull([10], 0)) as [10], sum(isnull([11], 0)) as [11], sum(isnull([12], 0)) as [12], sum(isnull([13], 0)) as [13], sum(isnull([14], 0)) as [14], sum(isnull([15], 0)) as [15],sum(isnull([16], 0)) as [16], sum(isnull([17], 0)) as [17], sum(isnull([18], 0)) as [18], sum(isnull([19], 0)) as [19], sum(isnull([20], 0)) as [20], sum(isnull([21], 0)) as [21], sum(isnull([22], 0)) as [22], sum(isnull([23], 0)) as [23], sum(isnull([24], 0)) as [24], sum(isnull([25], 0)) as [25], sum(isnull([26], 0)) as [26] from SIPGProyeccion P where Temporada = (select Temporada from CatSemanas where getdate() between Inicio and Fin) " +
                "and Fecha = (select max(Fecha) from SIPGProyeccion WHERE IdAgen = " + (short)Session["IdAgen"] + " ) group by Id,IdAgen, Cod_Prod,Cod_Campo,Sector,Fecha)V GROUP BY V.Id,V.IdAgen, V.Cod_Prod, V.Cod_Campo, V.Sector,V.[27],V.[28],V.[29],V.[30],V.[31],V.[32],V.[33],V.[34],V.[35],V.[36],V.[37],V.[38],V.[39],V.[40],V.[41],V.[42],V.[43],V.[44],V.[45],V.[46],V.[47],V.[48],V.[49],V.[50],V.[51],V.[52],V.[53],V.[1],V.[2],V.[3],V.[4],V.[5],V.[6],V.[7],V.[8],V.[9],V.[10],V.[11],V.[12],V.[13],V.[14],V.[15],V.[16],V.[17],V.[18],V.[19],V.[20],V.[21],V.[22],V.[23],V.[24],V.[25],V.[26])V UNPIVOT(Pronostico FOR Semana in ([27],[28],[29],[30],[31],[32],[33],[34],[35],[36],[37],[38],[39],[40],[41],[42],[43],[44],[45],[46],[47],[48],[49],[50],[51],[52],[53],[1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],[17],[18],[19],[20],[21],[22],[23],[24],[25],[26]))X where Pronostico <> 0)PA ON S.Semana = PA.Semana Left Join(select C.IdAgen, C.Cod_Prod, C.Cod_Campo, sum(C.Convertidas) as Entregado, C.Semana FROM(select IdAgen, Cod_Prod, Cod_Campo, sum(Convertidas) as Convertidas, Semana FROM UV_ProdRecepcion where CodEstatus <> 'C' and Temporada = (select Temporada from CatSemanas where getdate() between Inicio and Fin) group by IdAgen, Cod_prod, Cod_Campo, Semana)C GROUP BY C.IdAgen, C.Cod_Prod, C.Cod_Campo, C.Semana)EA ON PA.IdAgen = EA.IdAgen and PA.Cod_Prod = EA.Cod_Prod and PA.Cod_Campo = EA.Cod_Campo and S.Semana = EA.Semana WHERE getdate() between S.Inicio and S.Fin GROUP BY PA.Id,PA.IdAgen, PA.Cod_Prod, PA.Cod_Campo, PA.Sector,S.Semana)SA on V.Id = SA.Id " +
                "WHERE V.IdAgen = " + (short)Session["IdAgen"] + " and V.Cod_Prod = '" + Cod_Prod + "' group by V.Id, C.Cod_Prod, C.Cod_Campo, C.Descripcion, V.Sector, V.Ha, T.Descripcion, Pr.Descripcion, V.PronosticoT, V.EntregadoT, AA.PronosticoA, SA.Semana, SA.PronosticoSA, SA.EntregadoSA)V)V)V ORDER BY V.Id, V.Cod_Prod, V.Cod_Campo, V.Sector").ToList();
            return Json(curvaxcod, JsonRequestBehavior.AllowGet);
        }

        public ActionResult Semanas(SIPGProyeccion model, int Id_Proyeccion = 0)
        {
            if (Session["Nombre"] != null)
            {
                ViewData["Nombre"] = Session["Nombre"].ToString();
                var item = bd.SIPGProyeccion.Where(x => x.Id == Id_Proyeccion).First();
                return View(item);
            }
            else
            {
                return RedirectToAction("Index", "Login");
            }
        }

        [HttpPost]
        public ActionResult Semanas(SIPGProyeccion model, int Id_Proyeccion = 0, int cantidad = 0, string op = "", string Fecha_corte1 = "", string Fecha_defoliacion = "", string Fecha_corte2 = "", string Fecha_redefoliacion = "")
        {
            if (Session["Nombre"] != null)
            {
                ViewData["Nombre"] = Session["Nombre"].ToString();
                if (Id_Proyeccion != 0)
                {
                    try
                    {
                        double total = 0;
                        string query_up = "";
                        double cant = 0;

                        var semana_actual = bd.CatSemanas.Where(x => DateTime.Now >= x.Inicio && DateTime.Now <= x.Fin).Max(x => x.Semana);
                        var fecha_actual = bd.CatSemanas.Where(x => DateTime.Now >= x.Inicio && DateTime.Now <= x.Fin).First();

                        var query_p = bd.SIPGProyeccion.Where(x => x.Id == Id_Proyeccion).SelectMany(e => new[]
                       {
                        new { Semana = 26, Cantidad = e.C26 },
                        new { Semana = 27, Cantidad = e.C27 },
                        new { Semana = 28, Cantidad = e.C28 },
                        new { Semana = 29, Cantidad = e.C29 },
                        new { Semana = 30, Cantidad = e.C30 },
                        new { Semana = 31, Cantidad = e.C31 },
                        new { Semana = 32, Cantidad = e.C32 },
                        new { Semana = 33, Cantidad = e.C33 },
                        new { Semana = 34, Cantidad = e.C34 },
                        new { Semana = 35, Cantidad = e.C35 },
                        new { Semana = 36, Cantidad = e.C36 },
                        new { Semana = 37, Cantidad = e.C37 },
                        new { Semana = 38, Cantidad = e.C38 },
                        new { Semana = 39, Cantidad = e.C39 },
                        new { Semana = 40, Cantidad = e.C40 },
                        new { Semana = 41, Cantidad = e.C41 },
                        new { Semana = 42, Cantidad = e.C42 },
                        new { Semana = 43, Cantidad = e.C43 },
                        new { Semana = 44, Cantidad = e.C44 },
                        new { Semana = 45, Cantidad = e.C45 },
                        new { Semana = 46, Cantidad = e.C46 },
                        new { Semana = 47, Cantidad = e.C47 },
                        new { Semana = 48, Cantidad = e.C48 },
                        new { Semana = 49, Cantidad = e.C49 },
                        new { Semana = 50, Cantidad = e.C50 },
                        new { Semana = 51, Cantidad = e.C51 },
                        new { Semana = 52, Cantidad = e.C52 },
                        new { Semana = 1, Cantidad = e.C1 },
                        new { Semana = 2, Cantidad = e.C2 },
                        new { Semana = 3, Cantidad = e.C3 },
                        new { Semana = 4, Cantidad = e.C4 },
                        new { Semana = 5, Cantidad = e.C5 },
                        new { Semana = 6, Cantidad = e.C6 },
                        new { Semana = 7, Cantidad = e.C7 },
                        new { Semana = 8, Cantidad = e.C8 },
                        new { Semana = 9, Cantidad = e.C9 },
                        new { Semana = 10, Cantidad = e.C10 },
                        new { Semana = 11, Cantidad = e.C11 },
                        new { Semana = 12, Cantidad = e.C12 },
                        new { Semana = 13, Cantidad = e.C13 },
                        new { Semana = 14, Cantidad = e.C14 },
                        new { Semana = 15, Cantidad = e.C15 },
                        new { Semana = 16, Cantidad = e.C16 },
                        new { Semana = 17, Cantidad = e.C17 },
                        new { Semana = 18, Cantidad = e.C18 },
                        new { Semana = 19, Cantidad = e.C19 },
                        new { Semana = 20, Cantidad = e.C20 },
                        new { Semana = 21, Cantidad = e.C21 },
                        new { Semana = 22, Cantidad = e.C22 },
                        new { Semana = 23, Cantidad = e.C23 },
                        new { Semana = 24, Cantidad = e.C24 },
                        new { Semana = 25, Cantidad = e.C25 },
                        new { Semana = 26, Cantidad = e.C26 }
                    }).Where(x => x.Cantidad != null && x.Cantidad != 0);
                        var q = query_p.AsQueryable();

                        var query = bd.SIPGProyeccion.Where(x => x.Id == Id_Proyeccion).First();

                        if (q != null)
                        {
                            if (cantidad != 0)
                            {
                                var fecha = (from x in bd.CatSemanas
                                             where x.Inicio >= DateTime.Now && x.Temporada == fecha_actual.Temporada
                                             group x by new
                                             {
                                                 Semana = x.Semana
                                             } into x
                                             select new ClassCurva()
                                             {
                                                 Semana = x.Key.Semana
                                             });

                                fecha.ToList().OrderBy(x => x.Inicio);
                                var fecha_sem = fecha.AsQueryable();

                                foreach (var sem in fecha_sem)
                                {
                                    total = Convert.ToDouble(query_p.Sum(y => y.Cantidad));

                                    foreach (var item in q)
                                    {
                                        if (sem.Semana == item.Semana)
                                        {
                                            double cant_sem = Convert.ToDouble(item.Cantidad);
                                            double res = cant_sem / total;
                                            double valor = cantidad * res;

                                            if (op == "aumentar")
                                            {
                                                cant = (double)item.Cantidad + valor;
                                            }
                                            else if (op == "disminuir")
                                            {
                                                cant = (double)item.Cantidad - valor;
                                            }

                                            con.Open();
                                            query_up = "Update SIPGProyeccion SET [" + sem.Semana + "] = " + cant + " where Id= " + Id_Proyeccion + "";
                                            SqlCommand cmd = new SqlCommand(query_up, con);
                                            cmd.ExecuteNonQuery();
                                            con.Close();
                                        }
                                    }
                                }
                            }
                            else
                            {
                                Actualizar_fechas(Id_Proyeccion, Fecha_corte1, Fecha_corte2, Fecha_defoliacion, Fecha_redefoliacion);
                            }
                        }

                        TempData["sms"] = "Cambios guardados correctamente";
                        ViewBag.sms = TempData["sms"].ToString();
                    }

                    catch (Exception e)
                    {
                        e.ToString();
                        TempData["error"] = e.ToString();
                        ViewBag.error = TempData["error"].ToString();
                    }
                }
            }
            else
            {
                return RedirectToAction("Index", "Login");
            }
           
            return View();
        }

        private void Actualizar_fechas(int Id_Proyeccion = 0, string Fecha_corte1 = "", string Fecha_corte2 = "", string Fecha_defoliacion = "", string Fecha_redefoliacion = "")
        {
            List<double> lst_cant = new List<double>();
            List<byte> lst_sem = new List<byte>();
            List<byte> lst_sem2 = new List<byte>();
            int count = 0;

            var fecha_actual = bd.CatSemanas.Where(x => DateTime.Now >= x.Inicio && DateTime.Now <= x.Fin).First();

            var query_p = bd.SIPGProyeccion.Where(x => x.Id == Id_Proyeccion).SelectMany(e => new[]
             {
                        new { Semana = 26, Cantidad = e.C26 },
                        new { Semana = 27, Cantidad = e.C27 },
                        new { Semana = 28, Cantidad = e.C28 },
                        new { Semana = 29, Cantidad = e.C29 },
                        new { Semana = 30, Cantidad = e.C30 },
                        new { Semana = 31, Cantidad = e.C31 },
                        new { Semana = 32, Cantidad = e.C32 },
                        new { Semana = 33, Cantidad = e.C33 },
                        new { Semana = 34, Cantidad = e.C34 },
                        new { Semana = 35, Cantidad = e.C35 },
                        new { Semana = 36, Cantidad = e.C36 },
                        new { Semana = 37, Cantidad = e.C37 },
                        new { Semana = 38, Cantidad = e.C38 },
                        new { Semana = 39, Cantidad = e.C39 },
                        new { Semana = 40, Cantidad = e.C40 },
                        new { Semana = 41, Cantidad = e.C41 },
                        new { Semana = 42, Cantidad = e.C42 },
                        new { Semana = 43, Cantidad = e.C43 },
                        new { Semana = 44, Cantidad = e.C44 },
                        new { Semana = 45, Cantidad = e.C45 },
                        new { Semana = 46, Cantidad = e.C46 },
                        new { Semana = 47, Cantidad = e.C47 },
                        new { Semana = 48, Cantidad = e.C48 },
                        new { Semana = 49, Cantidad = e.C49 },
                        new { Semana = 50, Cantidad = e.C50 },
                        new { Semana = 51, Cantidad = e.C51 },
                        new { Semana = 52, Cantidad = e.C52 },
                        new { Semana = 1, Cantidad = e.C1 },
                        new { Semana = 2, Cantidad = e.C2 },
                        new { Semana = 3, Cantidad = e.C3 },
                        new { Semana = 4, Cantidad = e.C4 },
                        new { Semana = 5, Cantidad = e.C5 },
                        new { Semana = 6, Cantidad = e.C6 },
                        new { Semana = 7, Cantidad = e.C7 },
                        new { Semana = 8, Cantidad = e.C8 },
                        new { Semana = 9, Cantidad = e.C9 },
                        new { Semana = 10, Cantidad = e.C10 },
                        new { Semana = 11, Cantidad = e.C11 },
                        new { Semana = 12, Cantidad = e.C12 },
                        new { Semana = 13, Cantidad = e.C13 },
                        new { Semana = 14, Cantidad = e.C14 },
                        new { Semana = 15, Cantidad = e.C15 },
                        new { Semana = 16, Cantidad = e.C16 },
                        new { Semana = 17, Cantidad = e.C17 },
                        new { Semana = 18, Cantidad = e.C18 },
                        new { Semana = 19, Cantidad = e.C19 },
                        new { Semana = 20, Cantidad = e.C20 },
                        new { Semana = 21, Cantidad = e.C21 },
                        new { Semana = 22, Cantidad = e.C22 },
                        new { Semana = 23, Cantidad = e.C23 },
                        new { Semana = 24, Cantidad = e.C24 },
                        new { Semana = 25, Cantidad = e.C25 },
                        new { Semana = 26, Cantidad = e.C26 }
                    }).Where(x => x.Cantidad != null && x.Cantidad != 0);
            var q = query_p.AsQueryable();
            foreach (var item in q)
            {
                lst_cant.Add((double)item.Cantidad);
            }

            var query = bd.SIPGProyeccion.Where(x => x.Id == Id_Proyeccion).First();

            if (Fecha_corte1 != "")
            {
                DateTime corte1 = Convert.ToDateTime(Fecha_corte1);
                var list_semanas1= from x in bd.CatSemanas                                    
                                    group x by new
                                    {
                                        Inicio = x.Inicio,
                                        Semana = x.Semana
                                    } into x
                                    select new ClassCurva()
                                    {
                                        Inicio = x.Key.Inicio,
                                        Semana = x.Key.Semana
                                    };

                if (Fecha_corte1 != "" && Fecha_corte2 == "")
                {
                    count = lst_cant.Count;
                    list_semanas1 = (from x in bd.CatSemanas
                                     where x.Inicio >= corte1 && x.Temporada == fecha_actual.Temporada
                                     group x by new
                                     {
                                         Inicio = x.Inicio,
                                         Semana = x.Semana
                                     } into x
                                     select new ClassCurva()
                                     {
                                         Inicio = x.Key.Inicio,
                                         Semana = x.Key.Semana
                                     }).OrderBy(x => x.Inicio).Take(count);

                    list_semanas1.ToList();
                }
                else
                {                  
                    //obtener 9 semanas libres a partir de la nueva fecha
                    list_semanas1 = (from x in bd.CatSemanas
                                         where x.Inicio >= corte1 && x.Temporada == fecha_actual.Temporada
                                         group x by new
                                         {
                                             Inicio = x.Inicio,
                                             Semana = x.Semana
                                         } into x
                                         select new ClassCurva()
                                         {
                                             Inicio = x.Key.Inicio,
                                             Semana = x.Key.Semana
                                         }).OrderBy(x => x.Inicio).Take(9);

                    list_semanas1.ToList();
                    count = 9;
                }
                var lista_semanas1 = list_semanas1.AsQueryable();

                foreach (var item in lista_semanas1)
                {
                    lst_sem.Add((byte)item.Semana);
                }

                var sem1 = bd.CatSemanas.Where(x => corte1 >= x.Inicio && corte1 <= x.Fin).FirstOrDefault();
                query.Fecha_corte1 = corte1;
                query.Sem1 = sem1.Semana;
                if (Fecha_defoliacion != "")
                {
                    DateTime defoliacion = Convert.ToDateTime(Fecha_defoliacion);
                    query.Fecha_defoliacion = defoliacion;
                }
                bd.SaveChanges();
            }

            if (Fecha_corte2 != "")
            {
                DateTime corte2 = Convert.ToDateTime(Fecha_corte2);

                //obtener 8 semanas libres a partir de la nueva fecha
                var list_semanas2 = (from x in bd.CatSemanas
                                     where x.Inicio >= corte2 && x.Temporada == fecha_actual.Temporada
                                     group x by new
                                     {
                                         Inicio = x.Inicio,
                                         Semana = x.Semana
                                     } into x
                                     select new ClassCurva()
                                     {
                                         Inicio = x.Key.Inicio,
                                         Semana = x.Key.Semana
                                     }).OrderBy(x => x.Inicio).Take(8);

                list_semanas2.ToList();
                var lista_semanas2 = list_semanas2.AsQueryable();

                foreach (var item in lista_semanas2)
                {
                    lst_sem2.Add((byte)item.Semana);
                }

                var sem2 = bd.CatSemanas.Where(x => corte2 >= x.Inicio && corte2 <= x.Fin).FirstOrDefault();
                query.Fecha_corte2 = corte2;
                query.Sem2 = sem2.Semana;
                if (Fecha_redefoliacion != "")
                {
                    DateTime redefoliacion = Convert.ToDateTime(Fecha_redefoliacion);
                    query.Fecha_redefoliacion = redefoliacion;
                }
                bd.SaveChanges();
            }

            query.C27 = null;
            query.C28 = null;
            query.C29 = null;
            query.C30 = null;
            query.C31 = null;
            query.C32 = null;
            query.C33 = null;
            query.C34 = null;
            query.C35 = null;
            query.C36 = null;
            query.C37 = null;
            query.C38 = null;
            query.C39 = null;
            query.C40 = null;
            query.C41 = null;
            query.C42 = null;
            query.C43 = null;
            query.C44 = null;
            query.C45 = null;
            query.C46 = null;
            query.C47 = null;
            query.C48 = null;
            query.C49 = null;
            query.C50 = null;
            query.C51 = null;
            query.C52 = null;
            query.C53 = null;
            query.C1 = null;
            query.C2 = null;
            query.C3 = null;
            query.C4 = null;
            query.C5 = null;
            query.C6 = null;
            query.C7 = null;
            query.C8 = null;
            query.C9 = null;
            query.C10 = null;
            query.C11 = null;
            query.C12 = null;
            query.C13 = null;
            query.C14 = null;
            query.C15 = null;
            query.C16 = null;
            query.C17 = null;
            query.C18 = null;
            query.C19 = null;
            query.C20 = null;
            query.C21 = null;
            query.C22 = null;
            query.C23 = null;
            query.C24 = null;
            query.C25 = null;
            query.C26 = null;
            bd.SaveChanges();

            con.Open();
            //actualizar corte 1
            if (lst_sem.Count > 0)
            {
                for (int y = 0; y <= count - 1; y++)
                {
                    string query_up = "Update SIPGProyeccion SET [" + lst_sem[y] + "]  = " + lst_cant[y] + " where Id = " + Id_Proyeccion + " ";
                    SqlCommand cmd = new SqlCommand(query_up, con);
                    cmd.ExecuteNonQuery();
                }
            }
            //actualizar corte 2
            int z = 0;
            if (lst_sem2.Count > 0)
            {
                for (int y = 9; y <= lst_cant.Count - 1; y++)
                {
                    string query_up = "Update SIPGProyeccion SET [" + lst_sem2[z] + "]  = " + lst_cant[y] + " where Id = " + Id_Proyeccion + "";
                    SqlCommand cmd = new SqlCommand(query_up, con);
                    cmd.ExecuteNonQuery();
                    z++;
                }
            }
            con.Close();
        }

        public JsonResult Update_Semanas(int Id_Proyeccion=0, int Semana = 0, int Sem = 0, int Cantidad = 0)
        {
            string message = "";
            try
            {
                if (Sem != 0)
                {
                    con.Open();
                    string query_up = "Update SIPGProyeccion SET [" + Semana + "]  = null where Id = " + Id_Proyeccion + "";
                    SqlCommand cmd = new SqlCommand(query_up, con);
                    cmd.ExecuteNonQuery();

                    if (Semana == 52)
                    {
                        Semana = 0;
                    }

                    var fecha_actual = bd.CatSemanas.Where(x => DateTime.Now >= x.Inicio && DateTime.Now <= x.Fin).First();
                    var res = bd.CatSemanas.Where(x => x.Semana == Semana + 1 && x.Temporada == fecha_actual.Temporada).FirstOrDefault();

                    string update_fecha = "Update SIPGProyeccion SET Fecha_corte" + Sem + " = '" + res.Inicio.ToString("yyyy-MM-dd") + "', Sem" + Sem + " = " + res.Semana + " where Id = " + Id_Proyeccion + "";
                    SqlCommand cm = new SqlCommand(update_fecha, con);
                    cm.ExecuteNonQuery();

                    con.Close();
                    message = "SUCCESS";
                }
                else
                {
                    con.Open();
                    string query = "Update SIPGProyeccion SET [" + Semana + "]  = " + Cantidad + " where Id = " + Id_Proyeccion + "";
                    SqlCommand cmd = new SqlCommand(query, con);
                    cmd.ExecuteNonQuery();
                    con.Close();
                    message = "SUCCESS";
                }
            }
            catch (Exception e)
            {
                message = e.ToString();
            }
            return Json(new { Message = message, JsonRequestBehavior.AllowGet });
        }

        public void AddVisita(string Cod_Prod, int Cod_Campo, int Sector = 1)
        {
            try
            {
                SIPGVisitas.IdAgen = (short)Session["IdAgen"];
                SIPGVisitas.Cod_Prod = Cod_Prod;
                SIPGVisitas.Cod_Campo = (short?)Cod_Campo;
                SIPGVisitas.Sector = (short?)Sector;
                SIPGVisitas.Fecha = DateTime.Now;
                bd.SIPGVisitas.Add(SIPGVisitas);
                bd.SaveChanges();

                TempData["sms"] = "Cambios guardados correctamente";
                ViewBag.sms = TempData["sms"].ToString();               
            }
            catch (Exception e)
            {                
                TempData["error"] = e.ToString();
                ViewBag.sms = TempData["error"].ToString();
            }
       }
       public JsonResult Corte1List(int Id_Proyeccion = 0)
        {
            //List<ClassCurva> CorteList = bd.Database.SqlQuery<ClassCurva>("SELECT distinct S.Semana, round(V.Pronostico,0) as Cantidad, S.Inicio from CatSemanas S left join (select * from (select sum([27]) as [27], sum([28]) as [28], sum([29]) as [29],sum([30]) as [30],sum([31]) as [31],sum([32]) as [32],sum([33]) as [33],sum([34]) as [34], sum([35]) as [35],sum([36]) as [36],sum([37]) as [37],sum([38]) as [38], sum([39]) as [39],sum([40]) as [40],sum([41]) as [41],sum([42]) as [42],sum([43]) as [43],sum([44]) as [44],sum([45]) as [45],sum([46]) as [46],sum([47]) as [47],sum([48]) as [48], sum([49]) as [49],sum([50]) as [50],sum([51]) as [51],sum([52]) as [52],sum([1]) as [1], sum([2]) as [2], sum([3]) as [3], sum([4]) as [4], sum([5]) as [5], sum([6]) as [6], sum([7]) as [7], sum([8]) as [8], sum([9]) as [9], sum([10]) as [10], sum([11]) as [11], sum([12]) as [12], sum([13]) as [13], sum([14]) as [14], sum([15]) as [15], sum([16]) as [16], sum([17]) as [17], sum([18]) as [18], sum([19]) as [19], sum([20]) as [20], sum([21]) as [21], sum([22]) as [22], sum([23]) as [23], sum([24]) as [24], sum([25]) as [25], sum([26]) as [26] FROM SIPGProyeccion WHERE Temporada = (select Temporada from CatSemanas where getdate() between Inicio and Fin) and " +
            //    "cod_prod = '" + Cod_Prod + "' and cod_campo = " + Cod_Campo + " and sector = " + Sector + " and Fecha=(select max(Fecha) from SIPGProyeccion) group by[27],[28],[29],[30],[31],[32],[33],[34],[35],[36],[37],[38],[39],[40],[41],[42],[43],[44],[45],[46],[47],[48],[49],[50],[51],[52],[1], [2], [3], [4], [5], [6], [7], [8], [9], [10], [11], [12], [13], [14], [15], [16], [17], [18], [19], [20], [21], [22], [23], [24], [25], [26])V UNPIVOT(Pronostico FOR Semana in ([27],[28],[29],[30],[31],[32],[33],[34],[35],[36],[37],[38],[39],[40],[41],[42],[43],[44],[45],[46],[47],[48],[49],[50],[51],[52],[1], [2], [3], [4], [5], [6], [7], [8], [9], [10], [11], [12], [13], [14], [15], [16], [17], [18], [19], [20], [21], [22], [23], [24], [25], [26]))X )V on S.Semana = V.Semana WHERE V.Pronostico <> 0 and S.Temporada = (select Temporada from CatSemanas where getdate() between Inicio and Fin) and V.Semana between 27 and 52 order by S.Inicio").ToList();

            List<ClassCurva> CorteList = bd.Database.SqlQuery<ClassCurva>("SELECT distinct S.Semana, round(V.Pronostico,0) as Cantidad, S.Inicio from CatSemanas S left join (select * from (select sum([27]) as [27], sum([28]) as [28], sum([29]) as [29],sum([30]) as [30],sum([31]) as [31],sum([32]) as [32],sum([33]) as [33],sum([34]) as [34], sum([35]) as [35],sum([36]) as [36],sum([37]) as [37],sum([38]) as [38], sum([39]) as [39],sum([40]) as [40],sum([41]) as [41],sum([42]) as [42],sum([43]) as [43],sum([44]) as [44],sum([45]) as [45],sum([46]) as [46],sum([47]) as [47],sum([48]) as [48], sum([49]) as [49],sum([50]) as [50],sum([51]) as [51],sum([52]) as [52],sum([1]) as [1], sum([2]) as [2], sum([3]) as [3], sum([4]) as [4], sum([5]) as [5], sum([6]) as [6], sum([7]) as [7], sum([8]) as [8], sum([9]) as [9], sum([10]) as [10], sum([11]) as [11], sum([12]) as [12], sum([13]) as [13], sum([14]) as [14], sum([15]) as [15], sum([16]) as [16], sum([17]) as [17], sum([18]) as [18], sum([19]) as [19], sum([20]) as [20], sum([21]) as [21], sum([22]) as [22], sum([23]) as [23], sum([24]) as [24], sum([25]) as [25], sum([26]) as [26] FROM SIPGProyeccion " +
                "WHERE Temporada = (select Temporada from CatSemanas where getdate() between Inicio and Fin) and " +
               "Id = " + Id_Proyeccion + " group by[27],[28],[29],[30],[31],[32],[33],[34],[35],[36],[37],[38],[39],[40],[41],[42],[43],[44],[45],[46],[47],[48],[49],[50],[51],[52],[1], [2], [3], [4], [5], [6], [7], [8], [9], [10], [11], [12], [13], [14], [15], [16], [17], [18], [19], [20], [21], [22], [23], [24], [25], [26])V UNPIVOT(Pronostico FOR Semana in ([27],[28],[29],[30],[31],[32],[33],[34],[35],[36],[37],[38],[39],[40],[41],[42],[43],[44],[45],[46],[47],[48],[49],[50],[51],[52],[1], [2], [3], [4], [5], [6], [7], [8], [9], [10], [11], [12], [13], [14], [15], [16], [17], [18], [19], [20], [21], [22], [23], [24], [25], [26]))X )V on S.Semana = V.Semana WHERE V.Pronostico <> 0 and S.Temporada = (select Temporada from CatSemanas where getdate() between Inicio and Fin) and V.Semana between 27 and 52 order by S.Inicio").ToList();
            //List<ClassCurva> CorteList = bd.Database.SqlQuery<ClassCurva>("SELECT top 9 S.Semana, round(V.Pronostico,0) as Cantidad, S.Inicio from CatSemanas S left join (select * from (select sum([27]) as [27], sum([28]) as [28], sum([29]) as [29],sum([30]) as [30],sum([31]) as [31],sum([32]) as [32],sum([33]) as [33],sum([34]) as [34], sum([35]) as [35],sum([36]) as [36],sum([37]) as [37],sum([38]) as [38], sum([39]) as [39],sum([40]) as [40],sum([41]) as [41],sum([42]) as [42],sum([43]) as [43],sum([44]) as [44],sum([45]) as [45],sum([46]) as [46],sum([47]) as [47],sum([48]) as [48], sum([49]) as [49],sum([50]) as [50],sum([51]) as [51],sum([52]) as [52],sum([53]) as [53],sum([1]) as [1], sum([2]) as [2], sum([3]) as [3], sum([4]) as [4], sum([5]) as [5], sum([6]) as [6], sum([7]) as [7], sum([8]) as [8], sum([9]) as [9], sum([10]) as [10], sum([11]) as [11], sum([12]) as [12], sum([13]) as [13], sum([14]) as [14], sum([15]) as [15], sum([16]) as [16], sum([17]) as [17], sum([18]) as [18], sum([19]) as [19], sum([20]) as [20], sum([21]) as [21], sum([22]) as [22], sum([23]) as [23], sum([24]) as [24], sum([25]) as [25], sum([26]) as [26] FROM SIPGProyeccion " +
            //     "WHERE Temporada = (select Temporada from CatSemanas where getdate() between Inicio and Fin) and " +
            //    "Id = " + Id_Proyeccion + " group by[27],[28],[29],[30],[31],[32],[33],[34],[35],[36],[37],[38],[39],[40],[41],[42],[43],[44],[45],[46],[47],[48],[49],[50],[51],[52],[53],[1], [2], [3], [4], [5], [6], [7], [8], [9], [10], [11], [12], [13], [14], [15], [16], [17], [18], [19], [20], [21], [22], [23], [24], [25], [26])V UNPIVOT(Pronostico FOR Semana in ([27],[28],[29],[30],[31],[32],[33],[34],[35],[36],[37],[38],[39],[40],[41],[42],[43],[44],[45],[46],[47],[48],[49],[50],[51],[52],[53],[1], [2], [3], [4], [5], [6], [7], [8], [9], [10], [11], [12], [13], [14], [15], [16], [17], [18], [19], [20], [21], [22], [23], [24], [25], [26]))X )V on S.Semana = V.Semana WHERE V.Pronostico <> 0 and S.Temporada = (select Temporada from CatSemanas where getdate() between Inicio and Fin) order by S.Inicio").ToList();

            return Json(CorteList, JsonRequestBehavior.AllowGet);
        }
        public JsonResult Corte2List(int Id_Proyeccion = 0)
        {
            List<ClassCurva> CorteList = bd.Database.SqlQuery<ClassCurva>("SELECT distinct S.Semana, round(V.Pronostico,0) as Cantidad, S.Inicio from CatSemanas S left join (select * from (select sum([27]) as [27], sum([28]) as [28], sum([29]) as [29],sum([30]) as [30],sum([31]) as [31],sum([32]) as [32],sum([33]) as [33],sum([34]) as [34], sum([35]) as [35],sum([36]) as [36],sum([37]) as [37],sum([38]) as [38], sum([39]) as [39],sum([40]) as [40],sum([41]) as [41],sum([42]) as [42],sum([43]) as [43],sum([44]) as [44],sum([45]) as [45],sum([46]) as [46],sum([47]) as [47],sum([48]) as [48], sum([49]) as [49],sum([50]) as [50],sum([51]) as [51],sum([52]) as [52],sum([1]) as [1], sum([2]) as [2], sum([3]) as [3], sum([4]) as [4], sum([5]) as [5], sum([6]) as [6], sum([7]) as [7], sum([8]) as [8], sum([9]) as [9], sum([10]) as [10], sum([11]) as [11], sum([12]) as [12], sum([13]) as [13], sum([14]) as [14], sum([15]) as [15], sum([16]) as [16], sum([17]) as [17], sum([18]) as [18], sum([19]) as [19], sum([20]) as [20], sum([21]) as [21], sum([22]) as [22], sum([23]) as [23], sum([24]) as [24], sum([25]) as [25], sum([26]) as [26] FROM SIPGProyeccion " +
                "WHERE Temporada = (select Temporada from CatSemanas where getdate() between Inicio and Fin) and " +
               "Id = " + Id_Proyeccion + " group by[27],[28],[29],[30],[31],[32],[33],[34],[35],[36],[37],[38],[39],[40],[41],[42],[43],[44],[45],[46],[47],[48],[49],[50],[51],[52],[1], [2], [3], [4], [5], [6], [7], [8], [9], [10], [11], [12], [13], [14], [15], [16], [17], [18], [19], [20], [21], [22], [23], [24], [25], [26])V UNPIVOT(Pronostico FOR Semana in ([27],[28],[29],[30],[31],[32],[33],[34],[35],[36],[37],[38],[39],[40],[41],[42],[43],[44],[45],[46],[47],[48],[49],[50],[51],[52],[1], [2], [3], [4], [5], [6], [7], [8], [9], [10], [11], [12], [13], [14], [15], [16], [17], [18], [19], [20], [21], [22], [23], [24], [25], [26]))X )V on S.Semana = V.Semana WHERE V.Pronostico <> 0 and S.Temporada = (select Temporada from CatSemanas where getdate() between Inicio and Fin) and V.Semana between 1 and 26 order by S.Inicio").ToList();

            //List<ClassCurva> CorteList = bd.Database.SqlQuery<ClassCurva>("WITH t AS(select top 8 round(V.Pronostico,0) as Cantidad, S.Semana, S.Inicio from(select * from(select sum([27]) as [27], sum([28]) as [28], sum([29]) as [29], sum([30]) as [30], sum([31]) as [31],sum([32]) as [32],sum([33]) as [33],sum([34]) as [34], sum([35]) as [35],sum([36]) as [36],sum([37]) as [37],sum([38]) as [38], sum([39]) as [39],sum([40]) as [40],sum([41]) as [41],sum([42]) as [42],sum([43]) as [43],sum([44]) as [44],sum([45]) as [45],sum([46]) as [46],sum([47]) as [47],sum([48]) as [48], sum([49]) as [49],sum([50]) as [50],sum([51]) as [51],sum([52]) as [52],sum([53]) as [53],sum([1]) as [1], sum([2]) as [2], sum([3]) as [3], sum([4]) as [4], sum([5]) as [5], sum([6]) as [6], sum([7]) as [7], sum([8]) as [8], sum([9]) as [9], sum([10]) as [10], sum([11]) as [11], sum([12]) as [12], sum([13]) as [13], sum([14]) as [14], sum([15]) as [15], sum([16]) as [16], sum([17]) as [17], sum([18]) as [18], sum([19]) as [19], sum([20]) as [20], sum([21]) as [21], sum([22]) as [22], sum([23]) as [23], sum([24]) as [24], sum([25]) as [25], sum([26]) as [26] FROM SIPGProyeccion WHERE Temporada = (select Temporada from CatSemanas where getdate() between Inicio and Fin) and " +
            //"Id = "+ Id_Proyeccion + " group by[27],[28],[29],[30],[31],[32],[33],[34],[35],[36],[37],[38],[39],[40],[41],[42],[43],[44],[45],[46],[47],[48],[49],[50],[51],[52],[53],[1], [2], [3], [4], [5], [6], [7], [8], [9], [10], [11], [12], [13], [14], [15], [16], [17], [18], [19], [20], [21], [22], [23], [24], [25], [26])V UNPIVOT(Pronostico FOR Semana in ([27],[28],[29],[30],[31],[32],[33],[34],[35],[36],[37],[38],[39],[40],[41],[42],[43],[44],[45],[46],[47],[48],[49],[50],[51],[52],[1], [2], [3], [4], [5], [6], [7], [8], [9], [10], [11], [12], [13], [14], [15], [16], [17], [18], [19], [20], [21], [22], [23], [24], [25], [26]))X WHERE Pronostico <> 0)V Left join CatSemanas S on V.Semana = S.Semana WHERE S.Temporada = (select Temporada from CatSemanas where getdate() between Inicio and Fin) order by S.Inicio desc) SELECT * FROM t " +
            //"where Cantidad is not null order by Inicio").ToList();
            return Json(CorteList, JsonRequestBehavior.AllowGet);
        }
        public void VARIACION_ESTIMACION_BERRIES()
        {
            ViewData["Nombre"] = Session["Nombre"].ToString();

            ExcelPackage excel = new ExcelPackage();
            var fecha_actual = bd.CatSemanas.Where(a => DateTime.Now >= a.Inicio && DateTime.Now <= a.Fin).First();

            //ZARZAMORA
            var zarz = bd.Database.SqlQuery<ClassCurva>("Select I.Nombre as Asesor, P.Nombre as Productor, S.Cod_Prod, S.Cod_Campo, C.Descripcion as Campo, L.Descripcion as Localidad, isnull(S.Num_corte,0) as Num_corte,S.Sector,round(S.Ha,2) as Ha,isnull(T.Descripcion,'') AS Tipo, isnull(V.Descripcion, '') AS Producto, isnull(S.Numplantas_xha, 0) as Numplantas_xha, isnull(S.Manejo, '') as Manejo, isnull(S.Tipo_plantacion, '') as Tipo_plantacion, isnull(CONVERT(VARCHAR(10), S.Fecha_plantacion, 23), '') as Fecha_plantacion,isnull(CONVERT(VARCHAR(10), S.Fecha_poda, 23), '') as Fecha_poda, " +
               "isnull(CONVERT(VARCHAR(10), S.Fecha_defoliacion, 23), '') as Fecha_defoliacionR, isnull(CONVERT(VARCHAR(10), S.Fecha_corte1, 23), '') as Fecha_corte1R,isnull(CONVERT(VARCHAR(10), S.Fecha_redefoliacion, 23), '') as Fecha_redefoliacionR, isnull(CONVERT(VARCHAR(10), S.Fecha_corte2, 23), '') as Fecha_corte2R, isnull(S.Sem1, '') as Sem1, isnull(S.Sem2, '') as Sem2,isnull(S.Plantacion, 0) as Plantacion, isnull(S.Caja1, 0) as Caja1, isnull(S.Caja2, 0) as Caja2, isnull(S.Estructura, '') as Estructura, isnull(S.Tipo_certificacion, '') as Tipo_certificacion,isnull(S.Tesco, '') as Tesco,isnull(S.Edad_planta, 0) as Edad_planta,isnull(S.Tipo_plantacion2, '') as Tipo_plantacion2,isnull(CONVERT(VARCHAR(10), S.Fecha_podamediacaña, 23), '') as Fecha_podamediacaña,isnull(S.Temporada, '') as Temporada,isnull(Z.DescZona, '') as Zona,isnull(A.Acopio, '') as Acopio," +
               "isnull(S.[27], 0) as _27, isnull(S.[28], 0) as _28, isnull(S.[29], 0) as _29, isnull(S.[30], 0) as _30, isnull(S.[31], 0) as _31, isnull(S.[32], 0) as _32, isnull(S.[33], 0) _33, isnull(S.[34], 0) as _34, isnull(S.[35], 0) as _35,isnull(S.[36], 0) as _36, isnull(S.[37], 0) as _37, isnull(S.[38], 0) as _38, isnull(S.[39], 0) as _39, isnull(S.[40], 0) as _40, isnull(S.[41], 0) as _41, isnull(S.[42], 0) as _42, isnull(S.[43], 0) as _43, isnull(S.[44], 0) as _44, isnull(S.[45], 0) as _45, isnull(S.[46], 0) as _46,isnull(S.[47], 0) as _47, isnull(S.[48], 0) as _48, isnull(S.[49], 0) as _49, isnull(S.[50], 0) as _50, isnull(S.[51], 0) as _51, isnull(S.[52], 0) as _52, isnull(S.[1], 0) as _1, isnull(S.[2], 0) as _2, isnull(S.[3], 0) as _3, isnull(S.[4], 0) as _4, isnull(S.[5], 0) as _5,isnull(S.[6], 0) as _6, isnull(S.[7], 0) as _7, isnull(S.[8], 0) as _8, isnull(S.[9], 0) as _9, isnull(S.[10], 0) as _10, isnull(S.[11], 0) as _11, isnull(S.[12], 0) as _12, isnull(S.[13], 0) as _13, isnull(S.[14], 0) as _14, isnull(S.[15], 0) as _15, isnull(S.[16], 0) as _16,isnull(S.[17], 0) as _17, isnull(S.[18], 0) as _18, isnull(S.[19], 0) as _19, isnull(S.[20], 0) as _20, isnull(S.[21], 0) as _21, isnull(S.[22], 0) as _22, isnull(S.[23], 0) as _23, isnull(S.[24], 0) as _24, isnull(S.[25], 0) as _25, isnull(S.[26], 0) as _26 " +
               "FROM(select V.IdAgen, V.Cod_Prod, V.Cod_Campo, V.Num_corte, V.Sector, V.Ha, V.Numplantas_xha, V.Manejo, V.Tipo_plantacion, V.Fecha_plantacion, V.Fecha_poda, V.Fecha_defoliacion, V.Fecha_corte1, V.Fecha_redefoliacion, V.Fecha_corte2, V.Sem1,V.Sem2, V.Plantacion, V.Caja1, V.Caja2, V.Estructura, V.Tipo_certificacion, V.Tesco, V.Edad_planta, V.Tipo_plantacion2, V.Fecha_podamediacaña, V.Temporada," +
               "max(V.Fecha) as fecha, V.[27], V.[28], V.[29], V.[30], V.[31], V.[32], V.[33], V.[34], V.[35], V.[36], V.[37], V.[38], V.[39], V.[40], V.[41], V.[42], V.[43], V.[44], V.[45], V.[46], V.[47], V.[48], V.[49], V.[50], V.[51], V.[52], V.[1], V.[2], V.[3], V.[4], V.[5], V.[6], V.[7], V.[8], V.[9], V.[10], V.[11], V.[12], V.[13], V.[14], V.[15], V.[16], V.[17], V.[18], V.[19], V.[20], V.[21], V.[22], V.[23], V.[24], V.[25], V.[26] " +
               "From(select * from SIPGProyeccion P where Temporada = (select Temporada from CatSemanas where getdate() between Inicio and Fin)  and Fecha = (select max(Fecha) from SIPGProyeccion))V GROUP BY V.IdAgen,V.Cod_Prod,V.Cod_Campo,V.Num_corte,V.Sector,V.Ha,V.Numplantas_xha,V.Manejo,V.Tipo_plantacion,V.Fecha_plantacion,V.Fecha_poda,V.Fecha_poda,V.Fecha_defoliacion,V.Fecha_corte1,V.Fecha_redefoliacion,V.Fecha_corte2,V.Sem1, V.Sem2,V.Plantacion,V.Caja1,V.Caja2,V.Estructura,V.Tipo_certificacion,V.Tesco,V.Edad_planta,V.Tipo_plantacion2,V.Fecha_podamediacaña,V.Temporada,V.[27],V.[28],V.[29],V.[30],V.[31],V.[32],V.[33],V.[34],V.[35],V.[36],V.[37],V.[38],V.[39],V.[40],V.[41],V.[42],V.[43],V.[44],V.[45],V.[46],V.[47],V.[48],V.[49],V.[50],V.[51],V.[52],V.[1],V.[2],V.[3],V.[4],V.[5],V.[6],V.[7],V.[8],V.[9],V.[10],V.[11],V.[12],V.[13],V.[14],V.[15],V.[16],V.[17],V.[18],V.[19],V.[20],V.[21],V.[22],V.[23],V.[24],V.[25],V.[26])S " +
               "LEFT JOIN ProdAgenteCat I on S.IdAgen = I.IdAgen LEFT JOIN ProdCamposCat C on S.IdAgen = C.IdAgen and S.Cod_Prod = C.Cod_Prod AND S.Cod_Campo = C.Cod_Campo LEFT JOIN ProdProductoresCat P on S.Cod_Prod = P.Cod_Prod LEFT JOIN ProdZonasRastreoCat Z on C.IdZona = Z.IdZona LEFT JOIN CatAcopios A on C.IdAcopio = A.IdAcopio LEFT JOIN CatTiposProd T on C.Tipo = T.Tipo LEFT JOIN CatProductos V on C.Producto = V.Producto AND C.Tipo = V.Tipo LEFT JOIN CatLocalidades L on C.CodLocalidad = L.CodLocalidad " +
               "WHERE I.IdAgen = " + (short)Session["IdAgen"] + " and C.Tipo = 1 order by S.Cod_Prod, S.Cod_Campo, C.Descripcion, L.Descripcion, S.Num_corte,S.Sector").ToList();

            if (zarz.Count > 0)
            {
                ExcelWorksheet ws = excel.Workbook.Worksheets.Add("ZARZAMORA");
                ws.Cells["A1"].Value = "CURVA DE PRODUCCION " + fecha_actual.Temporada;
                ws.Cells["A2"].Value = "Ingeniero";
                ws.Cells["B2"].Value = "Codigo";
                ws.Cells["C2"].Value = "Productor";
                ws.Cells["D2"].Value = "Campo";
                ws.Cells["E2"].Value = "Ubicacion";
                ws.Cells["F2"].Value = "Num_corte";
                ws.Cells["G2"].Value = "Sector";
                ws.Cells["H2"].Value = "Superficie";
                ws.Cells["I2"].Value = "Cultivo";
                ws.Cells["J2"].Value = "Variedad";
                ws.Cells["K2"].Value = "Num plantas xha";
                ws.Cells["L2"].Value = "Manejo";
                ws.Cells["M2"].Value = "Tipo_plantacion";
                ws.Cells["N2"].Value = "Fecha_plantacion";
                ws.Cells["O2"].Value = "Fecha_poda";
                ws.Cells["P2"].Value = "Fecha_defoliacion";
                ws.Cells["Q2"].Value = "FechaIniciocorte1";
                ws.Cells["R2"].Value = "Fecha_redefoliacion";
                ws.Cells["S2"].Value = "FechaIniciocorte2";
                ws.Cells["T2"].Value = "Sem.Inic -1ra cosecha";
                ws.Cells["U2"].Value = "Sem.Inic - 2da cosecha";
                ws.Cells["V2"].Value = "Plantacion";
                ws.Cells["W2"].Value = "Caja1";
                ws.Cells["X2"].Value = "Caja2";
                ws.Cells["Y2"].Value = "Estructura";
                ws.Cells["Z2"].Value = "Tipo_certificacion";
                ws.Cells["AA2"].Value = "Programa_Tesco";
                ws.Cells["AB2"].Value = "Edad_planta";
                ws.Cells["AC2"].Value = "Tipo_plantacion2";
                ws.Cells["AD2"].Value = "Fecha_podamediacaña";
                ws.Cells["AE2"].Value = "Temporada";
                ws.Cells["AF2"].Value = "Zona";
                ws.Cells["AG2"].Value = "Acopio";
                ws.Cells["AH2"].Value = "27";
                ws.Cells["AI2"].Value = "28";
                ws.Cells["AJ2"].Value = "29";
                ws.Cells["AK2"].Value = "30";
                ws.Cells["AL2"].Value = "31";
                ws.Cells["AM2"].Value = "32";
                ws.Cells["AN2"].Value = "33";
                ws.Cells["AO2"].Value = "34";
                ws.Cells["AP2"].Value = "35";
                ws.Cells["AQ2"].Value = "36";
                ws.Cells["AR2"].Value = "37";
                ws.Cells["AS2"].Value = "38";
                ws.Cells["AT2"].Value = "39";
                ws.Cells["AU2"].Value = "40";
                ws.Cells["AV2"].Value = "41";
                ws.Cells["AW2"].Value = "42";
                ws.Cells["AX2"].Value = "43";
                ws.Cells["AY2"].Value = "44";
                ws.Cells["AZ2"].Value = "45";
                ws.Cells["BA2"].Value = "46";
                ws.Cells["BB2"].Value = "47";
                ws.Cells["BC2"].Value = "48";
                ws.Cells["BD2"].Value = "49";
                ws.Cells["BE2"].Value = "50";
                ws.Cells["BF2"].Value = "51";
                ws.Cells["BG2"].Value = "52";
                ws.Cells["BH2"].Value = "1";
                ws.Cells["BI2"].Value = "2";
                ws.Cells["BJ2"].Value = "3";
                ws.Cells["BK2"].Value = "4";
                ws.Cells["BL2"].Value = "5";
                ws.Cells["BM2"].Value = "6";
                ws.Cells["BN2"].Value = "7";
                ws.Cells["BO2"].Value = "8";
                ws.Cells["BP2"].Value = "9";
                ws.Cells["BQ2"].Value = "10";
                ws.Cells["BR2"].Value = "11";
                ws.Cells["BS2"].Value = "12";
                ws.Cells["BT2"].Value = "13";
                ws.Cells["BU2"].Value = "14";
                ws.Cells["BV2"].Value = "15";
                ws.Cells["BW2"].Value = "16";
                ws.Cells["BX2"].Value = "17";
                ws.Cells["BY2"].Value = "18";
                ws.Cells["BZ2"].Value = "19";
                ws.Cells["CA2"].Value = "20";
                ws.Cells["CB2"].Value = "21";
                ws.Cells["CC2"].Value = "22";
                ws.Cells["CD2"].Value = "23";
                ws.Cells["CE2"].Value = "24";
                ws.Cells["CF2"].Value = "25";
                ws.Cells["CG2"].Value = "26";
                int x = 3;
                foreach (var item in zarz)
                {
                    ws.Cells[string.Format("A{0}", x)].Value = item.Asesor;
                    ws.Cells[string.Format("B{0}", x)].Value = item.Cod_Prod;
                    ws.Cells[string.Format("C{0}", x)].Value = item.Productor;
                    ws.Cells[string.Format("D{0}", x)].Value = item.Cod_Campo;
                    ws.Cells[string.Format("E{0}", x)].Value = item.Localidad;
                    ws.Cells[string.Format("F{0}", x)].Value = item.Num_corte;
                    ws.Cells[string.Format("G{0}", x)].Value = item.Sector;
                    ws.Cells[string.Format("H{0}", x)].Value = item.Ha;
                    ws.Cells[string.Format("I{0}", x)].Value = item.Tipo;
                    ws.Cells[string.Format("J{0}", x)].Value = item.Producto;
                    ws.Cells[string.Format("K{0}", x)].Value = item.Numplantas_xha;
                    ws.Cells[string.Format("L{0}", x)].Value = item.Manejo;
                    ws.Cells[string.Format("M{0}", x)].Value = item.Tipo_plantacion;
                    ws.Cells[string.Format("N{0}", x)].Value = item.Fecha_plantacion;
                    ws.Cells[string.Format("O{0}", x)].Value = item.Fecha_poda;
                    ws.Cells[string.Format("P{0}", x)].Value = item.Fecha_defoliacionR;
                    ws.Cells[string.Format("Q{0}", x)].Value = item.Fecha_corte1R;
                    ws.Cells[string.Format("R{0}", x)].Value = item.Fecha_redefoliacionR;
                    ws.Cells[string.Format("S{0}", x)].Value = item.Fecha_corte2R;
                    ws.Cells[string.Format("T{0}", x)].Value = item.Sem1;
                    ws.Cells[string.Format("U{0}", x)].Value = item.Sem2;
                    ws.Cells[string.Format("V{0}", x)].Value = item.Plantacion;
                    ws.Cells[string.Format("W{0}", x)].Value = item.Caja1;
                    ws.Cells[string.Format("X{0}", x)].Value = item.Caja2;
                    ws.Cells[string.Format("Y{0}", x)].Value = item.Estructura;
                    ws.Cells[string.Format("Z{0}", x)].Value = item.Tipo_certificacion;
                    ws.Cells[string.Format("AA{0}", x)].Value = item.Tesco;
                    ws.Cells[string.Format("AB{0}", x)].Value = item.Edad_planta;
                    ws.Cells[string.Format("AC{0}", x)].Value = item.Tipo_plantacion2;
                    ws.Cells[string.Format("AD{0}", x)].Value = item.Fecha_podamediacaña;
                    ws.Cells[string.Format("AE{0}", x)].Value = item.Temporada;
                    ws.Cells[string.Format("AF{0}", x)].Value = item.Zona;
                    ws.Cells[string.Format("AG{0}", x)].Value = item.Acopio;
                    ws.Cells[string.Format("AH{0}", x)].Value = item._27;
                    ws.Cells[string.Format("AI{0}", x)].Value = item._28;
                    ws.Cells[string.Format("AJ{0}", x)].Value = item._29;
                    ws.Cells[string.Format("AK{0}", x)].Value = item._30;
                    ws.Cells[string.Format("AL{0}", x)].Value = item._31;
                    ws.Cells[string.Format("AM{0}", x)].Value = item._32;
                    ws.Cells[string.Format("AN{0}", x)].Value = item._33;
                    ws.Cells[string.Format("AO{0}", x)].Value = item._34;
                    ws.Cells[string.Format("AP{0}", x)].Value = item._35;
                    ws.Cells[string.Format("AQ{0}", x)].Value = item._36;
                    ws.Cells[string.Format("AR{0}", x)].Value = item._37;
                    ws.Cells[string.Format("AS{0}", x)].Value = item._38;
                    ws.Cells[string.Format("AT{0}", x)].Value = item._39;
                    ws.Cells[string.Format("AU{0}", x)].Value = item._40;
                    ws.Cells[string.Format("AV{0}", x)].Value = item._41;
                    ws.Cells[string.Format("AW{0}", x)].Value = item._42;
                    ws.Cells[string.Format("AX{0}", x)].Value = item._43;
                    ws.Cells[string.Format("AY{0}", x)].Value = item._44;
                    ws.Cells[string.Format("AZ{0}", x)].Value = item._45;
                    ws.Cells[string.Format("BA{0}", x)].Value = item._46;
                    ws.Cells[string.Format("BB{0}", x)].Value = item._47;
                    ws.Cells[string.Format("BC{0}", x)].Value = item._48;
                    ws.Cells[string.Format("BD{0}", x)].Value = item._49;
                    ws.Cells[string.Format("BE{0}", x)].Value = item._50;
                    ws.Cells[string.Format("BF{0}", x)].Value = item._51;
                    ws.Cells[string.Format("BG{0}", x)].Value = item._52;
                    ws.Cells[string.Format("BH{0}", x)].Value = item._1;
                    ws.Cells[string.Format("BI{0}", x)].Value = item._2;
                    ws.Cells[string.Format("BJ{0}", x)].Value = item._3;
                    ws.Cells[string.Format("BK{0}", x)].Value = item._4;
                    ws.Cells[string.Format("BL{0}", x)].Value = item._5;
                    ws.Cells[string.Format("BM{0}", x)].Value = item._6;
                    ws.Cells[string.Format("BN{0}", x)].Value = item._7;
                    ws.Cells[string.Format("BO{0}", x)].Value = item._8;
                    ws.Cells[string.Format("BP{0}", x)].Value = item._9;
                    ws.Cells[string.Format("BQ{0}", x)].Value = item._10;
                    ws.Cells[string.Format("BR{0}", x)].Value = item._11;
                    ws.Cells[string.Format("BS{0}", x)].Value = item._12;
                    ws.Cells[string.Format("BT{0}", x)].Value = item._13;
                    ws.Cells[string.Format("BU{0}", x)].Value = item._14;
                    ws.Cells[string.Format("BV{0}", x)].Value = item._15;
                    ws.Cells[string.Format("BW{0}", x)].Value = item._16;
                    ws.Cells[string.Format("BX{0}", x)].Value = item._17;
                    ws.Cells[string.Format("BY{0}", x)].Value = item._18;
                    ws.Cells[string.Format("BZ{0}", x)].Value = item._19;
                    ws.Cells[string.Format("CA{0}", x)].Value = item._20;
                    ws.Cells[string.Format("CB{0}", x)].Value = item._21;
                    ws.Cells[string.Format("CC{0}", x)].Value = item._22;
                    ws.Cells[string.Format("CD{0}", x)].Value = item._23;
                    ws.Cells[string.Format("CE{0}", x)].Value = item._24;
                    ws.Cells[string.Format("CF{0}", x)].Value = item._25;
                    ws.Cells[string.Format("CG{0}", x)].Value = item._26;
                    x++;
                }
                ws.Cells["A:CG"].AutoFitColumns();
            }
            //FRAMBUESA
            var fram = bd.Database.SqlQuery<ClassCurva>("Select I.Nombre as Asesor, P.Nombre as Productor, S.Cod_Prod, S.Cod_Campo, C.Descripcion as Campo, L.Descripcion as Localidad, isnull(S.Num_corte,0) as Num_corte,S.Sector,round(S.Ha,2) as Ha,isnull(T.Descripcion,'') AS Tipo, isnull(V.Descripcion, '') AS Producto, isnull(S.Numplantas_xha, 0) as Numplantas_xha, isnull(S.Manejo, '') as Manejo, isnull(S.Tipo_plantacion, '') as Tipo_plantacion, isnull(CONVERT(VARCHAR(10), S.Fecha_plantacion, 23), '') as Fecha_plantacion,isnull(CONVERT(VARCHAR(10), S.Fecha_poda, 23), '') as Fecha_poda, " +
              "isnull(CONVERT(VARCHAR(10), S.Fecha_defoliacion, 23), '') as Fecha_defoliacionR, isnull(CONVERT(VARCHAR(10), S.Fecha_corte1, 23), '') as Fecha_corte1R,isnull(CONVERT(VARCHAR(10), S.Fecha_redefoliacion, 23), '') as Fecha_redefoliacionR, isnull(CONVERT(VARCHAR(10), S.Fecha_corte2, 23), '') as Fecha_corte2R, isnull(S.Sem1, '') as Sem1, isnull(S.Sem2, '') as Sem2,isnull(S.Plantacion, 0) as Plantacion, isnull(S.Caja1, 0) as Caja1, isnull(S.Caja2, 0) as Caja2, isnull(S.Estructura, '') as Estructura, isnull(S.Tipo_certificacion, '') as Tipo_certificacion,isnull(S.Tesco, '') as Tesco,isnull(S.Edad_planta, 0) as Edad_planta,isnull(S.Tipo_plantacion2, '') as Tipo_plantacion2,isnull(CONVERT(VARCHAR(10), S.Fecha_podamediacaña, 23), '') as Fecha_podamediacaña,isnull(S.Temporada, '') as Temporada,isnull(Z.DescZona, '') as Zona,isnull(A.Acopio, '') as Acopio," +
              "isnull(S.[27], 0) as _27, isnull(S.[28], 0) as _28, isnull(S.[29], 0) as _29, isnull(S.[30], 0) as _30, isnull(S.[31], 0) as _31, isnull(S.[32], 0) as _32, isnull(S.[33], 0) _33, isnull(S.[34], 0) as _34, isnull(S.[35], 0) as _35,isnull(S.[36], 0) as _36, isnull(S.[37], 0) as _37, isnull(S.[38], 0) as _38, isnull(S.[39], 0) as _39, isnull(S.[40], 0) as _40, isnull(S.[41], 0) as _41, isnull(S.[42], 0) as _42, isnull(S.[43], 0) as _43, isnull(S.[44], 0) as _44, isnull(S.[45], 0) as _45, isnull(S.[46], 0) as _46,isnull(S.[47], 0) as _47, isnull(S.[48], 0) as _48, isnull(S.[49], 0) as _49, isnull(S.[50], 0) as _50, isnull(S.[51], 0) as _51, isnull(S.[52], 0) as _52, isnull(S.[1], 0) as _1, isnull(S.[2], 0) as _2, isnull(S.[3], 0) as _3, isnull(S.[4], 0) as _4, isnull(S.[5], 0) as _5,isnull(S.[6], 0) as _6, isnull(S.[7], 0) as _7, isnull(S.[8], 0) as _8, isnull(S.[9], 0) as _9, isnull(S.[10], 0) as _10, isnull(S.[11], 0) as _11, isnull(S.[12], 0) as _12, isnull(S.[13], 0) as _13, isnull(S.[14], 0) as _14, isnull(S.[15], 0) as _15, isnull(S.[16], 0) as _16,isnull(S.[17], 0) as _17, isnull(S.[18], 0) as _18, isnull(S.[19], 0) as _19, isnull(S.[20], 0) as _20, isnull(S.[21], 0) as _21, isnull(S.[22], 0) as _22, isnull(S.[23], 0) as _23, isnull(S.[24], 0) as _24, isnull(S.[25], 0) as _25, isnull(S.[26], 0) as _26 " +
              "FROM(select V.IdAgen, V.Cod_Prod, V.Cod_Campo, V.Num_corte, V.Sector, V.Ha, V.Numplantas_xha, V.Manejo, V.Tipo_plantacion, V.Fecha_plantacion, V.Fecha_poda, V.Fecha_defoliacion, V.Fecha_corte1, V.Fecha_redefoliacion, V.Fecha_corte2, V.Sem1,V.Sem2, V.Plantacion, V.Caja1, V.Caja2, V.Estructura, V.Tipo_certificacion, V.Tesco, V.Edad_planta, V.Tipo_plantacion2, V.Fecha_podamediacaña, V.Temporada," +
              "max(V.Fecha) as fecha, V.[27], V.[28], V.[29], V.[30], V.[31], V.[32], V.[33], V.[34], V.[35], V.[36], V.[37], V.[38], V.[39], V.[40], V.[41], V.[42], V.[43], V.[44], V.[45], V.[46], V.[47], V.[48], V.[49], V.[50], V.[51], V.[52], V.[1], V.[2], V.[3], V.[4], V.[5], V.[6], V.[7], V.[8], V.[9], V.[10], V.[11], V.[12], V.[13], V.[14], V.[15], V.[16], V.[17], V.[18], V.[19], V.[20], V.[21], V.[22], V.[23], V.[24], V.[25], V.[26] " +
              "From(select * from SIPGProyeccion P where Temporada = (select Temporada from CatSemanas where getdate() between Inicio and Fin)  and Fecha = (select max(Fecha) from SIPGProyeccion))V GROUP BY V.IdAgen,V.Cod_Prod,V.Cod_Campo,V.Num_corte,V.Sector,V.Ha,V.Numplantas_xha,V.Manejo,V.Tipo_plantacion,V.Fecha_plantacion,V.Fecha_poda,V.Fecha_poda,V.Fecha_defoliacion,V.Fecha_corte1,V.Fecha_redefoliacion,V.Fecha_corte2,V.Sem1, V.Sem2,V.Plantacion,V.Caja1,V.Caja2,V.Estructura,V.Tipo_certificacion,V.Tesco,V.Edad_planta,V.Tipo_plantacion2,V.Fecha_podamediacaña,V.Temporada,V.[27],V.[28],V.[29],V.[30],V.[31],V.[32],V.[33],V.[34],V.[35],V.[36],V.[37],V.[38],V.[39],V.[40],V.[41],V.[42],V.[43],V.[44],V.[45],V.[46],V.[47],V.[48],V.[49],V.[50],V.[51],V.[52],V.[1],V.[2],V.[3],V.[4],V.[5],V.[6],V.[7],V.[8],V.[9],V.[10],V.[11],V.[12],V.[13],V.[14],V.[15],V.[16],V.[17],V.[18],V.[19],V.[20],V.[21],V.[22],V.[23],V.[24],V.[25],V.[26])S " +
              "LEFT JOIN ProdAgenteCat I on S.IdAgen = I.IdAgen LEFT JOIN ProdCamposCat C on S.IdAgen = C.IdAgen and S.Cod_Prod = C.Cod_Prod AND S.Cod_Campo = C.Cod_Campo LEFT JOIN ProdProductoresCat P on S.Cod_Prod = P.Cod_Prod LEFT JOIN ProdZonasRastreoCat Z on C.IdZona = Z.IdZona LEFT JOIN CatAcopios A on C.IdAcopio = A.IdAcopio LEFT JOIN CatTiposProd T on C.Tipo = T.Tipo LEFT JOIN CatProductos V on C.Producto = V.Producto AND C.Tipo = V.Tipo LEFT JOIN CatLocalidades L on C.CodLocalidad = L.CodLocalidad " +
              "WHERE I.IdAgen = " + (short)Session["IdAgen"] + " and C.Tipo = 2 order by S.Cod_Prod, S.Cod_Campo, C.Descripcion, L.Descripcion, S.Num_corte,S.Sector").ToList();

            if (fram.Count > 0)
            {
                ExcelWorksheet ws = excel.Workbook.Worksheets.Add("FRAMBUESA");
                ws.Cells["A1"].Value = "CURVA DE PRODUCCION " + fecha_actual.Temporada;
                ws.Cells["A2"].Value = "Ingeniero";
                ws.Cells["B2"].Value = "Codigo";
                ws.Cells["C2"].Value = "Productor";
                ws.Cells["D2"].Value = "Campo";
                ws.Cells["E2"].Value = "Ubicacion";
                ws.Cells["F2"].Value = "Num_corte";
                ws.Cells["G2"].Value = "Sector";
                ws.Cells["H2"].Value = "Superficie";
                ws.Cells["I2"].Value = "Cultivo";
                ws.Cells["J2"].Value = "Variedad";
                ws.Cells["K2"].Value = "Num plantas xha";
                ws.Cells["L2"].Value = "Manejo";
                ws.Cells["M2"].Value = "Tipo_plantacion";
                ws.Cells["N2"].Value = "Fecha_plantacion";
                ws.Cells["O2"].Value = "Fecha_poda";
                ws.Cells["P2"].Value = "Fecha_defoliacion";
                ws.Cells["Q2"].Value = "FechaIniciocorte1";
                ws.Cells["R2"].Value = "Fecha_redefoliacion";
                ws.Cells["S2"].Value = "FechaIniciocorte2";
                ws.Cells["T2"].Value = "Sem.Inic -1ra cosecha";
                ws.Cells["U2"].Value = "Sem.Inic - 2da cosecha";
                ws.Cells["V2"].Value = "Plantacion";
                ws.Cells["W2"].Value = "Caja1";
                ws.Cells["X2"].Value = "Caja2";
                ws.Cells["Y2"].Value = "Estructura";
                ws.Cells["Z2"].Value = "Tipo_certificacion";
                ws.Cells["AA2"].Value = "Programa_Tesco";
                ws.Cells["AB2"].Value = "Edad_planta";
                ws.Cells["AC2"].Value = "Tipo_plantacion2";
                ws.Cells["AD2"].Value = "Fecha_podamediacaña";
                ws.Cells["AE2"].Value = "Temporada";
                ws.Cells["AF2"].Value = "Zona";
                ws.Cells["AG2"].Value = "Acopio";
                ws.Cells["AH2"].Value = "27";
                ws.Cells["AI2"].Value = "28";
                ws.Cells["AJ2"].Value = "29";
                ws.Cells["AK2"].Value = "30";
                ws.Cells["AL2"].Value = "31";
                ws.Cells["AM2"].Value = "32";
                ws.Cells["AN2"].Value = "33";
                ws.Cells["AO2"].Value = "34";
                ws.Cells["AP2"].Value = "35";
                ws.Cells["AQ2"].Value = "36";
                ws.Cells["AR2"].Value = "37";
                ws.Cells["AS2"].Value = "38";
                ws.Cells["AT2"].Value = "39";
                ws.Cells["AU2"].Value = "40";
                ws.Cells["AV2"].Value = "41";
                ws.Cells["AW2"].Value = "42";
                ws.Cells["AX2"].Value = "43";
                ws.Cells["AY2"].Value = "44";
                ws.Cells["AZ2"].Value = "45";
                ws.Cells["BA2"].Value = "46";
                ws.Cells["BB2"].Value = "47";
                ws.Cells["BC2"].Value = "48";
                ws.Cells["BD2"].Value = "49";
                ws.Cells["BE2"].Value = "50";
                ws.Cells["BF2"].Value = "51";
                ws.Cells["BG2"].Value = "52";
                ws.Cells["BH2"].Value = "1";
                ws.Cells["BI2"].Value = "2";
                ws.Cells["BJ2"].Value = "3";
                ws.Cells["BK2"].Value = "4";
                ws.Cells["BL2"].Value = "5";
                ws.Cells["BM2"].Value = "6";
                ws.Cells["BN2"].Value = "7";
                ws.Cells["BO2"].Value = "8";
                ws.Cells["BP2"].Value = "9";
                ws.Cells["BQ2"].Value = "10";
                ws.Cells["BR2"].Value = "11";
                ws.Cells["BS2"].Value = "12";
                ws.Cells["BT2"].Value = "13";
                ws.Cells["BU2"].Value = "14";
                ws.Cells["BV2"].Value = "15";
                ws.Cells["BW2"].Value = "16";
                ws.Cells["BX2"].Value = "17";
                ws.Cells["BY2"].Value = "18";
                ws.Cells["BZ2"].Value = "19";
                ws.Cells["CA2"].Value = "20";
                ws.Cells["CB2"].Value = "21";
                ws.Cells["CC2"].Value = "22";
                ws.Cells["CD2"].Value = "23";
                ws.Cells["CE2"].Value = "24";
                ws.Cells["CF2"].Value = "25";
                ws.Cells["CG2"].Value = "26";
                int x = 3;
                foreach (var item in fram)
                {
                    ws.Cells[string.Format("A{0}", x)].Value = item.Asesor;
                    ws.Cells[string.Format("B{0}", x)].Value = item.Cod_Prod;
                    ws.Cells[string.Format("C{0}", x)].Value = item.Productor;
                    ws.Cells[string.Format("D{0}", x)].Value = item.Cod_Campo;
                    ws.Cells[string.Format("E{0}", x)].Value = item.Localidad;
                    ws.Cells[string.Format("F{0}", x)].Value = item.Num_corte;
                    ws.Cells[string.Format("G{0}", x)].Value = item.Sector;
                    ws.Cells[string.Format("H{0}", x)].Value = item.Ha;
                    ws.Cells[string.Format("I{0}", x)].Value = item.Tipo;
                    ws.Cells[string.Format("J{0}", x)].Value = item.Producto;
                    ws.Cells[string.Format("K{0}", x)].Value = item.Numplantas_xha;
                    ws.Cells[string.Format("L{0}", x)].Value = item.Manejo;
                    ws.Cells[string.Format("M{0}", x)].Value = item.Tipo_plantacion;
                    ws.Cells[string.Format("N{0}", x)].Value = item.Fecha_plantacion;
                    ws.Cells[string.Format("O{0}", x)].Value = item.Fecha_poda;
                    ws.Cells[string.Format("P{0}", x)].Value = item.Fecha_defoliacionR;
                    ws.Cells[string.Format("Q{0}", x)].Value = item.Fecha_corte1R;
                    ws.Cells[string.Format("R{0}", x)].Value = item.Fecha_redefoliacionR;
                    ws.Cells[string.Format("S{0}", x)].Value = item.Fecha_corte2R;
                    ws.Cells[string.Format("T{0}", x)].Value = item.Sem1;
                    ws.Cells[string.Format("U{0}", x)].Value = item.Sem2;
                    ws.Cells[string.Format("V{0}", x)].Value = item.Plantacion;
                    ws.Cells[string.Format("W{0}", x)].Value = item.Caja1;
                    ws.Cells[string.Format("X{0}", x)].Value = item.Caja2;
                    ws.Cells[string.Format("Y{0}", x)].Value = item.Estructura;
                    ws.Cells[string.Format("Z{0}", x)].Value = item.Tipo_certificacion;
                    ws.Cells[string.Format("AA{0}", x)].Value = item.Tesco;
                    ws.Cells[string.Format("AB{0}", x)].Value = item.Edad_planta;
                    ws.Cells[string.Format("AC{0}", x)].Value = item.Tipo_plantacion2;
                    ws.Cells[string.Format("AD{0}", x)].Value = item.Fecha_podamediacaña;
                    ws.Cells[string.Format("AE{0}", x)].Value = item.Temporada;
                    ws.Cells[string.Format("AF{0}", x)].Value = item.Zona;
                    ws.Cells[string.Format("AG{0}", x)].Value = item.Acopio;
                    ws.Cells[string.Format("AH{0}", x)].Value = item._27;
                    ws.Cells[string.Format("AI{0}", x)].Value = item._28;
                    ws.Cells[string.Format("AJ{0}", x)].Value = item._29;
                    ws.Cells[string.Format("AK{0}", x)].Value = item._30;
                    ws.Cells[string.Format("AL{0}", x)].Value = item._31;
                    ws.Cells[string.Format("AM{0}", x)].Value = item._32;
                    ws.Cells[string.Format("AN{0}", x)].Value = item._33;
                    ws.Cells[string.Format("AO{0}", x)].Value = item._34;
                    ws.Cells[string.Format("AP{0}", x)].Value = item._35;
                    ws.Cells[string.Format("AQ{0}", x)].Value = item._36;
                    ws.Cells[string.Format("AR{0}", x)].Value = item._37;
                    ws.Cells[string.Format("AS{0}", x)].Value = item._38;
                    ws.Cells[string.Format("AT{0}", x)].Value = item._39;
                    ws.Cells[string.Format("AU{0}", x)].Value = item._40;
                    ws.Cells[string.Format("AV{0}", x)].Value = item._41;
                    ws.Cells[string.Format("AW{0}", x)].Value = item._42;
                    ws.Cells[string.Format("AX{0}", x)].Value = item._43;
                    ws.Cells[string.Format("AY{0}", x)].Value = item._44;
                    ws.Cells[string.Format("AZ{0}", x)].Value = item._45;
                    ws.Cells[string.Format("BA{0}", x)].Value = item._46;
                    ws.Cells[string.Format("BB{0}", x)].Value = item._47;
                    ws.Cells[string.Format("BC{0}", x)].Value = item._48;
                    ws.Cells[string.Format("BD{0}", x)].Value = item._49;
                    ws.Cells[string.Format("BE{0}", x)].Value = item._50;
                    ws.Cells[string.Format("BF{0}", x)].Value = item._51;
                    ws.Cells[string.Format("BG{0}", x)].Value = item._52;
                    ws.Cells[string.Format("BH{0}", x)].Value = item._1;
                    ws.Cells[string.Format("BI{0}", x)].Value = item._2;
                    ws.Cells[string.Format("BJ{0}", x)].Value = item._3;
                    ws.Cells[string.Format("BK{0}", x)].Value = item._4;
                    ws.Cells[string.Format("BL{0}", x)].Value = item._5;
                    ws.Cells[string.Format("BM{0}", x)].Value = item._6;
                    ws.Cells[string.Format("BN{0}", x)].Value = item._7;
                    ws.Cells[string.Format("BO{0}", x)].Value = item._8;
                    ws.Cells[string.Format("BP{0}", x)].Value = item._9;
                    ws.Cells[string.Format("BQ{0}", x)].Value = item._10;
                    ws.Cells[string.Format("BR{0}", x)].Value = item._11;
                    ws.Cells[string.Format("BS{0}", x)].Value = item._12;
                    ws.Cells[string.Format("BT{0}", x)].Value = item._13;
                    ws.Cells[string.Format("BU{0}", x)].Value = item._14;
                    ws.Cells[string.Format("BV{0}", x)].Value = item._15;
                    ws.Cells[string.Format("BW{0}", x)].Value = item._16;
                    ws.Cells[string.Format("BX{0}", x)].Value = item._17;
                    ws.Cells[string.Format("BY{0}", x)].Value = item._18;
                    ws.Cells[string.Format("BZ{0}", x)].Value = item._19;
                    ws.Cells[string.Format("CA{0}", x)].Value = item._20;
                    ws.Cells[string.Format("CB{0}", x)].Value = item._21;
                    ws.Cells[string.Format("CC{0}", x)].Value = item._22;
                    ws.Cells[string.Format("CD{0}", x)].Value = item._23;
                    ws.Cells[string.Format("CE{0}", x)].Value = item._24;
                    ws.Cells[string.Format("CF{0}", x)].Value = item._25;
                    ws.Cells[string.Format("CG{0}", x)].Value = item._26;
                    x++;
                }
                ws.Cells["A:CG"].AutoFitColumns();
            }

            //ARANDANO
            var aran = bd.Database.SqlQuery<ClassCurva>("Select I.Nombre as Asesor, P.Nombre as Productor, S.Cod_Prod, S.Cod_Campo, C.Descripcion as Campo, L.Descripcion as Localidad, isnull(S.Num_corte,0) as Num_corte,S.Sector,round(S.Ha,2) as Ha,isnull(T.Descripcion,'') AS Tipo, isnull(V.Descripcion, '') AS Producto, isnull(S.Numplantas_xha, 0) as Numplantas_xha, isnull(S.Manejo, '') as Manejo, isnull(S.Tipo_plantacion, '') as Tipo_plantacion, isnull(CONVERT(VARCHAR(10), S.Fecha_plantacion, 23), '') as Fecha_plantacion,isnull(CONVERT(VARCHAR(10), S.Fecha_poda, 23), '') as Fecha_poda, " +
              "isnull(CONVERT(VARCHAR(10), S.Fecha_defoliacion, 23), '') as Fecha_defoliacionR, isnull(CONVERT(VARCHAR(10), S.Fecha_corte1, 23), '') as Fecha_corte1R,isnull(CONVERT(VARCHAR(10), S.Fecha_redefoliacion, 23), '') as Fecha_redefoliacionR, isnull(CONVERT(VARCHAR(10), S.Fecha_corte2, 23), '') as Fecha_corte2R, isnull(S.Sem1, '') as Sem1, isnull(S.Sem2, '') as Sem2,isnull(S.Plantacion, 0) as Plantacion, isnull(S.Caja1, 0) as Caja1, isnull(S.Caja2, 0) as Caja2, isnull(S.Estructura, '') as Estructura, isnull(S.Tipo_certificacion, '') as Tipo_certificacion,isnull(S.Tesco, '') as Tesco,isnull(S.Edad_planta, 0) as Edad_planta,isnull(S.Tipo_plantacion2, '') as Tipo_plantacion2,isnull(CONVERT(VARCHAR(10), S.Fecha_podamediacaña, 23), '') as Fecha_podamediacaña,isnull(S.Temporada, '') as Temporada,isnull(Z.DescZona, '') as Zona,isnull(A.Acopio, '') as Acopio," +
              "isnull(S.[27], 0) as _27, isnull(S.[28], 0) as _28, isnull(S.[29], 0) as _29, isnull(S.[30], 0) as _30, isnull(S.[31], 0) as _31, isnull(S.[32], 0) as _32, isnull(S.[33], 0) _33, isnull(S.[34], 0) as _34, isnull(S.[35], 0) as _35,isnull(S.[36], 0) as _36, isnull(S.[37], 0) as _37, isnull(S.[38], 0) as _38, isnull(S.[39], 0) as _39, isnull(S.[40], 0) as _40, isnull(S.[41], 0) as _41, isnull(S.[42], 0) as _42, isnull(S.[43], 0) as _43, isnull(S.[44], 0) as _44, isnull(S.[45], 0) as _45, isnull(S.[46], 0) as _46,isnull(S.[47], 0) as _47, isnull(S.[48], 0) as _48, isnull(S.[49], 0) as _49, isnull(S.[50], 0) as _50, isnull(S.[51], 0) as _51, isnull(S.[52], 0) as _52, isnull(S.[1], 0) as _1, isnull(S.[2], 0) as _2, isnull(S.[3], 0) as _3, isnull(S.[4], 0) as _4, isnull(S.[5], 0) as _5,isnull(S.[6], 0) as _6, isnull(S.[7], 0) as _7, isnull(S.[8], 0) as _8, isnull(S.[9], 0) as _9, isnull(S.[10], 0) as _10, isnull(S.[11], 0) as _11, isnull(S.[12], 0) as _12, isnull(S.[13], 0) as _13, isnull(S.[14], 0) as _14, isnull(S.[15], 0) as _15, isnull(S.[16], 0) as _16,isnull(S.[17], 0) as _17, isnull(S.[18], 0) as _18, isnull(S.[19], 0) as _19, isnull(S.[20], 0) as _20, isnull(S.[21], 0) as _21, isnull(S.[22], 0) as _22, isnull(S.[23], 0) as _23, isnull(S.[24], 0) as _24, isnull(S.[25], 0) as _25, isnull(S.[26], 0) as _26 " +
              "FROM(select V.IdAgen, V.Cod_Prod, V.Cod_Campo, V.Num_corte, V.Sector, V.Ha, V.Numplantas_xha, V.Manejo, V.Tipo_plantacion, V.Fecha_plantacion, V.Fecha_poda, V.Fecha_defoliacion, V.Fecha_corte1, V.Fecha_redefoliacion, V.Fecha_corte2, V.Sem1,V.Sem2, V.Plantacion, V.Caja1, V.Caja2, V.Estructura, V.Tipo_certificacion, V.Tesco, V.Edad_planta, V.Tipo_plantacion2, V.Fecha_podamediacaña, V.Temporada," +
              "max(V.Fecha) as fecha, V.[27], V.[28], V.[29], V.[30], V.[31], V.[32], V.[33], V.[34], V.[35], V.[36], V.[37], V.[38], V.[39], V.[40], V.[41], V.[42], V.[43], V.[44], V.[45], V.[46], V.[47], V.[48], V.[49], V.[50], V.[51], V.[52], V.[1], V.[2], V.[3], V.[4], V.[5], V.[6], V.[7], V.[8], V.[9], V.[10], V.[11], V.[12], V.[13], V.[14], V.[15], V.[16], V.[17], V.[18], V.[19], V.[20], V.[21], V.[22], V.[23], V.[24], V.[25], V.[26] " +
              "From(select * from SIPGProyeccion P where Temporada = (select Temporada from CatSemanas where getdate() between Inicio and Fin)  and Fecha = (select max(Fecha) from SIPGProyeccion))V GROUP BY V.IdAgen,V.Cod_Prod,V.Cod_Campo,V.Num_corte,V.Sector,V.Ha,V.Numplantas_xha,V.Manejo,V.Tipo_plantacion,V.Fecha_plantacion,V.Fecha_poda,V.Fecha_poda,V.Fecha_defoliacion,V.Fecha_corte1,V.Fecha_redefoliacion,V.Fecha_corte2,V.Sem1, V.Sem2,V.Plantacion,V.Caja1,V.Caja2,V.Estructura,V.Tipo_certificacion,V.Tesco,V.Edad_planta,V.Tipo_plantacion2,V.Fecha_podamediacaña,V.Temporada,V.[27],V.[28],V.[29],V.[30],V.[31],V.[32],V.[33],V.[34],V.[35],V.[36],V.[37],V.[38],V.[39],V.[40],V.[41],V.[42],V.[43],V.[44],V.[45],V.[46],V.[47],V.[48],V.[49],V.[50],V.[51],V.[52],V.[1],V.[2],V.[3],V.[4],V.[5],V.[6],V.[7],V.[8],V.[9],V.[10],V.[11],V.[12],V.[13],V.[14],V.[15],V.[16],V.[17],V.[18],V.[19],V.[20],V.[21],V.[22],V.[23],V.[24],V.[25],V.[26])S " +
              "LEFT JOIN ProdAgenteCat I on S.IdAgen = I.IdAgen LEFT JOIN ProdCamposCat C on S.IdAgen = C.IdAgen and S.Cod_Prod = C.Cod_Prod AND S.Cod_Campo = C.Cod_Campo LEFT JOIN ProdProductoresCat P on S.Cod_Prod = P.Cod_Prod LEFT JOIN ProdZonasRastreoCat Z on C.IdZona = Z.IdZona LEFT JOIN CatAcopios A on C.IdAcopio = A.IdAcopio LEFT JOIN CatTiposProd T on C.Tipo = T.Tipo LEFT JOIN CatProductos V on C.Producto = V.Producto AND C.Tipo = V.Tipo LEFT JOIN CatLocalidades L on C.CodLocalidad = L.CodLocalidad " +
              "WHERE I.IdAgen = " + (short)Session["IdAgen"] + " and C.Tipo = 3 order by S.Cod_Prod, S.Cod_Campo, C.Descripcion, L.Descripcion, S.Num_corte,S.Sector").ToList();

            if (aran.Count > 0)
            {
                ExcelWorksheet ws = excel.Workbook.Worksheets.Add("ARANDANO");
                ws.Cells["A1"].Value = "CURVA DE PRODUCCION " + fecha_actual.Temporada;
                ws.Cells["A2"].Value = "Ingeniero";
                ws.Cells["B2"].Value = "Codigo";
                ws.Cells["C2"].Value = "Productor";
                ws.Cells["D2"].Value = "Campo";
                ws.Cells["E2"].Value = "Ubicacion";
                ws.Cells["F2"].Value = "Num_corte";
                ws.Cells["G2"].Value = "Sector";
                ws.Cells["H2"].Value = "Superficie";
                ws.Cells["I2"].Value = "Cultivo";
                ws.Cells["J2"].Value = "Variedad";
                ws.Cells["K2"].Value = "Num plantas xha";
                ws.Cells["L2"].Value = "Manejo";
                ws.Cells["M2"].Value = "Tipo_plantacion";
                ws.Cells["N2"].Value = "Fecha_plantacion";
                ws.Cells["O2"].Value = "Fecha_poda";
                ws.Cells["P2"].Value = "Fecha_defoliacion";
                ws.Cells["Q2"].Value = "FechaIniciocorte1";
                ws.Cells["R2"].Value = "Fecha_redefoliacion";
                ws.Cells["S2"].Value = "FechaIniciocorte2";
                ws.Cells["T2"].Value = "Sem.Inic -1ra cosecha";
                ws.Cells["U2"].Value = "Sem.Inic - 2da cosecha";
                ws.Cells["V2"].Value = "Plantacion";
                ws.Cells["W2"].Value = "Caja1";
                ws.Cells["X2"].Value = "Caja2";
                ws.Cells["Y2"].Value = "Estructura";
                ws.Cells["Z2"].Value = "Tipo_certificacion";
                ws.Cells["AA2"].Value = "Programa_Tesco";
                ws.Cells["AB2"].Value = "Edad_planta";
                ws.Cells["AC2"].Value = "Tipo_plantacion2";
                ws.Cells["AD2"].Value = "Fecha_podamediacaña";
                ws.Cells["AE2"].Value = "Temporada";
                ws.Cells["AF2"].Value = "Zona";
                ws.Cells["AG2"].Value = "Acopio";
                ws.Cells["AH2"].Value = "27";
                ws.Cells["AI2"].Value = "28";
                ws.Cells["AJ2"].Value = "29";
                ws.Cells["AK2"].Value = "30";
                ws.Cells["AL2"].Value = "31";
                ws.Cells["AM2"].Value = "32";
                ws.Cells["AN2"].Value = "33";
                ws.Cells["AO2"].Value = "34";
                ws.Cells["AP2"].Value = "35";
                ws.Cells["AQ2"].Value = "36";
                ws.Cells["AR2"].Value = "37";
                ws.Cells["AS2"].Value = "38";
                ws.Cells["AT2"].Value = "39";
                ws.Cells["AU2"].Value = "40";
                ws.Cells["AV2"].Value = "41";
                ws.Cells["AW2"].Value = "42";
                ws.Cells["AX2"].Value = "43";
                ws.Cells["AY2"].Value = "44";
                ws.Cells["AZ2"].Value = "45";
                ws.Cells["BA2"].Value = "46";
                ws.Cells["BB2"].Value = "47";
                ws.Cells["BC2"].Value = "48";
                ws.Cells["BD2"].Value = "49";
                ws.Cells["BE2"].Value = "50";
                ws.Cells["BF2"].Value = "51";
                ws.Cells["BG2"].Value = "52";
                ws.Cells["BH2"].Value = "1";
                ws.Cells["BI2"].Value = "2";
                ws.Cells["BJ2"].Value = "3";
                ws.Cells["BK2"].Value = "4";
                ws.Cells["BL2"].Value = "5";
                ws.Cells["BM2"].Value = "6";
                ws.Cells["BN2"].Value = "7";
                ws.Cells["BO2"].Value = "8";
                ws.Cells["BP2"].Value = "9";
                ws.Cells["BQ2"].Value = "10";
                ws.Cells["BR2"].Value = "11";
                ws.Cells["BS2"].Value = "12";
                ws.Cells["BT2"].Value = "13";
                ws.Cells["BU2"].Value = "14";
                ws.Cells["BV2"].Value = "15";
                ws.Cells["BW2"].Value = "16";
                ws.Cells["BX2"].Value = "17";
                ws.Cells["BY2"].Value = "18";
                ws.Cells["BZ2"].Value = "19";
                ws.Cells["CA2"].Value = "20";
                ws.Cells["CB2"].Value = "21";
                ws.Cells["CC2"].Value = "22";
                ws.Cells["CD2"].Value = "23";
                ws.Cells["CE2"].Value = "24";
                ws.Cells["CF2"].Value = "25";
                ws.Cells["CG2"].Value = "26";
                int x = 3;
                foreach (var item in aran)
                {
                    ws.Cells[string.Format("A{0}", x)].Value = item.Asesor;
                    ws.Cells[string.Format("B{0}", x)].Value = item.Cod_Prod;
                    ws.Cells[string.Format("C{0}", x)].Value = item.Productor;
                    ws.Cells[string.Format("D{0}", x)].Value = item.Cod_Campo;
                    ws.Cells[string.Format("E{0}", x)].Value = item.Localidad;
                    ws.Cells[string.Format("F{0}", x)].Value = item.Num_corte;
                    ws.Cells[string.Format("G{0}", x)].Value = item.Sector;
                    ws.Cells[string.Format("H{0}", x)].Value = item.Ha;
                    ws.Cells[string.Format("I{0}", x)].Value = item.Tipo;
                    ws.Cells[string.Format("J{0}", x)].Value = item.Producto;
                    ws.Cells[string.Format("K{0}", x)].Value = item.Numplantas_xha;
                    ws.Cells[string.Format("L{0}", x)].Value = item.Manejo;
                    ws.Cells[string.Format("M{0}", x)].Value = item.Tipo_plantacion;
                    ws.Cells[string.Format("N{0}", x)].Value = item.Fecha_plantacion;
                    ws.Cells[string.Format("O{0}", x)].Value = item.Fecha_poda;
                    ws.Cells[string.Format("P{0}", x)].Value = item.Fecha_defoliacionR;
                    ws.Cells[string.Format("Q{0}", x)].Value = item.Fecha_corte1R;
                    ws.Cells[string.Format("R{0}", x)].Value = item.Fecha_redefoliacionR;
                    ws.Cells[string.Format("S{0}", x)].Value = item.Fecha_corte2R;
                    ws.Cells[string.Format("T{0}", x)].Value = item.Sem1;
                    ws.Cells[string.Format("U{0}", x)].Value = item.Sem2;
                    ws.Cells[string.Format("V{0}", x)].Value = item.Plantacion;
                    ws.Cells[string.Format("W{0}", x)].Value = item.Caja1;
                    ws.Cells[string.Format("X{0}", x)].Value = item.Caja2;
                    ws.Cells[string.Format("Y{0}", x)].Value = item.Estructura;
                    ws.Cells[string.Format("Z{0}", x)].Value = item.Tipo_certificacion;
                    ws.Cells[string.Format("AA{0}", x)].Value = item.Tesco;
                    ws.Cells[string.Format("AB{0}", x)].Value = item.Edad_planta;
                    ws.Cells[string.Format("AC{0}", x)].Value = item.Tipo_plantacion2;
                    ws.Cells[string.Format("AD{0}", x)].Value = item.Fecha_podamediacaña;
                    ws.Cells[string.Format("AE{0}", x)].Value = item.Temporada;
                    ws.Cells[string.Format("AF{0}", x)].Value = item.Zona;
                    ws.Cells[string.Format("AG{0}", x)].Value = item.Acopio;
                    ws.Cells[string.Format("AH{0}", x)].Value = item._27;
                    ws.Cells[string.Format("AI{0}", x)].Value = item._28;
                    ws.Cells[string.Format("AJ{0}", x)].Value = item._29;
                    ws.Cells[string.Format("AK{0}", x)].Value = item._30;
                    ws.Cells[string.Format("AL{0}", x)].Value = item._31;
                    ws.Cells[string.Format("AM{0}", x)].Value = item._32;
                    ws.Cells[string.Format("AN{0}", x)].Value = item._33;
                    ws.Cells[string.Format("AO{0}", x)].Value = item._34;
                    ws.Cells[string.Format("AP{0}", x)].Value = item._35;
                    ws.Cells[string.Format("AQ{0}", x)].Value = item._36;
                    ws.Cells[string.Format("AR{0}", x)].Value = item._37;
                    ws.Cells[string.Format("AS{0}", x)].Value = item._38;
                    ws.Cells[string.Format("AT{0}", x)].Value = item._39;
                    ws.Cells[string.Format("AU{0}", x)].Value = item._40;
                    ws.Cells[string.Format("AV{0}", x)].Value = item._41;
                    ws.Cells[string.Format("AW{0}", x)].Value = item._42;
                    ws.Cells[string.Format("AX{0}", x)].Value = item._43;
                    ws.Cells[string.Format("AY{0}", x)].Value = item._44;
                    ws.Cells[string.Format("AZ{0}", x)].Value = item._45;
                    ws.Cells[string.Format("BA{0}", x)].Value = item._46;
                    ws.Cells[string.Format("BB{0}", x)].Value = item._47;
                    ws.Cells[string.Format("BC{0}", x)].Value = item._48;
                    ws.Cells[string.Format("BD{0}", x)].Value = item._49;
                    ws.Cells[string.Format("BE{0}", x)].Value = item._50;
                    ws.Cells[string.Format("BF{0}", x)].Value = item._51;
                    ws.Cells[string.Format("BG{0}", x)].Value = item._52;
                    ws.Cells[string.Format("BH{0}", x)].Value = item._1;
                    ws.Cells[string.Format("BI{0}", x)].Value = item._2;
                    ws.Cells[string.Format("BJ{0}", x)].Value = item._3;
                    ws.Cells[string.Format("BK{0}", x)].Value = item._4;
                    ws.Cells[string.Format("BL{0}", x)].Value = item._5;
                    ws.Cells[string.Format("BM{0}", x)].Value = item._6;
                    ws.Cells[string.Format("BN{0}", x)].Value = item._7;
                    ws.Cells[string.Format("BO{0}", x)].Value = item._8;
                    ws.Cells[string.Format("BP{0}", x)].Value = item._9;
                    ws.Cells[string.Format("BQ{0}", x)].Value = item._10;
                    ws.Cells[string.Format("BR{0}", x)].Value = item._11;
                    ws.Cells[string.Format("BS{0}", x)].Value = item._12;
                    ws.Cells[string.Format("BT{0}", x)].Value = item._13;
                    ws.Cells[string.Format("BU{0}", x)].Value = item._14;
                    ws.Cells[string.Format("BV{0}", x)].Value = item._15;
                    ws.Cells[string.Format("BW{0}", x)].Value = item._16;
                    ws.Cells[string.Format("BX{0}", x)].Value = item._17;
                    ws.Cells[string.Format("BY{0}", x)].Value = item._18;
                    ws.Cells[string.Format("BZ{0}", x)].Value = item._19;
                    ws.Cells[string.Format("CA{0}", x)].Value = item._20;
                    ws.Cells[string.Format("CB{0}", x)].Value = item._21;
                    ws.Cells[string.Format("CC{0}", x)].Value = item._22;
                    ws.Cells[string.Format("CD{0}", x)].Value = item._23;
                    ws.Cells[string.Format("CE{0}", x)].Value = item._24;
                    ws.Cells[string.Format("CF{0}", x)].Value = item._25;
                    ws.Cells[string.Format("CG{0}", x)].Value = item._26;
                    x++;
                }
                ws.Cells["A:CG"].AutoFitColumns();
            }

            //FRESA
            var fresa = bd.Database.SqlQuery<ClassCurva>("Select I.Nombre as Asesor, P.Nombre as Productor, S.Cod_Prod, S.Cod_Campo, C.Descripcion as Campo, L.Descripcion as Localidad, isnull(S.Num_corte,0) as Num_corte,S.Sector,round(S.Ha,2) as Ha,isnull(T.Descripcion,'') AS Tipo, isnull(V.Descripcion, '') AS Producto, isnull(S.Numplantas_xha, 0) as Numplantas_xha, isnull(S.Manejo, '') as Manejo, isnull(S.Tipo_plantacion, '') as Tipo_plantacion, isnull(CONVERT(VARCHAR(10), S.Fecha_plantacion, 23), '') as Fecha_plantacion,isnull(CONVERT(VARCHAR(10), S.Fecha_poda, 23), '') as Fecha_poda, " +
              "isnull(CONVERT(VARCHAR(10), S.Fecha_defoliacion, 23), '') as Fecha_defoliacionR, isnull(CONVERT(VARCHAR(10), S.Fecha_corte1, 23), '') as Fecha_corte1R,isnull(CONVERT(VARCHAR(10), S.Fecha_redefoliacion, 23), '') as Fecha_redefoliacionR, isnull(CONVERT(VARCHAR(10), S.Fecha_corte2, 23), '') as Fecha_corte2R, isnull(S.Sem1, '') as Sem1, isnull(S.Sem2, '') as Sem2,isnull(S.Plantacion, 0) as Plantacion, isnull(S.Caja1, 0) as Caja1, isnull(S.Caja2, 0) as Caja2, isnull(S.Estructura, '') as Estructura, isnull(S.Tipo_certificacion, '') as Tipo_certificacion,isnull(S.Tesco, '') as Tesco,isnull(S.Edad_planta, 0) as Edad_planta,isnull(S.Tipo_plantacion2, '') as Tipo_plantacion2,isnull(CONVERT(VARCHAR(10), S.Fecha_podamediacaña, 23), '') as Fecha_podamediacaña,isnull(S.Temporada, '') as Temporada,isnull(Z.DescZona, '') as Zona,isnull(A.Acopio, '') as Acopio," +
              "isnull(S.[27], 0) as _27, isnull(S.[28], 0) as _28, isnull(S.[29], 0) as _29, isnull(S.[30], 0) as _30, isnull(S.[31], 0) as _31, isnull(S.[32], 0) as _32, isnull(S.[33], 0) _33, isnull(S.[34], 0) as _34, isnull(S.[35], 0) as _35,isnull(S.[36], 0) as _36, isnull(S.[37], 0) as _37, isnull(S.[38], 0) as _38, isnull(S.[39], 0) as _39, isnull(S.[40], 0) as _40, isnull(S.[41], 0) as _41, isnull(S.[42], 0) as _42, isnull(S.[43], 0) as _43, isnull(S.[44], 0) as _44, isnull(S.[45], 0) as _45, isnull(S.[46], 0) as _46,isnull(S.[47], 0) as _47, isnull(S.[48], 0) as _48, isnull(S.[49], 0) as _49, isnull(S.[50], 0) as _50, isnull(S.[51], 0) as _51, isnull(S.[52], 0) as _52, isnull(S.[1], 0) as _1, isnull(S.[2], 0) as _2, isnull(S.[3], 0) as _3, isnull(S.[4], 0) as _4, isnull(S.[5], 0) as _5,isnull(S.[6], 0) as _6, isnull(S.[7], 0) as _7, isnull(S.[8], 0) as _8, isnull(S.[9], 0) as _9, isnull(S.[10], 0) as _10, isnull(S.[11], 0) as _11, isnull(S.[12], 0) as _12, isnull(S.[13], 0) as _13, isnull(S.[14], 0) as _14, isnull(S.[15], 0) as _15, isnull(S.[16], 0) as _16,isnull(S.[17], 0) as _17, isnull(S.[18], 0) as _18, isnull(S.[19], 0) as _19, isnull(S.[20], 0) as _20, isnull(S.[21], 0) as _21, isnull(S.[22], 0) as _22, isnull(S.[23], 0) as _23, isnull(S.[24], 0) as _24, isnull(S.[25], 0) as _25, isnull(S.[26], 0) as _26 " +
              "FROM(select V.IdAgen, V.Cod_Prod, V.Cod_Campo, V.Num_corte, V.Sector, V.Ha, V.Numplantas_xha, V.Manejo, V.Tipo_plantacion, V.Fecha_plantacion, V.Fecha_poda, V.Fecha_defoliacion, V.Fecha_corte1, V.Fecha_redefoliacion, V.Fecha_corte2, V.Sem1,V.Sem2, V.Plantacion, V.Caja1, V.Caja2, V.Estructura, V.Tipo_certificacion, V.Tesco, V.Edad_planta, V.Tipo_plantacion2, V.Fecha_podamediacaña, V.Temporada," +
              "max(V.Fecha) as fecha, V.[27], V.[28], V.[29], V.[30], V.[31], V.[32], V.[33], V.[34], V.[35], V.[36], V.[37], V.[38], V.[39], V.[40], V.[41], V.[42], V.[43], V.[44], V.[45], V.[46], V.[47], V.[48], V.[49], V.[50], V.[51], V.[52], V.[1], V.[2], V.[3], V.[4], V.[5], V.[6], V.[7], V.[8], V.[9], V.[10], V.[11], V.[12], V.[13], V.[14], V.[15], V.[16], V.[17], V.[18], V.[19], V.[20], V.[21], V.[22], V.[23], V.[24], V.[25], V.[26] " +
              "From(select * from SIPGProyeccion P where Temporada = (select Temporada from CatSemanas where getdate() between Inicio and Fin)  and Fecha = (select max(Fecha) from SIPGProyeccion))V GROUP BY V.IdAgen,V.Cod_Prod,V.Cod_Campo,V.Num_corte,V.Sector,V.Ha,V.Numplantas_xha,V.Manejo,V.Tipo_plantacion,V.Fecha_plantacion,V.Fecha_poda,V.Fecha_poda,V.Fecha_defoliacion,V.Fecha_corte1,V.Fecha_redefoliacion,V.Fecha_corte2,V.Sem1, V.Sem2,V.Plantacion,V.Caja1,V.Caja2,V.Estructura,V.Tipo_certificacion,V.Tesco,V.Edad_planta,V.Tipo_plantacion2,V.Fecha_podamediacaña,V.Temporada,V.[27],V.[28],V.[29],V.[30],V.[31],V.[32],V.[33],V.[34],V.[35],V.[36],V.[37],V.[38],V.[39],V.[40],V.[41],V.[42],V.[43],V.[44],V.[45],V.[46],V.[47],V.[48],V.[49],V.[50],V.[51],V.[52],V.[1],V.[2],V.[3],V.[4],V.[5],V.[6],V.[7],V.[8],V.[9],V.[10],V.[11],V.[12],V.[13],V.[14],V.[15],V.[16],V.[17],V.[18],V.[19],V.[20],V.[21],V.[22],V.[23],V.[24],V.[25],V.[26])S " +
              "LEFT JOIN ProdAgenteCat I on S.IdAgen = I.IdAgen LEFT JOIN ProdCamposCat C on S.IdAgen = C.IdAgen and S.Cod_Prod = C.Cod_Prod AND S.Cod_Campo = C.Cod_Campo LEFT JOIN ProdProductoresCat P on S.Cod_Prod = P.Cod_Prod LEFT JOIN ProdZonasRastreoCat Z on C.IdZona = Z.IdZona LEFT JOIN CatAcopios A on C.IdAcopio = A.IdAcopio LEFT JOIN CatTiposProd T on C.Tipo = T.Tipo LEFT JOIN CatProductos V on C.Producto = V.Producto AND C.Tipo = V.Tipo LEFT JOIN CatLocalidades L on C.CodLocalidad = L.CodLocalidad " +
              "WHERE I.IdAgen = " + (short)Session["IdAgen"] + " and C.Tipo = 4 order by S.Cod_Prod, S.Cod_Campo, C.Descripcion, L.Descripcion, S.Num_corte,S.Sector").ToList();

            if (fresa.Count > 0)
            {
                ExcelWorksheet ws = excel.Workbook.Worksheets.Add("FRESA");
                ws.Cells["A1"].Value = "CURVA DE PRODUCCION " + fecha_actual.Temporada;
                ws.Cells["A2"].Value = "Ingeniero";
                ws.Cells["B2"].Value = "Codigo";
                ws.Cells["C2"].Value = "Productor";
                ws.Cells["D2"].Value = "Campo";
                ws.Cells["E2"].Value = "Ubicacion";
                ws.Cells["F2"].Value = "Num_corte";
                ws.Cells["G2"].Value = "Sector";
                ws.Cells["H2"].Value = "Superficie";
                ws.Cells["I2"].Value = "Cultivo";
                ws.Cells["J2"].Value = "Variedad";
                ws.Cells["K2"].Value = "Num plantas xha";
                ws.Cells["L2"].Value = "Manejo";
                ws.Cells["M2"].Value = "Tipo_plantacion";
                ws.Cells["N2"].Value = "Fecha_plantacion";
                ws.Cells["O2"].Value = "Fecha_poda";
                ws.Cells["P2"].Value = "Fecha_defoliacion";
                ws.Cells["Q2"].Value = "FechaIniciocorte1";
                ws.Cells["R2"].Value = "Fecha_redefoliacion";
                ws.Cells["S2"].Value = "FechaIniciocorte2";
                ws.Cells["T2"].Value = "Sem.Inic -1ra cosecha";
                ws.Cells["U2"].Value = "Sem.Inic - 2da cosecha";
                ws.Cells["V2"].Value = "Plantacion";
                ws.Cells["W2"].Value = "Caja1";
                ws.Cells["X2"].Value = "Caja2";
                ws.Cells["Y2"].Value = "Estructura";
                ws.Cells["Z2"].Value = "Tipo_certificacion";
                ws.Cells["AA2"].Value = "Programa_Tesco";
                ws.Cells["AB2"].Value = "Edad_planta";
                ws.Cells["AC2"].Value = "Tipo_plantacion2";
                ws.Cells["AD2"].Value = "Fecha_podamediacaña";
                ws.Cells["AE2"].Value = "Temporada";
                ws.Cells["AF2"].Value = "Zona";
                ws.Cells["AG2"].Value = "Acopio";
                ws.Cells["AH2"].Value = "27";
                ws.Cells["AI2"].Value = "28";
                ws.Cells["AJ2"].Value = "29";
                ws.Cells["AK2"].Value = "30";
                ws.Cells["AL2"].Value = "31";
                ws.Cells["AM2"].Value = "32";
                ws.Cells["AN2"].Value = "33";
                ws.Cells["AO2"].Value = "34";
                ws.Cells["AP2"].Value = "35";
                ws.Cells["AQ2"].Value = "36";
                ws.Cells["AR2"].Value = "37";
                ws.Cells["AS2"].Value = "38";
                ws.Cells["AT2"].Value = "39";
                ws.Cells["AU2"].Value = "40";
                ws.Cells["AV2"].Value = "41";
                ws.Cells["AW2"].Value = "42";
                ws.Cells["AX2"].Value = "43";
                ws.Cells["AY2"].Value = "44";
                ws.Cells["AZ2"].Value = "45";
                ws.Cells["BA2"].Value = "46";
                ws.Cells["BB2"].Value = "47";
                ws.Cells["BC2"].Value = "48";
                ws.Cells["BD2"].Value = "49";
                ws.Cells["BE2"].Value = "50";
                ws.Cells["BF2"].Value = "51";
                ws.Cells["BG2"].Value = "52";
                ws.Cells["BH2"].Value = "1";
                ws.Cells["BI2"].Value = "2";
                ws.Cells["BJ2"].Value = "3";
                ws.Cells["BK2"].Value = "4";
                ws.Cells["BL2"].Value = "5";
                ws.Cells["BM2"].Value = "6";
                ws.Cells["BN2"].Value = "7";
                ws.Cells["BO2"].Value = "8";
                ws.Cells["BP2"].Value = "9";
                ws.Cells["BQ2"].Value = "10";
                ws.Cells["BR2"].Value = "11";
                ws.Cells["BS2"].Value = "12";
                ws.Cells["BT2"].Value = "13";
                ws.Cells["BU2"].Value = "14";
                ws.Cells["BV2"].Value = "15";
                ws.Cells["BW2"].Value = "16";
                ws.Cells["BX2"].Value = "17";
                ws.Cells["BY2"].Value = "18";
                ws.Cells["BZ2"].Value = "19";
                ws.Cells["CA2"].Value = "20";
                ws.Cells["CB2"].Value = "21";
                ws.Cells["CC2"].Value = "22";
                ws.Cells["CD2"].Value = "23";
                ws.Cells["CE2"].Value = "24";
                ws.Cells["CF2"].Value = "25";
                ws.Cells["CG2"].Value = "26";
                int x = 3;
                foreach (var item in fresa)
                {
                    ws.Cells[string.Format("A{0}", x)].Value = item.Asesor;
                    ws.Cells[string.Format("B{0}", x)].Value = item.Cod_Prod;
                    ws.Cells[string.Format("C{0}", x)].Value = item.Productor;
                    ws.Cells[string.Format("D{0}", x)].Value = item.Cod_Campo;
                    ws.Cells[string.Format("E{0}", x)].Value = item.Localidad;
                    ws.Cells[string.Format("F{0}", x)].Value = item.Num_corte;
                    ws.Cells[string.Format("G{0}", x)].Value = item.Sector;
                    ws.Cells[string.Format("H{0}", x)].Value = item.Ha;
                    ws.Cells[string.Format("I{0}", x)].Value = item.Tipo;
                    ws.Cells[string.Format("J{0}", x)].Value = item.Producto;
                    ws.Cells[string.Format("K{0}", x)].Value = item.Numplantas_xha;
                    ws.Cells[string.Format("L{0}", x)].Value = item.Manejo;
                    ws.Cells[string.Format("M{0}", x)].Value = item.Tipo_plantacion;
                    ws.Cells[string.Format("N{0}", x)].Value = item.Fecha_plantacion;
                    ws.Cells[string.Format("O{0}", x)].Value = item.Fecha_poda;
                    ws.Cells[string.Format("P{0}", x)].Value = item.Fecha_defoliacionR;
                    ws.Cells[string.Format("Q{0}", x)].Value = item.Fecha_corte1R;
                    ws.Cells[string.Format("R{0}", x)].Value = item.Fecha_redefoliacionR;
                    ws.Cells[string.Format("S{0}", x)].Value = item.Fecha_corte2R;
                    ws.Cells[string.Format("T{0}", x)].Value = item.Sem1;
                    ws.Cells[string.Format("U{0}", x)].Value = item.Sem2;
                    ws.Cells[string.Format("V{0}", x)].Value = item.Plantacion;
                    ws.Cells[string.Format("W{0}", x)].Value = item.Caja1;
                    ws.Cells[string.Format("X{0}", x)].Value = item.Caja2;
                    ws.Cells[string.Format("Y{0}", x)].Value = item.Estructura;
                    ws.Cells[string.Format("Z{0}", x)].Value = item.Tipo_certificacion;
                    ws.Cells[string.Format("AA{0}", x)].Value = item.Tesco;
                    ws.Cells[string.Format("AB{0}", x)].Value = item.Edad_planta;
                    ws.Cells[string.Format("AC{0}", x)].Value = item.Tipo_plantacion2;
                    ws.Cells[string.Format("AD{0}", x)].Value = item.Fecha_podamediacaña;
                    ws.Cells[string.Format("AE{0}", x)].Value = item.Temporada;
                    ws.Cells[string.Format("AF{0}", x)].Value = item.Zona;
                    ws.Cells[string.Format("AG{0}", x)].Value = item.Acopio;
                    ws.Cells[string.Format("AH{0}", x)].Value = item._27;
                    ws.Cells[string.Format("AI{0}", x)].Value = item._28;
                    ws.Cells[string.Format("AJ{0}", x)].Value = item._29;
                    ws.Cells[string.Format("AK{0}", x)].Value = item._30;
                    ws.Cells[string.Format("AL{0}", x)].Value = item._31;
                    ws.Cells[string.Format("AM{0}", x)].Value = item._32;
                    ws.Cells[string.Format("AN{0}", x)].Value = item._33;
                    ws.Cells[string.Format("AO{0}", x)].Value = item._34;
                    ws.Cells[string.Format("AP{0}", x)].Value = item._35;
                    ws.Cells[string.Format("AQ{0}", x)].Value = item._36;
                    ws.Cells[string.Format("AR{0}", x)].Value = item._37;
                    ws.Cells[string.Format("AS{0}", x)].Value = item._38;
                    ws.Cells[string.Format("AT{0}", x)].Value = item._39;
                    ws.Cells[string.Format("AU{0}", x)].Value = item._40;
                    ws.Cells[string.Format("AV{0}", x)].Value = item._41;
                    ws.Cells[string.Format("AW{0}", x)].Value = item._42;
                    ws.Cells[string.Format("AX{0}", x)].Value = item._43;
                    ws.Cells[string.Format("AY{0}", x)].Value = item._44;
                    ws.Cells[string.Format("AZ{0}", x)].Value = item._45;
                    ws.Cells[string.Format("BA{0}", x)].Value = item._46;
                    ws.Cells[string.Format("BB{0}", x)].Value = item._47;
                    ws.Cells[string.Format("BC{0}", x)].Value = item._48;
                    ws.Cells[string.Format("BD{0}", x)].Value = item._49;
                    ws.Cells[string.Format("BE{0}", x)].Value = item._50;
                    ws.Cells[string.Format("BF{0}", x)].Value = item._51;
                    ws.Cells[string.Format("BG{0}", x)].Value = item._52;
                    ws.Cells[string.Format("BH{0}", x)].Value = item._1;
                    ws.Cells[string.Format("BI{0}", x)].Value = item._2;
                    ws.Cells[string.Format("BJ{0}", x)].Value = item._3;
                    ws.Cells[string.Format("BK{0}", x)].Value = item._4;
                    ws.Cells[string.Format("BL{0}", x)].Value = item._5;
                    ws.Cells[string.Format("BM{0}", x)].Value = item._6;
                    ws.Cells[string.Format("BN{0}", x)].Value = item._7;
                    ws.Cells[string.Format("BO{0}", x)].Value = item._8;
                    ws.Cells[string.Format("BP{0}", x)].Value = item._9;
                    ws.Cells[string.Format("BQ{0}", x)].Value = item._10;
                    ws.Cells[string.Format("BR{0}", x)].Value = item._11;
                    ws.Cells[string.Format("BS{0}", x)].Value = item._12;
                    ws.Cells[string.Format("BT{0}", x)].Value = item._13;
                    ws.Cells[string.Format("BU{0}", x)].Value = item._14;
                    ws.Cells[string.Format("BV{0}", x)].Value = item._15;
                    ws.Cells[string.Format("BW{0}", x)].Value = item._16;
                    ws.Cells[string.Format("BX{0}", x)].Value = item._17;
                    ws.Cells[string.Format("BY{0}", x)].Value = item._18;
                    ws.Cells[string.Format("BZ{0}", x)].Value = item._19;
                    ws.Cells[string.Format("CA{0}", x)].Value = item._20;
                    ws.Cells[string.Format("CB{0}", x)].Value = item._21;
                    ws.Cells[string.Format("CC{0}", x)].Value = item._22;
                    ws.Cells[string.Format("CD{0}", x)].Value = item._23;
                    ws.Cells[string.Format("CE{0}", x)].Value = item._24;
                    ws.Cells[string.Format("CF{0}", x)].Value = item._25;
                    ws.Cells[string.Format("CG{0}", x)].Value = item._26;
                    x++;
                }
                ws.Cells["A:CG"].AutoFitColumns();
            }
            //ws.Row(2).Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            //ws.Row(2).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(string.Format("white")));

            Response.Clear();
            Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            Response.AddHeader("content-disposition", "attachment: filename=" + "ExcelReport.xlsx");
            Response.BinaryWrite(excel.GetAsByteArray());
            Response.End();
        }

        //isabel
        public ActionResult Activos_Curva(int agente = 0)
        {
            if (Session["Nombre"] != null)
            {
                ViewData["Nombre"] = Session["Nombre"].ToString();
                List<ProdAgenteCat> lst_Agentes = new List<ProdAgenteCat>();
                lst_Agentes = bd.ProdAgenteCat.Where(x => x.Depto == "P" && x.Activo == true).OrderBy(i => i.IdAgen).ToList();
                ViewBag.List_Agente = lst_Agentes;

                CurvaList(agente);
            }
            else
            {
                return RedirectToAction("Index", "Login");
            }
            return View();
        }

        [HttpPost]
        public ActionResult Activos_Curva(int agente = 0, string codigo = "", string check = "")
        {
            JsonResult dtaEjecucionTarea = default(JsonResult);

            //if (Session["Nombre"] != null)
            //{
            //    ViewData["Nombre"] = Session["Nombre"].ToString();               

            //    List<ProdAgenteCat> lst_Agentes = new List<ProdAgenteCat>();
            //    lst_Agentes = bd.ProdAgenteCat.Where(x => x.Depto == "P" && x.Activo==true).OrderBy(i => i.Nombre).ToList();
            //    ViewBag.List_Agente = lst_Agentes;

            CurvaList(agente);

            if (!String.IsNullOrEmpty(check))
            {
                if (Update_estado_M(agente, codigo))
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
            //}
            //else
            //{
            //    return RedirectToAction("Index", "Login");
            //}

            return dtaEjecucionTarea;
        }
        public bool Update_estado_M(int agente = 0, string codigo = "")
        {
            if (agente != 0 && codigo != null)
            {
                var query = from x in bd.SIPGProyeccion
                            where x.IdAgen == agente && x.Cod_Prod == codigo
                            select x;
                foreach (SIPGProyeccion item in query)
                {
                    item.Estado_M = "A";
                    item.Fecha = DateTime.Now;
                }
                bd.SaveChanges();
                return true;
            }
            else
            {
                return false;
            }
        }

    }
}
