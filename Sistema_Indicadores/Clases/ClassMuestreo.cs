using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Sistema_Indicadores.Clases
{
    public class ClassMuestreo
    {
        public int? IdMuestreo { get; set; }
        public int? IdAnalisis_Residuo { get; set; }
        public string Asesor { get; set; }
        public string AsesorC { get; set; }
        public string AsesorCS { get; set; }
        public string AsesorI { get; set; }
        public string Cod_Prod { get; set; }
        public string Productor { get; set; }
        public Int16? Cod_Empresa { get; set; }
        public Int16? Cod_Campo { get; set; }
        public string Campo { get; set; }
        public decimal? Ha { get; set; }
        public string Ubicacion { get; set; }
        public string Telefono { get; set; }
        public string Liberacion { get; set; }
        public string Estatus { get; set; }
        public string Calidad_fruta { get; set; }
        public string Tarjeta { get; set; }
        public int? IdSector { get; set; }
        public Int16? Sector { get; set; }
        public string CodZona { get; set; }
        public string DescZona { get; set; } 
        public string Tipo { get; set; }
        public string Producto { get; set; }
        public string Zona { get; set; }
        public DateTime? Fecha_envio { get; set; }
        public DateTime? Fecha_entrega { get; set; }
        public DateTime? Fecha_solicitud { get; set; }
        public DateTime? Fecha_ejecucion { get; set; }
        public int? Num_analisis { get; set; }
        public string Laboratorio { get; set; }
        public string Comentarios { get; set; }
        public string Trazas { get; set; }
        public Int16? IdAgen { get; set; }
        public Int16? IdAgenC { get; set; }
        public Int16? IdAgenI { get; set; }
        public DateTime? Fecha { get; set; }
        public DateTime? Inicio_cosecha { get; set; }
        public string Incidencia { get; set; }
        public string Propuesta { get; set; }
        public Int16? IdAgen_Tarjeta { get; set; }
        public string Liberar_Tarjeta { get; set; }
        public string Fecha_real { get; set; }
        public int? Dias { get; set; }
        public DateTime? Fecha_analisis { get; set; }
        public DateTime? LiberacionUSA { get; set; }
        public DateTime? LiberacionEU { get; set; }
        public string Analisis { get; set; }
        public int? proximos_liberar { get; set; }
        public string Compras_oportunidad { get; set; }
 
        public string SectorList { get; set; }
        public int? IdRegion { get; set; }
        public string Folio { get; set; }
    }
}