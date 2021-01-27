using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Sistema_Indicadores.Clases
{
    public class ClassCurva
    {
        public int Id_Proyeccion { get; set; }
        public int Id_Semana{ get; set; }       
        public short? IdAgen { get; set; }
        public string Asesor { get; set; }        
        public string Cod_Prod { get; set; }        
        public string Productor { get; set; }
        public Int16? Cod_Campo { get; set; }
        public string Campo { get; set; }
        public string Localidad { get; set; }
        public int? Sector { get; set; }
        public double? Ha { get; set; }
        public string Tipo { get; set; }
        public string Producto { get; set; }
        public string DescTipo { get; set; }
        public string DescProducto { get; set; }      
        public double? PronosticoT { get; set; }
        public double? DiferenciaT { get; set; }
        public double? PronosticoA { get; set; }
        public double? EntregadoT { get; set; }
        public double? DiferenciaA { get; set; }
        public byte? Semana { get; set; }
        public double? Cantidad { get; set; }
        public double? PronosticoSA { get; set; }
        public double? EntregadoSA { get; set; }
        public double? DiferenciaSA { get; set; }
        public string Fecha { get; set; }
        public string SaldoFinal { get; set; }
        public DateTime? Fecha_defoliacion { get; set; }
        public DateTime? Fecha_corte1 { get; set; }
        public DateTime? Fecha_redefoliacion { get; set; }
        public DateTime? Fecha_corte2 { get; set; }
        public DateTime? Fecha_Analisis { get; set; }
        public DateTime? Inicio { get; set; }

        public string Fecha_defoliacionR { get; set; }
        public string Fecha_corte1R { get; set; }
        public string Fecha_redefoliacionR { get; set; }
        public string Fecha_corte2R { get; set; }

        public int? Num_corte { get; set; }
        public string Numplantas_xha { get; set; }
        public string Manejo { get; set; }
        public string Tipo_plantacion { get; set; }
        public string Fecha_plantacion { get; set; }
        public string Fecha_poda { get; set; }
        public int? Sem1 { get; set; }
        public int? Sem2 { get; set; }
        public string Plantacion { get; set; }
        public double Caja1 { get; set; }
        public double Caja2 { get; set; }
        public string Estructura { get; set; }
        public string Tipo_certificacion { get; set; }
        public string Tesco { get; set; }
        public string Edad_planta { get; set; }
        public string Tipo_plantacion2 { get; set; }
        public string Fecha_podamediacaña { get; set; }
        public string Temporada { get; set; }
        public string Zona { get; set; }
        public string Acopio { get; set; }
        public double? _27 { get; set; }
        public double? _28 { get; set; }
        public double? _29 { get; set; }
        public double? _30 { get; set; }
        public double? _31 { get; set; }
        public double? _32 { get; set; }
        public double? _33 { get; set; }
        public double? _34 { get; set; }
        public double? _35 { get; set; }
        public double? _36 { get; set; }
        public double? _37 { get; set; }
        public double? _38 { get; set; }
        public double? _39 { get; set; }
        public double? _40 { get; set; }
        public double? _41 { get; set; }
        public double? _42 { get; set; }
        public double? _43 { get; set; }
        public double? _44 { get; set; }
        public double? _45 { get; set; }
        public double? _46 { get; set; }
        public double? _47 { get; set; }
        public double? _48 { get; set; }
        public double? _49 { get; set; }
        public double? _50 { get; set; }
        public double? _51 { get; set; }
        public double? _52 { get; set; }
        public double? _53 { get; set; }
        public double? _1 { get; set; }
        public double? _2 { get; set; }
        public double? _3 { get; set; }
        public double? _4 { get; set; }
        public double? _5 { get; set; }
        public double? _6 { get; set; }
        public double? _7 { get; set; }
        public double? _8 { get; set; }
        public double? _9 { get; set; }
        public double? _10 { get; set; }
        public double? _11 { get; set; }
        public double? _12 { get; set; }
        public double? _13 { get; set; }
        public double? _14 { get; set; }
        public double? _15 { get; set; }
        public double? _16 { get; set; }
        public double? _17 { get; set; }
        public double? _18 { get; set; }
        public double? _19 { get; set; }
        public double? _20 { get; set; }
        public double? _21 { get; set; }
        public double? _22 { get; set; }
        public double? _23 { get; set; }
        public double? _24 { get; set; }
        public double? _25 { get; set; }
        public double? _26 { get; set; }
    }
}