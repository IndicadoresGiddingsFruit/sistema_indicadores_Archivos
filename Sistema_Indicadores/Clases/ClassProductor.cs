using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Sistema_Indicadores.Clases
{
    public class ClassProductor
    {
        public int Id { get; set; }
        public short? IdAgen { get; set; }
        public string Asesor { get; set; }
        public string Cod_Prod { get; set; }
        public string Productor { get; set; }
        public string Estatus { get; set; }
        public string Comentarios { get; set; }
        public Int16? Cod_Campo { get; set; }
        public string Campo { get; set; }
        public DateTime? Fecha { get; set; }
        public int? dias { get; set; }
        public double? cjs1 { get; set; }
        public double? cjs2 { get; set; }
    }
}