//------------------------------------------------------------------------------
// <auto-generated>
//     Este código se generó a partir de una plantilla.
//
//     Los cambios manuales en este archivo pueden causar un comportamiento inesperado de la aplicación.
//     Los cambios manuales en este archivo se sobrescribirán si se regenera el código.
// </auto-generated>
//------------------------------------------------------------------------------

namespace Sistema_Indicadores.Models
{
    using System;
    using System.Collections.Generic;
    
    public partial class Seguimiento_financ
    {
        public int Id { get; set; }
        public string Cod_Prod { get; set; }
        public short Cod_Campo { get; set; }
        public string Comentarios { get; set; }
        public Nullable<System.DateTime> Fecha { get; set; }
        public short Cod_Empresa { get; set; }
        public string Estatus { get; set; }
        public Nullable<short> IdAgen { get; set; }
        public string AP { get; set; }
        public Nullable<System.DateTime> Fecha_Up { get; set; }
    }
}
