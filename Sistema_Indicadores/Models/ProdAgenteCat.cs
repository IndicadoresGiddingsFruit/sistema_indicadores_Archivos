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
    
    public partial class ProdAgenteCat
    {
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2214:DoNotCallOverridableMethodsInConstructors")]
        public ProdAgenteCat()
        {
            this.ProdMuestreo = new HashSet<ProdMuestreo>();
            this.ProdMuestreo1 = new HashSet<ProdMuestreo>();
            this.ProdProductoresCat = new HashSet<ProdProductoresCat>();
        }
    
        public short IdAgen { get; set; }
        public string Nombre { get; set; }
        public string Depto { get; set; }
        public string Abrev { get; set; }
        public Nullable<bool> Activo { get; set; }
        public short IdRegion { get; set; }
        public string Codigo { get; set; }
    
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<ProdMuestreo> ProdMuestreo { get; set; }
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<ProdMuestreo> ProdMuestreo1 { get; set; }
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<ProdProductoresCat> ProdProductoresCat { get; set; }
    }
}
