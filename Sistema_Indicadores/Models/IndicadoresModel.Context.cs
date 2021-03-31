﻿//------------------------------------------------------------------------------
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
    using System.Data.Entity;
    using System.Data.Entity.Infrastructure;
    using System.Data.Entity.Core.Objects;
    using System.Linq;
    
    public partial class SeasonSun1213Entities16 : DbContext
    {
        public SeasonSun1213Entities16()
            : base("name=SeasonSun1213Entities16")
        {
        }
    
        protected override void OnModelCreating(DbModelBuilder modelBuilder)
        {
            throw new UnintentionalCodeFirstException();
        }
    
        public virtual DbSet<CatAcopios> CatAcopios { get; set; }
        public virtual DbSet<CatProductos> CatProductos { get; set; }
        public virtual DbSet<CatSemanas> CatSemanas { get; set; }
        public virtual DbSet<CatTiposProd> CatTiposProd { get; set; }
        public virtual DbSet<ProdAgenteCat> ProdAgenteCat { get; set; }
        public virtual DbSet<ProdCamposCat> ProdCamposCat { get; set; }
        public virtual DbSet<ProdProductoresCat> ProdProductoresCat { get; set; }
        public virtual DbSet<SIPGVisitas> SIPGVisitas { get; set; }
        public virtual DbSet<SIPGComentarios> SIPGComentarios { get; set; }
        public virtual DbSet<ProdVisitasCab> ProdVisitasCab { get; set; }
        public virtual DbSet<Seguimiento_financ> Seguimiento_financ { get; set; }
        public virtual DbSet<SIPGUsuarios> SIPGUsuarios { get; set; }
        public virtual DbSet<ProdProyeccion> ProdProyeccion { get; set; }
        public virtual DbSet<UV_ProdRecepcion> UV_ProdRecepcion { get; set; }
    
        [DbFunction("SeasonSun1213Entities16", "fnRptSaldosFinanciamiento")]
        public virtual IQueryable<fnRptSaldosFinanciamiento_Result> fnRptSaldosFinanciamiento(Nullable<System.DateTime> fechaIni, Nullable<System.DateTime> fechaFin, Nullable<System.DateTime> fecFinanciamiento, Nullable<short> semana)
        {
            var fechaIniParameter = fechaIni.HasValue ?
                new ObjectParameter("FechaIni", fechaIni) :
                new ObjectParameter("FechaIni", typeof(System.DateTime));
    
            var fechaFinParameter = fechaFin.HasValue ?
                new ObjectParameter("FechaFin", fechaFin) :
                new ObjectParameter("FechaFin", typeof(System.DateTime));
    
            var fecFinanciamientoParameter = fecFinanciamiento.HasValue ?
                new ObjectParameter("FecFinanciamiento", fecFinanciamiento) :
                new ObjectParameter("FecFinanciamiento", typeof(System.DateTime));
    
            var semanaParameter = semana.HasValue ?
                new ObjectParameter("Semana", semana) :
                new ObjectParameter("Semana", typeof(short));
    
            return ((IObjectContextAdapter)this).ObjectContext.CreateQuery<fnRptSaldosFinanciamiento_Result>("[SeasonSun1213Entities16].[fnRptSaldosFinanciamiento](@FechaIni, @FechaFin, @FecFinanciamiento, @Semana)", fechaIniParameter, fechaFinParameter, fecFinanciamientoParameter, semanaParameter);
        }
    }
}
