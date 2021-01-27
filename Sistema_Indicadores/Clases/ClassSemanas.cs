using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Sistema_Indicadores.Clases
{
    public class ClassSemanas
    {
        public int Id { get; set; }
        public int Semana { get; set; }
        public string Cantidad { get; set; }

        public ClassSemanas ShallowCopy() => (ClassSemanas)this.MemberwiseClone();
    }
}
