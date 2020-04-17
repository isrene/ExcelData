using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace EmpleadosMan.Models
{
    public class Empleados
    {
        public int Cont { get; set; }
        public string Nombre { get; set; }
        public string Empresa { get; set; }
        public string FIngreso { get; set; }
        public string AfiliadoIMSS { get; set; }
        public string Estatus { get; set; }
        public decimal SDI { get; set; }
        public string NSS { get; set; }
        public string CURP { get; set; }
    }
}