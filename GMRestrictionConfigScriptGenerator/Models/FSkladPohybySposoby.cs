using System;
using System.Collections.Generic;

namespace GMRestrictionConfigScriptGenerator.Models
{
    public partial class FSkladPohybySposoby
    {
        public int Operacia { get; set; }
        public int Sposob { get; set; }
        public int Povoleny { get; set; }
        public string Skratka { get; set; }
        public string Nazov { get; set; }
        public int? MapovaniePohyb { get; set; }
        public string SqlFilter { get; set; }
        public string SqlFilterPartner { get; set; }
        public int Typ { get; set; }
    }
}
