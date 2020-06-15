using System;
using System.Collections.Generic;

namespace GMRestrictionConfigScriptGenerator.Models
{
    public partial class UplSubcategories
    {
        public int Id { get; set; }
        public string Title { get; set; }
        public int Number { get; set; }
        public int IdCtCategories { get; set; }
        public string Notes { get; set; }
    }
}
