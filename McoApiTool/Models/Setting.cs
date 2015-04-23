namespace McoApiTool.Models
{
    using System;
    using System.Collections.Generic;

    public partial class Setting
    {
        public int Id { get; set; }
        public string Color { get; set; }
        public string Language { get; set; }

        //public virtual User User { get; set; }
    }
}
