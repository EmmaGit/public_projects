namespace McoApiTool.Models
{
    using System;
    using System.Collections.Generic;

    public partial class Mediary
    {
        public Mediary()
        {
            this.Resources = new HashSet<Resource>();
        }

        public int Id { get; set; }
        public string Name { get; set; }
        public string Settings { get; set; }
        public System.DateTime CreationTime { get; set; }
        public System.DateTime UpdateTime { get; set; }
        public int UserId { get; set; }

        public virtual User User { get; set; }
        public virtual ICollection<Resource> Resources { get; set; }
    }
}