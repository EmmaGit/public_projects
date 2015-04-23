namespace McoApiTool.Models
{
    using System;
    using System.Collections.Generic;

    public partial class Role
    {
        public Role()
        {
            //this.Users = new HashSet<User>();
        }

        public int Id { get; set; }
        public string Name { get; set; }

        //public virtual ICollection<User> Users { get; set; }
    }
}