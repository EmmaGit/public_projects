namespace McoApiTool.Models
{
    using System;
    using System.Collections.Generic;

    public partial class Post : Feed
    {
        public Post()
        {
            this.Comments = new HashSet<Comment>();
        }

        public int Availability { get; set; }

        public virtual ICollection<Comment> Comments { get; set; }
    }
}