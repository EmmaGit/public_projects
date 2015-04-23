namespace McoApiTool.Models
{
    using System;
    using System.Collections.Generic;

    public partial class Feed
    {
        public Feed()
        {
            this.Likes = new HashSet<Like>();
        }

        public int Id { get; set; }
        public string Title { get; set; }
        public string Content { get; set; }
        public System.DateTime CreationTime { get; set; }
        public System.DateTime UpdateTime { get; set; }
        public int UserId { get; set; }

        public virtual User User { get; set; }
        public virtual ICollection<Like> Likes { get; set; }
    }
}
