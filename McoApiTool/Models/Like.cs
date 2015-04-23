namespace McoApiTool.Models
{
    using System;
    using System.Collections.Generic;

    public partial class Like
    {
        public int Id { get; set; }
        public int UserId { get; set; }
        public int FeedId { get; set; }

        public virtual User User { get; set; }
        public virtual Feed Feed { get; set; }
    }
}
