namespace McoApiTool.Models
{
    using System;
    using System.Collections.Generic;

    public partial class Comment : Feed
    {
        public int PostId { get; set; }

        public virtual Post Post { get; set; }
    }
}
