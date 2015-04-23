namespace McoApiTool.Models
{
    using System;
    using System.Collections.Generic;

    public partial class Exchange
    {
        public Exchange()
        {
            this.Messages = new HashSet<Message>();
            this.Recipients = new HashSet<User>();
        }

        public int Id { get; set; }

        public virtual ICollection<Message> Messages { get; set; }
        public virtual ICollection<User> Recipients { get; set; }
    }
}