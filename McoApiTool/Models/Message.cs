namespace McoApiTool.Models
{
    using System;
    using System.Collections.Generic;

    public partial class Message : Feed
    {
        public int ExchangeId { get; set; }

        public virtual Exchange Exchange { get; set; }
    }
}