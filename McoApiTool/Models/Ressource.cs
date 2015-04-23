namespace McoApiTool.Models
{
    using System;
    using System.Collections.Generic;
    
    public partial class Resource
    {
        public int Id { get; set; }
        public string Extension { get; set; }
        public string Type { get; set; }
        public string Filename { get; set; }
        public string Path { get; set; }
        public string Size { get; set; }
        public System.DateTime CreationTime { get; set; }
        public System.DateTime UpdateTime { get; set; }
        public int MediaryId { get; set; }
    
        public virtual Mediary Mediary { get; set; }
    }
}
