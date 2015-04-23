using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Web;

namespace McoApiTool.Models
{
    public class McoApiToolContext : DbContext
    {
        // You can add custom code to this file. Changes will not be overwritten.
        // 
        // If you want Entity Framework to drop and regenerate your database
        // automatically whenever you change your model schema, please use data migrations.
        // For more information refer to the documentation:
        // http://msdn.microsoft.com/en-us/data/jj591621.aspx
    
        public McoApiToolContext() : base("name=McoApiToolContext")
        {
        }

        public System.Data.Entity.DbSet<McoApiTool.Models.Role> Roles { get; set; }

        public System.Data.Entity.DbSet<McoApiTool.Models.User> Users { get; set; }

        public System.Data.Entity.DbSet<McoApiTool.Models.Exchange> Exchanges { get; set; }

        public System.Data.Entity.DbSet<McoApiTool.Models.Setting> Settings { get; set; }

        public System.Data.Entity.DbSet<McoApiTool.Models.Mediary> Mediaries { get; set; }

        public System.Data.Entity.DbSet<McoApiTool.Models.Resource> Resources { get; set; }

        public System.Data.Entity.DbSet<McoApiTool.Models.Feed> Feeds { get; set; }

        public System.Data.Entity.DbSet<McoApiTool.Models.Like> Likes { get; set; }
    
    }
}
