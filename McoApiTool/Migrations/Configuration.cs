namespace McoApiTool.Migrations
{
    using System;
    using System.Data.Entity;
    using System.Data.Entity.Migrations;
    using System.Linq;
    using McoApiTool.Models;

    internal sealed class Configuration : DbMigrationsConfiguration<McoApiTool.Models.McoApiToolContext>
    {
        public Configuration()
        {
            AutomaticMigrationsEnabled = false;
        }

        protected override void Seed(McoApiTool.Models.McoApiToolContext context)
        {
            context.Roles.AddOrUpdate(x => x.Id,
                new Role() { Id = 1, Name = "WebMaster" },
                new Role() { Id = 2, Name = "Administrator" },
                new Role() { Id = 3, Name = "Standard" },
                new Role() { Id = 4, Name = "Visitor" }
            );
        }
    }
}
