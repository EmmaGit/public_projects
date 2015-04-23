namespace McoApiTool.Migrations
{
    using System;
    using System.Data.Entity.Migrations;
    
    public partial class Initial : DbMigration
    {
        public override void Up()
        {
            CreateTable(
                "dbo.Exchanges",
                c => new
                    {
                        Id = c.Int(nullable: false, identity: true),
                    })
                .PrimaryKey(t => t.Id);
            
            CreateTable(
                "dbo.Feeds",
                c => new
                    {
                        Id = c.Int(nullable: false, identity: true),
                        Title = c.String(),
                        Content = c.String(),
                        CreationTime = c.DateTime(nullable: false),
                        UpdateTime = c.DateTime(nullable: false),
                        UserId = c.Int(nullable: false),
                        ExchangeId = c.Int(),
                        PostId = c.Int(),
                        Availability = c.Int(),
                        Discriminator = c.String(nullable: false, maxLength: 128),
                    })
                .PrimaryKey(t => t.Id)
                .ForeignKey("dbo.Exchanges", t => t.ExchangeId, cascadeDelete: true)
                .ForeignKey("dbo.Users", t => t.UserId, cascadeDelete: true)
                .ForeignKey("dbo.Feeds", t => t.PostId)
                .Index(t => t.UserId)
                .Index(t => t.ExchangeId)
                .Index(t => t.PostId);
            
            CreateTable(
                "dbo.Likes",
                c => new
                    {
                        Id = c.Int(nullable: false, identity: true),
                        UserId = c.Int(nullable: false),
                        FeedId = c.Int(nullable: false),
                    })
                .PrimaryKey(t => t.Id)
                .ForeignKey("dbo.Feeds", t => t.FeedId, cascadeDelete: false)
                .ForeignKey("dbo.Users", t => t.UserId, cascadeDelete: false)
                .Index(t => t.UserId)
                .Index(t => t.FeedId);
            
            CreateTable(
                "dbo.Users",
                c => new
                    {
                        Id = c.Int(nullable: false, identity: true),
                        Username = c.String(),
                        Email = c.String(),
                        FirstConnection = c.DateTime(nullable: false),
                        LastConnection = c.DateTime(nullable: false),
                        Picture = c.String(),
                        Hostname = c.String(),
                        RoleId = c.Int(nullable: false),
                        Exchange_Id = c.Int(),
                        Setting_Id = c.Int(),
                    })
                .PrimaryKey(t => t.Id)
                .ForeignKey("dbo.Exchanges", t => t.Exchange_Id)
                .ForeignKey("dbo.Roles", t => t.RoleId, cascadeDelete: true)
                .ForeignKey("dbo.Settings", t => t.Setting_Id)
                .Index(t => t.RoleId)
                .Index(t => t.Exchange_Id)
                .Index(t => t.Setting_Id);
            
            CreateTable(
                "dbo.Mediaries",
                c => new
                    {
                        Id = c.Int(nullable: false, identity: true),
                        Name = c.String(),
                        Settings = c.String(),
                        CreationTime = c.DateTime(nullable: false),
                        UpdateTime = c.DateTime(nullable: false),
                        UserId = c.Int(nullable: false),
                    })
                .PrimaryKey(t => t.Id)
                .ForeignKey("dbo.Users", t => t.UserId, cascadeDelete: true)
                .Index(t => t.UserId);
            
            CreateTable(
                "dbo.Resources",
                c => new
                    {
                        Id = c.Int(nullable: false, identity: true),
                        Extension = c.String(),
                        Type = c.String(),
                        Filename = c.String(),
                        Path = c.String(),
                        Size = c.String(),
                        CreationTime = c.DateTime(nullable: false),
                        UpdateTime = c.DateTime(nullable: false),
                        MediaryId = c.Int(nullable: false),
                    })
                .PrimaryKey(t => t.Id)
                .ForeignKey("dbo.Mediaries", t => t.MediaryId, cascadeDelete: true)
                .Index(t => t.MediaryId);
            
            CreateTable(
                "dbo.Roles",
                c => new
                    {
                        Id = c.Int(nullable: false, identity: true),
                        Name = c.String(),
                    })
                .PrimaryKey(t => t.Id);
            
            CreateTable(
                "dbo.Settings",
                c => new
                    {
                        Id = c.Int(nullable: false, identity: true),
                        Color = c.String(),
                        Language = c.String(),
                    })
                .PrimaryKey(t => t.Id);
            
        }
        
        public override void Down()
        {
            DropForeignKey("dbo.Feeds", "PostId", "dbo.Feeds");
            DropForeignKey("dbo.Users", "Setting_Id", "dbo.Settings");
            DropForeignKey("dbo.Users", "RoleId", "dbo.Roles");
            DropForeignKey("dbo.Mediaries", "UserId", "dbo.Users");
            DropForeignKey("dbo.Resources", "MediaryId", "dbo.Mediaries");
            DropForeignKey("dbo.Likes", "UserId", "dbo.Users");
            DropForeignKey("dbo.Feeds", "UserId", "dbo.Users");
            DropForeignKey("dbo.Users", "Exchange_Id", "dbo.Exchanges");
            DropForeignKey("dbo.Likes", "FeedId", "dbo.Feeds");
            DropForeignKey("dbo.Feeds", "ExchangeId", "dbo.Exchanges");
            DropIndex("dbo.Resources", new[] { "MediaryId" });
            DropIndex("dbo.Mediaries", new[] { "UserId" });
            DropIndex("dbo.Users", new[] { "Setting_Id" });
            DropIndex("dbo.Users", new[] { "Exchange_Id" });
            DropIndex("dbo.Users", new[] { "RoleId" });
            DropIndex("dbo.Likes", new[] { "FeedId" });
            DropIndex("dbo.Likes", new[] { "UserId" });
            DropIndex("dbo.Feeds", new[] { "PostId" });
            DropIndex("dbo.Feeds", new[] { "ExchangeId" });
            DropIndex("dbo.Feeds", new[] { "UserId" });
            DropTable("dbo.Settings");
            DropTable("dbo.Roles");
            DropTable("dbo.Resources");
            DropTable("dbo.Mediaries");
            DropTable("dbo.Users");
            DropTable("dbo.Likes");
            DropTable("dbo.Feeds");
            DropTable("dbo.Exchanges");
        }
    }
}
