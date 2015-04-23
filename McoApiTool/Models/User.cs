namespace McoApiTool.Models
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;

    public partial class User
    {
        public User()
        {
            this.Mediaries = new HashSet<Mediary>();
            this.Feeds = new HashSet<Feed>();
            this.Likes = new HashSet<Like>();
        }
        public int Id { get; set; }
        public string Username { get; set; }
        public string Email { get; set; }
        public DateTime FirstConnection { get; set; }
        public DateTime LastConnection { get; set; }
        public string Picture { get; set; }
        public string Hostname { get; set; }
        public int RoleId { get; set; }

        public virtual Role Role { get; set; }
        public virtual Setting Setting { get; set; }
        public virtual ICollection<Mediary> Mediaries { get; set; }
        public virtual ICollection<Feed> Feeds { get; set; }
        public virtual ICollection<Like> Likes { get; set; }
        public virtual Exchange Exchange { get; set; }
    }

    public class UserDTO
    {
        public int Id { get; set; }
        public string Username { get; set; }
        public string Email { get; set; }
        public DateTime FirstConnection { get; set; }
        public DateTime LastConnection { get; set; }
        public string Picture { get; set; }
        public string Hostname { get; set; }
        public string Role { get; set; }
    }
}
