using AutoMapper;
using AutoMapper.QueryableExtensions;
using McoApiTool.Models;
using System;
using System.Data.Entity;
using System.Data.Entity.Infrastructure;
using System.Linq;
using System.Net;
using System.Threading.Tasks;
using System.Web.Http;
using System.Web.Http.Description;
using System.Web.Mvc;

namespace McoApiTool.Controllers
{
    public class UsersController : ApiController
    {
        private McoApiToolContext db = new McoApiToolContext();

        // GET: api/Users
        public IQueryable<UserDTO> GetUsers()
        {
            Mapper.CreateMap<User, UserDTO>()
                .ForMember(dto => dto.Role, conf => conf.MapFrom(ol => ol.Role.Name));
            return db.Users.Project().To<UserDTO>();
        }

        // GET: api/Users/5
        [ResponseType(typeof(UserDTO))]
        public async Task<IHttpActionResult> GetUser(int id)
        {
            User user = await db.Users.FindAsync(id);
            if (user == null)
            {
                return NotFound();
            }
            Mapper.CreateMap<User, UserDTO>();
            UserDTO dto = Mapper.Map<User, UserDTO>(user);
            return Ok(dto);
        }

        // PUT: api/Users/5
        [ResponseType(typeof(void))]
        public async Task<IHttpActionResult> PutUser(int id, User user)
        {
            if (!ModelState.IsValid)
            {
                return BadRequest(ModelState);
            }

            if (id != user.Id)
            {
                return BadRequest();
            }

            db.Entry(user).State = EntityState.Modified;

            try
            {
                await db.SaveChangesAsync();
            }
            catch (DbUpdateConcurrencyException)
            {
                if (!UserExists(id))
                {
                    return NotFound();
                }
                else
                {
                    throw;
                }
            }

            return StatusCode(HttpStatusCode.NoContent);
        }

        // POST: api/Users
        [ResponseType(typeof(UserDTO))]
        public async Task<IHttpActionResult> PostUser(User user)
        {
            if (!ModelState.IsValid)
            {
                return BadRequest(ModelState);
            }
            user.FirstConnection = DateTime.Now;
            user.LastConnection = DateTime.Now;
            user.Hostname = (user.Hostname == null || user.Hostname.Trim() == "") ? 
                User.Identity.Name : user.Hostname.ToUpper();
            user.RoleId = (user.RoleId == 0) ? 
                db.Roles.Where(r => r.Name == "Standard").FirstOrDefault().Id: user.RoleId;
            db.Users.Add(user);
            try { await db.SaveChangesAsync(); }
            catch (Exception exception) 
            {
            
            }
            Mapper.CreateMap<User, UserDTO>();
            UserDTO dto = Mapper.Map<User, UserDTO>(user);
            return CreatedAtRoute("DefaultApi", new { id = dto.Id }, dto);
        }

        // DELETE: api/Users/5
        [ResponseType(typeof(User))]
        public async Task<IHttpActionResult> DeleteUser(int id)
        {
            User user = await db.Users.FindAsync(id);
            if (user == null)
            {
                return NotFound();
            }

            db.Users.Remove(user);
            await db.SaveChangesAsync();

            return Ok(user);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                db.Dispose();
            }
            base.Dispose(disposing);
        }

        private bool UserExists(int id)
        {
            return db.Users.Count(e => e.Id == id) > 0;
        }
    }
}