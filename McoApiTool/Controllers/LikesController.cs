using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.Data.Entity.Infrastructure;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using System.Web.Http;
using System.Web.Http.Description;
using McoApiTool.Models;

namespace McoApiTool.Controllers
{
    public class LikesController : ApiController
    {
        private McoApiToolContext db = new McoApiToolContext();

        // GET: api/Likes
        public IQueryable<Like> GetLikes()
        {
            return db.Likes;
        }

        // GET: api/Likes/5
        [ResponseType(typeof(Like))]
        public async Task<IHttpActionResult> GetLike(int id)
        {
            Like like = await db.Likes.FindAsync(id);
            if (like == null)
            {
                return NotFound();
            }

            return Ok(like);
        }

        // PUT: api/Likes/5
        [ResponseType(typeof(void))]
        public async Task<IHttpActionResult> PutLike(int id, Like like)
        {
            if (!ModelState.IsValid)
            {
                return BadRequest(ModelState);
            }

            if (id != like.Id)
            {
                return BadRequest();
            }

            db.Entry(like).State = EntityState.Modified;

            try
            {
                await db.SaveChangesAsync();
            }
            catch (DbUpdateConcurrencyException)
            {
                if (!LikeExists(id))
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

        // POST: api/Likes
        [ResponseType(typeof(Like))]
        public async Task<IHttpActionResult> PostLike(Like like)
        {
            if (!ModelState.IsValid)
            {
                return BadRequest(ModelState);
            }

            db.Likes.Add(like);
            await db.SaveChangesAsync();

            return CreatedAtRoute("DefaultApi", new { id = like.Id }, like);
        }

        // DELETE: api/Likes/5
        [ResponseType(typeof(Like))]
        public async Task<IHttpActionResult> DeleteLike(int id)
        {
            Like like = await db.Likes.FindAsync(id);
            if (like == null)
            {
                return NotFound();
            }

            db.Likes.Remove(like);
            await db.SaveChangesAsync();

            return Ok(like);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                db.Dispose();
            }
            base.Dispose(disposing);
        }

        private bool LikeExists(int id)
        {
            return db.Likes.Count(e => e.Id == id) > 0;
        }
    }
}