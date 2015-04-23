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
    public class FeedsController : ApiController
    {
        private McoApiToolContext db = new McoApiToolContext();

        // GET: api/Feeds
        public IQueryable<Feed> GetFeeds()
        {
            return db.Feeds;
        }

        // GET: api/Feeds/5
        [ResponseType(typeof(Feed))]
        public async Task<IHttpActionResult> GetFeed(int id)
        {
            Feed feed = await db.Feeds.FindAsync(id);
            if (feed == null)
            {
                return NotFound();
            }

            return Ok(feed);
        }

        // PUT: api/Feeds/5
        [ResponseType(typeof(void))]
        public async Task<IHttpActionResult> PutFeed(int id, Feed feed)
        {
            if (!ModelState.IsValid)
            {
                return BadRequest(ModelState);
            }

            if (id != feed.Id)
            {
                return BadRequest();
            }

            db.Entry(feed).State = EntityState.Modified;

            try
            {
                await db.SaveChangesAsync();
            }
            catch (DbUpdateConcurrencyException)
            {
                if (!FeedExists(id))
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

        // POST: api/Feeds
        [ResponseType(typeof(Feed))]
        public async Task<IHttpActionResult> PostFeed(Feed feed)
        {
            if (!ModelState.IsValid)
            {
                return BadRequest(ModelState);
            }

            db.Feeds.Add(feed);
            await db.SaveChangesAsync();

            return CreatedAtRoute("DefaultApi", new { id = feed.Id }, feed);
        }

        // DELETE: api/Feeds/5
        [ResponseType(typeof(Feed))]
        public async Task<IHttpActionResult> DeleteFeed(int id)
        {
            Feed feed = await db.Feeds.FindAsync(id);
            if (feed == null)
            {
                return NotFound();
            }

            db.Feeds.Remove(feed);
            await db.SaveChangesAsync();

            return Ok(feed);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                db.Dispose();
            }
            base.Dispose(disposing);
        }

        private bool FeedExists(int id)
        {
            return db.Feeds.Count(e => e.Id == id) > 0;
        }
    }
}