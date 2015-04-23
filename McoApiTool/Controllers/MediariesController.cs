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
    public class MediariesController : ApiController
    {
        private McoApiToolContext db = new McoApiToolContext();

        // GET: api/Mediaries
        public IQueryable<Mediary> GetMediaries()
        {
            return db.Mediaries;
        }

        // GET: api/Mediaries/5
        [ResponseType(typeof(Mediary))]
        public async Task<IHttpActionResult> GetMediary(int id)
        {
            Mediary mediary = await db.Mediaries.FindAsync(id);
            if (mediary == null)
            {
                return NotFound();
            }

            return Ok(mediary);
        }

        // PUT: api/Mediaries/5
        [ResponseType(typeof(void))]
        public async Task<IHttpActionResult> PutMediary(int id, Mediary mediary)
        {
            if (!ModelState.IsValid)
            {
                return BadRequest(ModelState);
            }

            if (id != mediary.Id)
            {
                return BadRequest();
            }

            db.Entry(mediary).State = EntityState.Modified;

            try
            {
                await db.SaveChangesAsync();
            }
            catch (DbUpdateConcurrencyException)
            {
                if (!MediaryExists(id))
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

        // POST: api/Mediaries
        [ResponseType(typeof(Mediary))]
        public async Task<IHttpActionResult> PostMediary(Mediary mediary)
        {
            if (!ModelState.IsValid)
            {
                return BadRequest(ModelState);
            }

            db.Mediaries.Add(mediary);
            await db.SaveChangesAsync();

            return CreatedAtRoute("DefaultApi", new { id = mediary.Id }, mediary);
        }

        // DELETE: api/Mediaries/5
        [ResponseType(typeof(Mediary))]
        public async Task<IHttpActionResult> DeleteMediary(int id)
        {
            Mediary mediary = await db.Mediaries.FindAsync(id);
            if (mediary == null)
            {
                return NotFound();
            }

            db.Mediaries.Remove(mediary);
            await db.SaveChangesAsync();

            return Ok(mediary);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                db.Dispose();
            }
            base.Dispose(disposing);
        }

        private bool MediaryExists(int id)
        {
            return db.Mediaries.Count(e => e.Id == id) > 0;
        }
    }
}