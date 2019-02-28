using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.Linq;
using System.Threading.Tasks;
using System.Net;
using System.Web;
using System.Web.Mvc;
using FinchInventory.Models;
using FinchInventory.Controllers;

namespace Finch_Inventory.Controllers
{
    public class GoalsController : BaseController
    {
        private FinchDbContext db = new FinchDbContext();

        // GET: Goals
        public async Task<ActionResult> Index()
        {
            var goals = db.Goals.Include(g => g.Machine).Include(g => g.Position);
            return View(await goals.ToListAsync());
        }

        // GET: Goals/Details/5
        public async Task<ActionResult> Details(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Goal goal = await db.Goals.FindAsync(id);
            if (goal == null)
            {
                return HttpNotFound();
            }
            return View(goal);
        }

        // GET: Goals/Create
        public ActionResult Create()
        {
            ViewBag.PM_ID = new SelectList(db.Machines, "ID", "Machine1");
            ViewBag.PositionID = new SelectList(db.Positions, "ID", "Position1");
            return View();
        }

        // POST: Goals/Create
        // To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        // more details see https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<ActionResult> Create([Bind(Include = "ID,PositionID,PM_ID,Goal1")] Goal goal)
        {
            if (ModelState.IsValid)
            {
                db.Goals.Add(goal);
                await db.SaveChangesAsync();
                return RedirectToAction("Index");
            }

            ViewBag.PM_ID = new SelectList(db.Machines, "ID", "Machine1", goal.PM_ID);
            ViewBag.PositionID = new SelectList(db.Positions, "ID", "Position1", goal.PositionID);
            return View(goal);
        }

        // GET: Goals/Edit/5
        public async Task<ActionResult> Edit(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Goal goal = await db.Goals.FindAsync(id);
            if (goal == null)
            {
                return HttpNotFound();
            }
            ViewBag.PM_ID = new SelectList(db.Machines, "ID", "Machine1", goal.PM_ID);
            ViewBag.PositionID = new SelectList(db.Positions, "ID", "Position1", goal.PositionID);
            return View(goal);
        }

        // POST: Goals/Edit/5
        // To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        // more details see https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<ActionResult> Edit([Bind(Include = "ID,PositionID,PM_ID,Goal1")] Goal goal)
        {
            if (ModelState.IsValid)
            {
                db.Entry(goal).State = EntityState.Modified;
                await db.SaveChangesAsync();
                return RedirectToAction("Index");
            }
            ViewBag.PM_ID = new SelectList(db.Machines, "ID", "Machine1", goal.PM_ID);
            ViewBag.PositionID = new SelectList(db.Positions, "ID", "Position1", goal.PositionID);
            return View(goal);
        }

        // GET: Goals/Delete/5
        public async Task<ActionResult> Delete(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Goal goal = await db.Goals.FindAsync(id);
            if (goal == null)
            {
                return HttpNotFound();
            }
            return View(goal);
        }

        // POST: Goals/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public async Task<ActionResult> DeleteConfirmed(int id)
        {
            Goal goal = await db.Goals.FindAsync(id);
            db.Goals.Remove(goal);
            await db.SaveChangesAsync();
            return RedirectToAction("Index");
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                db.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}
