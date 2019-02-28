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
    public class UsersController : BaseController
    {
        private FinchDbContext db = new FinchDbContext();

        // GET: Users
        public async Task<ActionResult> Index()
        {
            if (ViewBag.UserRoles.Contains(3) || ViewBag.UserRoles.Contains(4))
            {
                return View(await db.Users.ToListAsync());
            }
            return View("Error");
        }

        // GET: Users/Details/5
        public async Task<ActionResult> Details(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            User user = await db.Users.FindAsync(id);
            if (user == null)
            {
                return HttpNotFound();
            }
            return View(user);
        }

        // GET: Users/Create
        public ActionResult Create()
        {
            return View();
        }

        // POST: Users/Create
        // To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        // more details see https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<ActionResult> Create([Bind(Include = "ID,FirstName,LastName,UserName")] User user)
        {
            if (ModelState.IsValid)
            {
                db.Users.Add(user);
                await db.SaveChangesAsync();
                return RedirectToAction("Index");
            }

            return View(user);
        }

        // GET: Users/Edit/5
        public async Task<ActionResult> Edit(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            User user = await db.Users.FindAsync(id);
            if (user == null)
            {
                return HttpNotFound();
            }
            return View(user);
        }


        // Post: Users/Edit/5
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<ActionResult> Edit(int ID, string FirstName, string LastName, string UserName, int[] UserRoles)
        {
            if (ID > 0 && !string.IsNullOrEmpty(FirstName) && !string.IsNullOrEmpty(LastName) && !string.IsNullOrEmpty(UserName) && UserRoles.Length > 0)
            {
                var existing = db.Users.Find(ID);
                if (existing != null)
                {
                    try
                    {
                        existing.FirstName = FirstName;
                        existing.LastName = LastName;
                        existing.UserName = UserName;
                        foreach (var role in UserRoles)
                        {
                            var userRole = db.UserRoles.Where(r => r.UserID == ID && r.RoleID == role).Any();
                            if (!userRole)
                            {
                                var newRole = new UserRole(ID, role);
                                db.UserRoles.Add(newRole);
                            }
                        }
                    }
                    catch (Exception)
                    {

                        throw;
                    }
                }
                await db.SaveChangesAsync();
                return RedirectToAction("Index");
            }
            return View();
        }

        // GET: Users/Delete/5
        public async Task<ActionResult> Delete(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            User user = await db.Users.FindAsync(id);
            if (user == null)
            {
                return HttpNotFound();
            }
            return View(user);
        }

        // POST: Users/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public async Task<ActionResult> DeleteConfirmed(int id)
        {
            User user = await db.Users.FindAsync(id);
            db.Users.Remove(user);
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
