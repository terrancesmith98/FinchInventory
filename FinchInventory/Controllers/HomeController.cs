using FinchInventory.Controllers;
using FinchInventory.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.Mvc;

namespace Finch_Inventory.Controllers
{
    public class HomeController : BaseController
    {
        private static FinchDbContext db = new FinchDbContext();
        public ActionResult Index()
        {
            db.Dispose();
            db = new FinchDbContext();
            if (ViewBag.CurrUser != null)
            {
                var pmNumber = Request["pmNumber"];
                var type = Request["type"];
                var inventory = new List<Clothing>();
                if (!string.IsNullOrEmpty(pmNumber))
                {
                    var pm = Convert.ToInt32(pmNumber);
                    if (!string.IsNullOrEmpty(type))
                    {
                        var rollType = 0;
                        switch (type)
                        {
                            case "rolls":
                                rollType = 1;
                                break;
                            case "clothing":
                                rollType = 2;
                                break;
                        }

                        inventory = db.Clothings.Where(c => c.PM_Number == pm).Where(c => c.RollTypeID == rollType).ToList();

                    }
                    else inventory = db.Clothings.Where(c => c.PM_Number == pm).ToList();
                }
                else if(!string.IsNullOrEmpty(type))
                {
                    var rollType = 0;
                    switch (type)
                    {
                        case "rolls":
                            rollType = 1;
                            break;
                        case "clothing":
                            rollType = 2;
                            break;
                    }
                    inventory = db.Clothings.Where(c => c.RollTypeID == rollType).ToList();
                }
                else inventory = db.Clothings.Where(c => c.PM_Number == 1 && c.RollTypeID == 2).ToList();

                ViewBag.Inventory = inventory;
                return View();

            }
            return View("Error");
        }



        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }
    }
}