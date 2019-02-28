using FinchInventory.Controllers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Finch_Inventory.Controllers
{
    public class HomeController : BaseController
    {
        public ActionResult Index()
        {
            if (ViewBag.CurrUser != null)
            {
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