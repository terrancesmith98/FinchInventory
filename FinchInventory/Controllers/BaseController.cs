using FinchInventory.Models;
using System;
using System.DirectoryServices.AccountManagement;
using System.Linq;
using System.Web.Mvc;

namespace FinchInventory.Controllers
{
    public class BaseController : Controller
    {
        private static FinchDbContext db = new FinchDbContext();

        protected override void OnActionExecuting(ActionExecutingContext filterContext)
        {
            //string UserEmail = "terry.smith@finchpaper.com";


#if (!DEBUG)
            var UserDomain = Environment.UserDomainName.ToString();
            PrincipalContext ctx = new PrincipalContext(ContextType.Domain);
            var user = UserPrincipal.FindByIdentity(ctx, User.Identity.Name);
            var UserEmail = user.EmailAddress;
#endif

#if (DEBUG)
            //var UserEmail = "terry.smith@finchpaper.com";
            var UserEmail = "mike.hammond@finchpaper.com";
#endif

            db.Dispose();
            db = new FinchDbContext();

            if (UserEmail != null)
            {
                var currUser = db.Users.SingleOrDefault(u => u.UserName == UserEmail);
                if (currUser != null)
                {
                    var roles = db.UserRoles.Where(r => r.UserID == currUser.ID).Select(r => r.RoleID).ToList();
                    ViewBag.CurrUser = currUser;
                    ViewBag.UserRoles = roles;
                    ViewBag.RolesList = new MultiSelectList(db.Roles.OrderBy(r => r.ID).Select(r => r.Role1).ToList());

                }
                else
                {
                    ViewBag.ErrorMessage = $"There is no user account found that matches the current logged in user.";
                    ViewBag.ErrorDetails = $@"The current logged in user has the email address {UserEmail}.
                                            Please contact an admin for assistance. Current admins include:";
                }
            }
            else
            {
                ViewBag.ErrorMessage = $"There is no user account found that matches the current logged in user.";
                ViewBag.ErrorDetails = $"The current logged in user has the email address {UserEmail}.";
            }
            ViewBag.Admins = db.Admins.ToList();
        }
    }
}