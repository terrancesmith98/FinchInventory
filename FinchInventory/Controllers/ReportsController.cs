using CsvHelper;
using FinchInventory.CustomClasses;
using FinchInventory.Models;
using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Tables;
using MigraDoc.Rendering;
using PdfSharp.Drawing;
using PdfSharp.Drawing.Layout;
using PdfSharp.Pdf;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Web.Mvc;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace Finch_Inventory.Controllers
{

    public class ReportsController : Controller
    {
        private static FinchDbContext db = new FinchDbContext();
        // GET: Reports
        public ActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public string CreateInventoryAuditReport()
        {
            var inventoryItems = db.Clothings.Where(c => c.StatusID != 2 && c.StatusID != 3).ToList();
            //create Migradoc Document
            Document document = Documents.InventoryAudit(inventoryItems);
            //create PDF Renderer
            PdfDocumentRenderer renderer = new PdfDocumentRenderer(true);
            renderer.Document = document;
            renderer.RenderDocument();
            var fileName = "Finch_Inventory_Audit.pdf";
            var filePath = Path.Combine(Server.MapPath("~/Content/Reports"), fileName);
            renderer.PdfDocument.Save(filePath);

            return fileName;
        }

        [HttpPost]
        public string CreateWeeklyPMReport()
        {
            try
            {
                var currentClothing = db.Clothings.Where(c => c.StatusID == 2 && c.RollTypeID == 2).ToList();
                var currentRolls = db.Clothings.Where(c => c.StatusID == 2 && c.RollTypeID == 1).ToList();
                //create Migradoc Document
                Document document = Documents.WeeklyPMReport(currentClothing, currentRolls);
                PageSetup pageSetup = document.DefaultPageSetup.Clone();
                // set orientation

                pageSetup.Orientation = MigraDoc.DocumentObjectModel.Orientation.Landscape;
                //create PDF renderer
                PdfDocumentRenderer renderer = new PdfDocumentRenderer(true);
                renderer.Document = document;
                renderer.RenderDocument();
                var fileName = "Finch_Weekly_PM_Report.pdf";
                var filePath = Path.Combine(Server.MapPath("~/Content/Reports"), fileName);
                renderer.PdfDocument.Save(filePath);

                return fileName;
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
            
        }

        [HttpGet]
        public FileStreamResult InventoryToExcel()
        {
            var clothings = WriteCsvToMemory(db.Clothings.ToList());
            var memoryStream = new MemoryStream(clothings);
            var today = (DateTime.Now.ToShortDateString().Replace("/", "-"));

            var fileName = $"clothing-roll-inventory_{today}.csv";

            return new FileStreamResult(memoryStream, "text/csv") { FileDownloadName = fileName };
        }

        public FileStreamResult HistoricalToExcel()
        {
            var clothings = WriteCsvToMemory(db.Clothings.Where(c => c.StatusID == 3).ToList());
            var memoryStream = new MemoryStream(clothings);
            var today = (DateTime.Now.ToShortDateString().Replace("/", "-"));

            var fileName = $"clothing-roll-historical_{today}.csv";

            return new FileStreamResult(memoryStream, "text/csv") { FileDownloadName = fileName };
        }

        public byte[] WriteCsvToMemory(List<Clothing> clothing)
        {
            using (var memoryStream = new MemoryStream())
            using (var streamWriter = new StreamWriter(memoryStream))
            using (var csvWriter = new CsvWriter(streamWriter))
            {
                csvWriter.WriteRecords(clothing);
                streamWriter.Flush();
                return memoryStream.ToArray();
            }
        }
    }
}