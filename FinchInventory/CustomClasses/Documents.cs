using MigraDoc.DocumentObjectModel.Tables;
using System;
using System.Collections.Generic;
using System.Linq;
using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Shapes.Charts;
using MigraDoc.DocumentObjectModel.Shapes;
using FinchInventory.Models;
using System.IO;

namespace FinchInventory.CustomClasses
{
    public class Documents
    {
        internal static Document InventoryAudit(List<Clothing> clothings)
        {
            //create new Migradoc document
            Document document = new Document();
            document.Info.Title = "Clothing/Roll Inventory Audit Report";
            document.Info.Subject = "Displays an inventory per Paper Machine.";
            document.Info.Author = "Terry Smith Custom Applications";
            Styles.DefineStyles(document);
            DefineInventoryAuditContentSection(document);

            //add report heading
            document.LastSection.AddParagraph("Clothing/Roll Inventory as of  - " + DateTime.Now.ToShortDateString(), "Heading1");
            document.LastSection.AddParagraph("", "FooterText");

            //add main content tables

            //machine 1 clothing inventory
            document.LastSection.AddParagraph("", "FooterText");
            document.LastSection.AddParagraph("Clothing Inventory - Machine 1", "Heading3");
            document.LastSection.AddParagraph("", "FooterText");
            var clothings1 = clothings.Where(x => x.PM_Number == 1 && x.RollTypeID == 2).ToList();
            Table table1 = Tables.BuildInventoryAuditTable(clothings1);
            document.LastSection.Add(table1);
            document.LastSection.AddPageBreak();

            //machine 2 clothing inventory
            document.LastSection.AddParagraph("", "FooterText");
            document.LastSection.AddParagraph("Clothing Inventory - Machine 2", "Heading3");
            document.LastSection.AddParagraph("", "FooterText");
            var clothings2 = clothings.Where(x => x.PM_Number == 2 && x.RollTypeID == 2).ToList();
            Table table2 = Tables.BuildInventoryAuditTable(clothings2);
            document.LastSection.Add(table2);
            document.LastSection.AddPageBreak();


            //machine 3 clothing inventory
            document.LastSection.AddParagraph("", "FooterText");
            document.LastSection.AddParagraph("Clothing Inventory - Machine 3", "Heading3");
            document.LastSection.AddParagraph("", "FooterText");
            var clothings3 = clothings.Where(x => x.PM_Number == 3 && x.RollTypeID == 2).ToList();
            Table table3 = Tables.BuildInventoryAuditTable(clothings3);
            document.LastSection.Add(table3);
            document.LastSection.AddPageBreak();


            //machine 4 clothing inventory
            document.LastSection.AddParagraph("", "FooterText");
            document.LastSection.AddParagraph("Clothing Inventory - Machine 4", "Heading3");
            document.LastSection.AddParagraph("", "FooterText");
            var clothings4 = clothings.Where(x => x.PM_Number == 4 && x.RollTypeID == 2).ToList();
            Table table4 = Tables.BuildInventoryAuditTable(clothings4);
            document.LastSection.Add(table4);
            document.LastSection.AddPageBreak();

            //machine 1 roll inventory
            document.LastSection.AddParagraph("", "FooterText");
            document.LastSection.AddParagraph("Roll Inventory - Machine 1", "Heading3");
            document.LastSection.AddParagraph("", "FooterText");
            var rolls1 = clothings.Where(x => x.PM_Number == 1 && x.RollTypeID == 1).ToList();
            Table table5 = Tables.BuildInventoryAuditTable(rolls1);
            document.LastSection.Add(table5);
            document.LastSection.AddPageBreak();

            //machine 2 clothing inventory
            document.LastSection.AddParagraph("", "FooterText");
            document.LastSection.AddParagraph("Roll Inventory - Machine 2", "Heading3");
            document.LastSection.AddParagraph("", "FooterText");
            var rolls2 = clothings.Where(x => x.PM_Number == 2 && x.RollTypeID == 1).ToList();
            Table table6 = Tables.BuildInventoryAuditTable(rolls2);
            document.LastSection.Add(table6);
            document.LastSection.AddPageBreak();


            //machine 3 clothing inventory
            document.LastSection.AddParagraph("", "FooterText");
            document.LastSection.AddParagraph("Roll Inventory - Machine 3", "Heading3");
            document.LastSection.AddParagraph("", "FooterText");
            var rolls3 = clothings.Where(x => x.PM_Number == 3 && x.RollTypeID == 1).ToList();
            Table table7 = Tables.BuildInventoryAuditTable(rolls3);
            document.LastSection.Add(table7);
            document.LastSection.AddPageBreak();


            //machine 4 clothing inventory
            document.LastSection.AddParagraph("", "FooterText");
            document.LastSection.AddParagraph("Roll Inventory - Machine 4", "Heading3");
            document.LastSection.AddParagraph("", "FooterText");
            var rolls4 = clothings.Where(x => x.PM_Number == 4 && x.RollTypeID == 1).ToList();
            Table table8 = Tables.BuildInventoryAuditTable(rolls4);
            document.LastSection.Add(table8);

            return document;
        }

        private static void DefineInventoryAuditContentSection(Document document)
        {
            Section section = document.AddSection();
            section.PageSetup.HeaderDistance = 2;
            section.PageSetup.LeftMargin = 10;
            section.PageSetup.TopMargin = 30;
            section.PageSetup.RightMargin = 10;
            section.PageSetup.DifferentFirstPageHeaderFooter = true;


            //Create report footer
            HeaderFooter footer = section.Footers.FirstPage;
            Paragraph paragraph = new Paragraph();
            paragraph.AddFormattedText("Report Printed on: " + DateTime.Now.ToShortDateString(), "FooterText");
            paragraph.Format.Alignment = ParagraphAlignment.Center;
            footer.Add(paragraph);


        }

        internal static Document WeeklyPMReport(List<Clothing> clothings, List<Clothing> rolls)
        {
            //create new Migradoc document
            Document document = new Document();
            document.Info.Title = "Weekly PM Report";
            document.Info.Subject = "Displays a weekly paper maching item aging report.";
            document.Info.Author = "Terry Smith Custom Applications";
            Styles.DefineStyles(document);
            DefineWeeklyPMClothingContentSection(document);

            //add report heading
            document.LastSection.AddParagraph("Weekly Paper Machine Report  - " + DateTime.Now.ToShortDateString(), "Heading1");
            document.LastSection.AddParagraph("", "FooterText");

            //add main content table
            document.LastSection.AddParagraph("", "FooterText");
            document.LastSection.AddParagraph("Clothing", "Heading3");
            document.LastSection.AddParagraph("", "FooterText");
            Table tableClothing = Tables.BuildWeeklyPMClothingTable(clothings);
            document.LastSection.Add(tableClothing);
            document.LastSection.AddPageBreak();
            document.LastSection.AddParagraph("", "FooterText");
            document.LastSection.AddParagraph("Rolls", "Heading3");
            document.LastSection.AddParagraph("", "FooterText");
            Table tableRolls = Tables.BuildWeeklyPMRollsTable(rolls);
            document.LastSection.Add(tableRolls);
            return document;
        }

        private static void DefineWeeklyPMClothingContentSection(Document document)
        {
            Section section = document.AddSection();
            section.PageSetup.HeaderDistance = 2;
            section.PageSetup.LeftMargin = 10;
            section.PageSetup.TopMargin = 30;
            section.PageSetup.RightMargin = 10;
            section.PageSetup.DifferentFirstPageHeaderFooter = true;
            section.PageSetup.Orientation = Orientation.Landscape;


            //Create report footer
            HeaderFooter footer = section.Footers.FirstPage;
            Paragraph paragraph = new Paragraph();
            paragraph.AddFormattedText("Report Printed on: " + DateTime.Now.ToShortDateString(), "FooterText");
            paragraph.Format.Alignment = ParagraphAlignment.Center;
            footer.Add(paragraph);
        }
    }
}