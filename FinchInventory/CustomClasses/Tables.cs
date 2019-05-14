using FinchInventory.Models;
using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Tables;
using System;
using System.Collections.Generic;
using System.Linq;

namespace FinchInventory.CustomClasses
{
    public class Tables
    {
        static FinchDbContext db = new FinchDbContext();

        internal static Table BuildInventoryAuditTable(List<Clothing> clothings)
        {
            Table auditTable = new Table();
            auditTable.Style = "Table";
            auditTable.Borders.Width = .075;
            auditTable.Rows.Alignment = RowAlignment.Center;

            //define header columns
            Column posCol = auditTable.AddColumn(Unit.FromCentimeter(6));
            posCol.Format.Alignment = ParagraphAlignment.Left;
            posCol.Format.Font.Size = 10;

            Column typeCol = auditTable.AddColumn(Unit.FromCentimeter(7));
            typeCol.Format.Alignment = ParagraphAlignment.Left;
            typeCol.Format.Font.Size = 9.5;

            Column locCol = auditTable.AddColumn(Unit.FromCentimeter(3));
            locCol.Format.Alignment = ParagraphAlignment.Center;
            locCol.Format.Font.Size = 10;

            Column recCol = auditTable.AddColumn(Unit.FromCentimeter(3));
            recCol.Format.Alignment = ParagraphAlignment.Center;
            recCol.Format.Font.Size = 10;

            //define headers row
            Row headers = auditTable.AddRow();
            headers.Shading.Color = Colors.LightGray;
            headers.Height = 10;

            //define header row cells
            Cell posHeader = headers.Cells[0];
            posHeader.Format.Font.Size = 10;
            posHeader.VerticalAlignment = VerticalAlignment.Center;
            posHeader.AddParagraph("Position");

            Cell typeHeader = headers.Cells[1];
            typeHeader.Format.Font.Size = 10;
            typeHeader.VerticalAlignment = VerticalAlignment.Center;
            typeHeader.AddParagraph("Type");

            Cell locHeader = headers.Cells[2];
            locHeader.Format.Font.Size = 10;
            locHeader.VerticalAlignment = VerticalAlignment.Center;
            locHeader.AddParagraph("Location");

            Cell recHeader = headers.Cells[3];
            recHeader.Format.Font.Size = 10;
            recHeader.VerticalAlignment = VerticalAlignment.Center;
            recHeader.AddParagraph("Date Received");

            //fill content cells
            var rowCounter = 1;
            var currentPosID = 0;
            foreach (var clothing in clothings.OrderBy(x => x.PositionID).ThenBy(x => x.Date_Received))
            {
                if (clothing.PositionID != currentPosID && rowCounter != 1)
                {
                    Row rBlank = auditTable.AddRow();
                    Row rBlank2 = auditTable.AddRow();
                }
                currentPosID = clothing.PositionID;
                Row r = auditTable.AddRow();
                r.Height = 13;
                if (rowCounter % 2 == 0)
                {
                    r.Format.Shading.Color = Colors.LightGray;
                }


                Cell cPos = r.Cells[0];
                cPos.Format.Font.Size = 8.5;
                cPos.AddParagraph($"{clothing.Position.Position1} - {clothing.Dimensions}");
                cPos.VerticalAlignment = VerticalAlignment.Center;

                Cell cType = r.Cells[1];
                cType.Format.Font.Size = 8.5;
                cType.AddParagraph($"{clothing.Manufacturer.Name} - #{clothing.Serial_Number} - {clothing.Type.Type1}");
                cType.VerticalAlignment = VerticalAlignment.Center;

                Cell cLoc = r.Cells[2];
                cLoc.Format.Font.Size = 8.5;
                cLoc.AddParagraph(clothing.Location.Location1);
                cLoc.VerticalAlignment = VerticalAlignment.Center;

                Cell cRec = r.Cells[3];
                cRec.Format.Font.Size = 8.5;
                var dateReceived = new DateTime();
                if (clothing.Date_Received != null)
                {
                    dateReceived = (DateTime)clothing.Date_Received;
                    cRec.AddParagraph($"{dateReceived.ToShortDateString()}");
                    cRec.VerticalAlignment = VerticalAlignment.Center;
                }
                   

                rowCounter++;
            }

            return auditTable;
        }

        internal static Table BuildWeeklyPMClothingTable(List<Clothing> clothings)
        {
            var weeklyPMTable = new Table
            {
                Style = "Table"
            };
            weeklyPMTable.Borders.Width = .075;
            weeklyPMTable.Rows.Alignment = RowAlignment.Left;
            weeklyPMTable.Rows.LeftIndent = 20;

            #region Header Columns
            //define header columns
            var posCol = weeklyPMTable.AddColumn(Unit.FromCentimeter(3.5));
            posCol.Format.Alignment = ParagraphAlignment.Center;
            posCol.Format.Font.Size = 10;

            var pm1Col = weeklyPMTable.AddColumn(Unit.FromCentimeter(2));
            pm1Col.Format.Alignment = ParagraphAlignment.Center;
            pm1Col.Format.Font.Size = 10;

            var pm1Colb = weeklyPMTable.AddColumn(Unit.FromCentimeter(2));
            pm1Colb.Format.Alignment = ParagraphAlignment.Center;
            pm1Colb.Format.Font.Size = 10;

            var pm1Colc = weeklyPMTable.AddColumn(Unit.FromCentimeter(2));
            pm1Colc.Format.Alignment = ParagraphAlignment.Center;
            pm1Colc.Format.Font.Size = 10;

            var pm2col = weeklyPMTable.AddColumn(Unit.FromCentimeter(2));
            pm2col.Format.Alignment = ParagraphAlignment.Center;
            pm2col.Format.Font.Size = 10;

            var pm2colb = weeklyPMTable.AddColumn(Unit.FromCentimeter(2));
            pm2colb.Format.Alignment = ParagraphAlignment.Center;
            pm2colb.Format.Font.Size = 10;

            var pm2colc = weeklyPMTable.AddColumn(Unit.FromCentimeter(2));
            pm2colc.Format.Alignment = ParagraphAlignment.Center;
            pm2colc.Format.Font.Size = 10;

            var pm3col = weeklyPMTable.AddColumn(Unit.FromCentimeter(2));
            pm3col.Format.Alignment = ParagraphAlignment.Center;
            pm3col.Format.Font.Size = 10;

            var pm3colb = weeklyPMTable.AddColumn(Unit.FromCentimeter(2));
            pm3colb.Format.Alignment = ParagraphAlignment.Center;
            pm3colb.Format.Font.Size = 10;

            var pm3colc = weeklyPMTable.AddColumn(Unit.FromCentimeter(2));
            pm3colc.Format.Alignment = ParagraphAlignment.Center;
            pm3colc.Format.Font.Size = 10;

            var pm4col = weeklyPMTable.AddColumn(Unit.FromCentimeter(2));
            pm4col.Format.Alignment = ParagraphAlignment.Center;
            pm4col.Format.Font.Size = 10;

            var pm4colb = weeklyPMTable.AddColumn(Unit.FromCentimeter(2));
            pm4colb.Format.Alignment = ParagraphAlignment.Center;
            pm4colb.Format.Font.Size = 10;

            var pm4colc = weeklyPMTable.AddColumn(Unit.FromCentimeter(2));
            pm4colc.Format.Alignment = ParagraphAlignment.Center;
            pm4colc.Format.Font.Size = 10;

            #endregion

            #region Header Rows

            //define headers row
            var headers = weeklyPMTable.AddRow();
            headers.Shading.Color = Colors.LightGray;
            headers.Height = 15;

            //define header row cells
            var posHeader = headers.Cells[0];
            posHeader.Format.Font.Size = 10;
            posHeader.VerticalAlignment = VerticalAlignment.Center;
            posHeader.MergeDown = 1;
            posHeader.AddParagraph("Position");

            var pm1Header = headers.Cells[1];
            pm1Header.Format.Font.Size = 10;
            pm1Header.MergeRight = 2;
            pm1Header.VerticalAlignment = VerticalAlignment.Center;
            pm1Header.AddParagraph("#1 PM Days Run");

            var pm2Header = headers.Cells[4];
            pm2Header.Format.Font.Size = 10;
            pm2Header.MergeRight = 2;
            pm2Header.VerticalAlignment = VerticalAlignment.Center;
            pm2Header.AddParagraph("#2 PM Days Run");

            var pm3Header = headers.Cells[7];
            pm3Header.Format.Font.Size = 10;
            pm3Header.MergeRight = 2;
            pm3Header.VerticalAlignment = VerticalAlignment.Center;
            pm3Header.AddParagraph("#3 PM Days Run");

            var pm4Header = headers.Cells[10];
            pm4Header.Format.Font.Size = 10;
            pm4Header.MergeRight = 2;
            pm4Header.VerticalAlignment = VerticalAlignment.Center;
            pm4Header.AddParagraph("#4 PM Days Run");

            //define header row 2
            var headers2 = weeklyPMTable.AddRow();
            headers2.Height = 10;

            //define 2nd header row cells
            var pm1PastHeader = headers2.Cells[1];
            pm1PastHeader.Format.Font.Size = 8;
            pm1PastHeader.VerticalAlignment = VerticalAlignment.Center;
            pm1PastHeader.AddParagraph("PAST");

            var pm1CurrentHeader = headers2.Cells[2];
            pm1CurrentHeader.Format.Font.Size = 8;
            pm1CurrentHeader.VerticalAlignment = VerticalAlignment.Center;
            pm1CurrentHeader.AddParagraph("CURRENT");

            var pm1GoalHeader = headers2.Cells[3];
            pm1GoalHeader.Format.Font.Size = 8;
            pm1GoalHeader.VerticalAlignment = VerticalAlignment.Center;
            pm1GoalHeader.AddParagraph("GOAL");

            var pm2PastHeader = headers2.Cells[4];
            pm2PastHeader.Format.Font.Size = 8;
            pm2PastHeader.VerticalAlignment = VerticalAlignment.Center;
            pm2PastHeader.AddParagraph("PAST");

            var pm2CurrentHeader = headers2.Cells[5];
            pm2CurrentHeader.Format.Font.Size = 8;
            pm2CurrentHeader.VerticalAlignment = VerticalAlignment.Center;
            pm2CurrentHeader.AddParagraph("CURRENT");

            var pm2GoalHeader = headers2.Cells[6];
            pm2GoalHeader.Format.Font.Size = 8;
            pm2GoalHeader.VerticalAlignment = VerticalAlignment.Center;
            pm2GoalHeader.AddParagraph("GOAL");

            var pm3PastHeader = headers2.Cells[7];
            pm3PastHeader.Format.Font.Size = 8;
            pm3PastHeader.VerticalAlignment = VerticalAlignment.Center;
            pm3PastHeader.AddParagraph("PAST");

            var pm3CurrentHeader = headers2.Cells[8];
            pm3CurrentHeader.Format.Font.Size = 8;
            pm3CurrentHeader.VerticalAlignment = VerticalAlignment.Center;
            pm3CurrentHeader.AddParagraph("CURRENT");

            var pm3GoalHeader = headers2.Cells[9];
            pm3GoalHeader.Format.Font.Size = 8;
            pm3GoalHeader.VerticalAlignment = VerticalAlignment.Center;
            pm3GoalHeader.AddParagraph("GOAL");

            var pm4PastHeader = headers2.Cells[10];
            pm4PastHeader.Format.Font.Size = 8;
            pm4PastHeader.VerticalAlignment = VerticalAlignment.Center;
            pm4PastHeader.AddParagraph("PAST");

            var pm4CurrentHeader = headers2.Cells[11];
            pm4CurrentHeader.Format.Font.Size = 8;
            pm4CurrentHeader.VerticalAlignment = VerticalAlignment.Center;
            pm4CurrentHeader.AddParagraph("CURRENT");

            var pm4GoalHeader = headers2.Cells[12];
            pm4GoalHeader.Format.Font.Size = 8;
            pm4GoalHeader.VerticalAlignment = VerticalAlignment.Center;
            pm4GoalHeader.AddParagraph("GOAL");
            #endregion

            //define Content Rows

            #region Wire Position
            // ----- Wire Position Row ----- //
            var clothingWire1 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "wire" && c.PM_Number == 1);
            var clothingWire2 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "wire" && c.PM_Number == 2);
            var clothingWire3 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "wire" && c.PM_Number == 3);
            var clothingWire4 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "wire" && c.PM_Number == 4);
            var rowWire = weeklyPMTable.AddRow();
            rowWire.Height = 10;
            var rowWire2 = weeklyPMTable.AddRow();
            rowWire2.Height = 10;

            var cellWire = rowWire.Cells[0];
            cellWire.Format.Font.Size = 8;
            cellWire.VerticalAlignment = VerticalAlignment.Center;
            cellWire.MergeDown = 1;
            cellWire.AddParagraph("WIRE");

            #region Wire PM 1
            // ----- Wire Position PM 1 ----- //
            var cellWire1Past = rowWire.Cells[1];
            cellWire1Past.Format.Font.Size = 7.5;
            cellWire1Past.VerticalAlignment = VerticalAlignment.Center;
            var pastAverage = GetPastAverage(1, 2);
            if (pastAverage > 0)
                cellWire1Past.AddParagraph($"{pastAverage}");

            var cellWire1Current = rowWire.Cells[2];
            cellWire1Current.Format.Font.Size = 7.5;
            cellWire1Current.VerticalAlignment = VerticalAlignment.Center;

            var currentAge = GetCurrentAge(clothingWire1);
            var positionGoal = GetPositionGoal(1, 2);
            if (currentAge > 0)
                cellWire1Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cellWire1Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cellWire1Current.Shading.Color = Colors.Yellow;
            }
            

            var cellWire1Goal = rowWire.Cells[3];
            cellWire1Goal.Format.Font.Size = 7.5;
            cellWire1Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cellWire1Goal.AddParagraph($"{positionGoal}");

            var cellWire1Info = rowWire2.Cells[1];
            cellWire1Info.Format.Font.Size = 7.5;
            cellWire1Info.MergeRight = 2;
            cellWire1Info.Shading.Color = Colors.LightBlue;
            cellWire1Info.AddParagraph($"{GetClothingInfo(clothingWire1)}");
            #endregion

            #region Wire PM 2

            // ----- Wire Position PM 2 ----- //
            var cellWire2Past = rowWire.Cells[4];
            cellWire2Past.Format.Font.Size = 7.5;
            cellWire2Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(2, 2);
            if (pastAverage > 0)
                cellWire2Past.AddParagraph($"{pastAverage}");

            var cellWire2Current = rowWire.Cells[5];
            cellWire2Current.Format.Font.Size = 7.5;
            cellWire2Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothingWire2);
            positionGoal = GetPositionGoal(2, 2);
            if (currentAge > 0)
                cellWire2Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cellWire2Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cellWire2Current.Shading.Color = Colors.Yellow;
            }
            

            var cellWire2Goal = rowWire.Cells[6];
            cellWire2Goal.Format.Font.Size = 7.5;
            cellWire2Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cellWire2Goal.AddParagraph($"{positionGoal}");

            var cellWire2Info = rowWire2.Cells[4];
            cellWire2Info.Format.Font.Size = 7.5;
            cellWire2Info.MergeRight = 2;
            cellWire2Info.Shading.Color = Colors.LightBlue;
            cellWire2Info.AddParagraph($"{GetClothingInfo(clothingWire2)}");
            #endregion

            #region Wire PM 3
            // ----- Wire Position PM 3 ----- //
            var cellWire3Past = rowWire.Cells[7];
            cellWire3Past.Format.Font.Size = 7.5;
            cellWire3Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(3, 2);
            if (pastAverage > 0)
                cellWire3Past.AddParagraph($"{pastAverage}");

            var cellWire3Current = rowWire.Cells[8];
            cellWire3Current.Format.Font.Size = 7.5;
            cellWire3Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothingWire3);
            positionGoal = GetPositionGoal(3, 2);
            if (currentAge > 0)
                cellWire3Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cellWire3Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cellWire3Current.Shading.Color = Colors.Yellow;
            }
            

            var cellWire3Goal = rowWire.Cells[9];
            cellWire3Goal.Format.Font.Size = 7.5;
            cellWire3Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cellWire3Goal.AddParagraph($"{positionGoal}");

            var cellWire3Info = rowWire2.Cells[7];
            cellWire3Info.Format.Font.Size = 7.5;
            cellWire3Info.MergeRight = 2;
            cellWire3Info.Shading.Color = Colors.LightBlue;
            cellWire3Info.AddParagraph($"{GetClothingInfo(clothingWire3)}");
            #endregion

            #region Wire PM 4
            // ----- Wire Position PM 4 ----- //
            var cellWire4Past = rowWire.Cells[10];
            cellWire4Past.Format.Font.Size = 7.5;
            cellWire4Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(4, 2);
            if (pastAverage > 0)
                cellWire4Past.AddParagraph($"{pastAverage}");

            var cellWire4Current = rowWire.Cells[11];
            cellWire4Current.Format.Font.Size = 7.5;
            cellWire4Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothingWire4);
            positionGoal = GetPositionGoal(4, 2);
            if (currentAge > 0)
                cellWire4Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cellWire4Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cellWire4Current.Shading.Color = Colors.Yellow;
            }
            

            var cellWire4Goal = rowWire.Cells[12];
            cellWire4Goal.Format.Font.Size = 7.5;
            cellWire4Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cellWire4Goal.AddParagraph($"{positionGoal}");

            var cellWire4Info = rowWire2.Cells[10];
            cellWire4Info.Format.Font.Size = 7.5;
            cellWire4Info.MergeRight = 2;
            cellWire4Info.Shading.Color = Colors.LightBlue;
            cellWire4Info.AddParagraph($"{GetClothingInfo(clothingWire4)}");
            #endregion

            #endregion // Wire Position

            #region 1st Press Position

            // ----- 1st Press Position Row ----- //
            // add spacer row
            var spacer = weeklyPMTable.AddRow();
            spacer.Height = 12;
            var cellSpacer = spacer.Cells[0];
            cellSpacer.MergeRight = 12;
            cellSpacer.AddParagraph("");

            var clothingFirstPress1 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "1st press felt" && c.PM_Number == 1);
            var clothingFirstPress2 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "1st press felt" && c.PM_Number == 2);
            var clothingFirstPress3 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "1st press felt" && c.PM_Number == 3);
            var clothingFirstPress4 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "1st press felt" && c.PM_Number == 4);
            var rowFirstPress = weeklyPMTable.AddRow();
            rowFirstPress.Height = 10;
            var rowFirstPress2 = weeklyPMTable.AddRow();
            rowFirstPress2.Height = 10;

            var cell1Press = rowFirstPress.Cells[0];
            cell1Press.Format.Font.Size = 8;
            cell1Press.VerticalAlignment = VerticalAlignment.Center;
            cell1Press.MergeDown = 1;
            cell1Press.AddParagraph("1ST PRESS");

            #region 1st Press PM 1
            // ----- 1st Press Position PM 1 ----- //
            var cell1Press1Past = rowFirstPress.Cells[1];
            cell1Press1Past.Format.Font.Size = 8;
            cell1Press1Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(1, 3);
            if (pastAverage > 0)
                cell1Press1Past.AddParagraph($"{pastAverage}");

            var cell1Press1Current = rowFirstPress.Cells[2];
            cell1Press1Current.Format.Font.Size = 8;
            cell1Press1Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothingFirstPress1);
            positionGoal = GetPositionGoal(1, 3);
            if (currentAge > 0)
                cell1Press1Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cell1Press1Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cell1Press1Current.Shading.Color = Colors.Yellow;
            }
            

            var cell1Press1Goal = rowFirstPress.Cells[3];
            cell1Press1Goal.Format.Font.Size = 8;
            cell1Press1Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cell1Press1Goal.AddParagraph($"{positionGoal}");

            var cell1Press1Info = rowFirstPress2.Cells[1];
            cell1Press1Info.Format.Font.Size = 8;
            cell1Press1Info.MergeRight = 2;
            cell1Press1Info.Shading.Color = Colors.LightBlue;
            cell1Press1Info.AddParagraph($"{GetClothingInfo(clothingFirstPress1)}");
            #endregion

            #region 1st Press PM 2
            // ----- 1st Press Position PM 2 ----- //
            var cell1Press2Past = rowFirstPress.Cells[4];
            cell1Press2Past.Format.Font.Size = 8;
            cell1Press2Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(2, 3);
            if (pastAverage > 0)
                cell1Press2Past.AddParagraph($"{GetPastAverage(2, 3)}");

            var cell1Press2Current = rowFirstPress.Cells[5];
            cell1Press2Current.Format.Font.Size = 8;
            cell1Press2Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothingFirstPress2);
            positionGoal = GetPositionGoal(2, 3);
            if (currentAge > 0)
                cell1Press2Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cell1Press2Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cell1Press2Current.Shading.Color = Colors.Yellow;
            }
            

            var cell1Press2Goal = rowFirstPress.Cells[6];
            cell1Press2Goal.Format.Font.Size = 8;
            cell1Press2Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cell1Press2Goal.AddParagraph($"{positionGoal}");

            var cell1Press2Info = rowFirstPress2.Cells[4];
            cell1Press2Info.Format.Font.Size = 8;
            cell1Press2Info.MergeRight = 2;
            cell1Press2Info.Shading.Color = Colors.LightBlue;
            cell1Press2Info.AddParagraph($"{GetClothingInfo(clothingFirstPress2)}");
            #endregion

            #region 1st Press PM 3
            // ----- 1st Press Position PM 3 ----- //
            var cell1Press3Past = rowFirstPress.Cells[7];
            cell1Press3Past.Format.Font.Size = 8;
            cell1Press3Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(3, 3);
            if (pastAverage > 0)
                cell1Press3Past.AddParagraph($"{pastAverage}");

            var cell1Press3Current = rowFirstPress.Cells[8];
            cell1Press3Current.Format.Font.Size = 8;
            cell1Press3Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothingFirstPress3);
            positionGoal = GetPositionGoal(3, 3);
            if (currentAge > 0)
                cell1Press3Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cell1Press3Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cell1Press3Current.Shading.Color = Colors.Yellow;
            }
            

            var cell1Press3Goal = rowFirstPress.Cells[9];
            cell1Press3Goal.Format.Font.Size = 8;
            cell1Press3Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cell1Press3Goal.AddParagraph($"{positionGoal}");

            var cell1Press3Info = rowFirstPress2.Cells[7];
            cell1Press3Info.Format.Font.Size = 8;
            cell1Press3Info.MergeRight = 2;
            cell1Press3Info.Shading.Color = Colors.LightBlue;
            cell1Press3Info.AddParagraph($"{GetClothingInfo(clothingFirstPress3)}");
            #endregion

            #region 1st Press PM 4
            // ----- 1st Press Position PM 4 ----- //
            var cell1Press4Past = rowFirstPress.Cells[10];
            cell1Press4Past.Format.Font.Size = 8;
            cell1Press4Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(4, 3);
            if (pastAverage > 0)
                cell1Press4Past.AddParagraph($"{pastAverage}");

            var cell1Press4Current = rowFirstPress.Cells[11];
            cell1Press4Current.Format.Font.Size = 8;
            cell1Press4Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothingFirstPress4);
            positionGoal = GetPositionGoal(4, 3);
            if (currentAge > 0)
                cell1Press4Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cell1Press4Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cell1Press4Current.Shading.Color = Colors.Yellow;
            }
           

            var cell1Press4Goal = rowFirstPress.Cells[12];
            cell1Press4Goal.Format.Font.Size = 8;
            cell1Press4Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cell1Press4Goal.AddParagraph($"{positionGoal}");

            var cell1Press4Info = rowFirstPress2.Cells[10];
            cell1Press4Info.Format.Font.Size = 8;
            cell1Press4Info.MergeRight = 2;
            cell1Press4Info.Shading.Color = Colors.LightBlue;
            cell1Press4Info.AddParagraph($"{GetClothingInfo(clothingFirstPress4)}");
            #endregion

            #endregion //1st Press Position

            #region 2nd Press Position
            // add spacer row
            spacer = weeklyPMTable.AddRow();
            spacer.Height = 12;
            cellSpacer = spacer.Cells[0];
            cellSpacer.MergeRight = 12;
            cellSpacer.AddParagraph("");

            var clothingSecondPress1 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "2nd press felt" && c.PM_Number == 1);
            var clothingSecondPress2 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "2nd press felt" && c.PM_Number == 2);
            var clothingSecondPress3 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "2nd press felt" && c.PM_Number == 3);
            var clothingSecondPress4 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "2nd press felt" && c.PM_Number == 4);
            var rowSecondPress = weeklyPMTable.AddRow();
            rowSecondPress.Height = 10;
            var rowSecondPress2 = weeklyPMTable.AddRow();
            rowSecondPress2.Height = 10;

            var cell2Press = rowSecondPress.Cells[0];
            cell2Press.Format.Font.Size = 8;
            cell2Press.VerticalAlignment = VerticalAlignment.Center;
            cell2Press.MergeDown = 1;
            cell2Press.AddParagraph("2ND PRESS");

            #region 2nd Press PM 1
            // ----- 2nd Press Position PM 1 ----- //
            var cell2Press1Past = rowSecondPress.Cells[1];
            cell2Press1Past.Format.Font.Size = 8;
            cell2Press1Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(1, 4);
            if (pastAverage > 0)
                cell2Press1Past.AddParagraph($"{pastAverage}");

            var cell2Press1Current = rowSecondPress.Cells[2];
            cell2Press1Current.Format.Font.Size = 8;
            cell2Press1Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothingSecondPress1);
            positionGoal = GetPositionGoal(1, 4);
            if (currentAge > 0)
                cell2Press1Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cell2Press1Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cell2Press1Current.Shading.Color = Colors.Yellow;
            }
            

            var cell2Press1Goal = rowSecondPress.Cells[3];
            cell2Press1Goal.Format.Font.Size = 8;
            cell2Press1Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cell2Press1Goal.AddParagraph($"{positionGoal}");

            var cell2Press1Info = rowSecondPress2.Cells[1];
            cell2Press1Info.Format.Font.Size = 8;
            cell2Press1Info.MergeRight = 2;
            cell2Press1Info.Shading.Color = Colors.LightBlue;
            cell2Press1Info.AddParagraph($"{GetClothingInfo(clothingSecondPress1)}");
            #endregion

            #region 2nd Press PM 2
            // ----- 2nd Press Position PM 2 ----- //
            var cell2Press2Past = rowSecondPress.Cells[4];
            cell2Press2Past.Format.Font.Size = 8;
            cell2Press2Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(2, 4);
            if (pastAverage > 0)
                cell2Press2Past.AddParagraph($"{pastAverage}");

            var cell2Press2Current = rowSecondPress.Cells[5];
            cell2Press2Current.Format.Font.Size = 8;
            cell2Press2Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothingSecondPress2);
            positionGoal = GetPositionGoal(2, 4);
            if (currentAge > 0)
                cell2Press2Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cell2Press2Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cell2Press2Current.Shading.Color = Colors.Yellow;
            }
            

            var cell2Press2Goal = rowSecondPress.Cells[6];
            cell2Press2Goal.Format.Font.Size = 8;
            cell2Press2Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cell2Press2Goal.AddParagraph($"{positionGoal}");

            var cell2Press2Info = rowSecondPress2.Cells[4];
            cell2Press2Info.Format.Font.Size = 8;
            cell2Press2Info.MergeRight = 2;
            cell2Press2Info.Shading.Color = Colors.LightBlue;
            cell2Press2Info.AddParagraph($"{GetClothingInfo(clothingSecondPress2)}");
            #endregion

            #region 2nd Press PM 3
            // ----- 2nd Press Position PM 3 ----- //
            var cell2Press3Past = rowSecondPress.Cells[7];
            cell2Press3Past.Format.Font.Size = 8;
            cell2Press3Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(3, 4);
            if (pastAverage > 0)
                cell2Press3Past.AddParagraph($"{pastAverage}");

            var cell2Press3Current = rowSecondPress.Cells[8];
            cell2Press3Current.Format.Font.Size = 8;
            cell2Press3Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothingSecondPress3);
            positionGoal = GetPositionGoal(3, 4);
            if (currentAge > 0)
                cell2Press3Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cell2Press3Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cell2Press3Current.Shading.Color = Colors.Yellow;
            }
            

            var cell2Press3Goal = rowSecondPress.Cells[9];
            cell2Press3Goal.Format.Font.Size = 8;
            cell2Press3Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cell2Press3Goal.AddParagraph($"{positionGoal}");

            var cell2Press3Info = rowSecondPress2.Cells[7];
            cell2Press3Info.Format.Font.Size = 8;
            cell2Press3Info.MergeRight = 2;
            cell2Press3Info.Shading.Color = Colors.LightBlue;
            cell2Press3Info.AddParagraph($"{GetClothingInfo(clothingSecondPress3)}");
            #endregion

            #region 2nd Press PM 4
            // ----- 2nd Press Position PM 4 ----- //
            var cell2Press4Past = rowSecondPress.Cells[10];
            cell2Press4Past.Format.Font.Size = 8;
            cell2Press4Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(4, 4);
            if (pastAverage > 0)
                cell2Press4Past.AddParagraph($"{pastAverage}");

            var cell2Press4Current = rowSecondPress.Cells[11];
            cell2Press4Current.Format.Font.Size = 8;
            cell2Press4Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothingSecondPress4);
            positionGoal = GetPositionGoal(4, 4);
            if (currentAge > 0)
                cell2Press4Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cell2Press4Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cell2Press4Current.Shading.Color = Colors.Yellow;
            }
            

            var cell2Press4Goal = rowSecondPress.Cells[12];
            cell2Press4Goal.Format.Font.Size = 8;
            cell2Press4Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cell2Press4Goal.AddParagraph($"{positionGoal}");

            var cell2Press4Info = rowSecondPress2.Cells[10];
            cell2Press4Info.Format.Font.Size = 8;
            cell2Press4Info.MergeRight = 2;
            cell2Press4Info.Shading.Color = Colors.LightBlue;
            cell2Press4Info.AddParagraph($"{GetClothingInfo(clothingSecondPress4)}");
            #endregion

            #endregion // 2nd Press Position

            #region 3rd Press Position

            // add spacer row
            spacer = weeklyPMTable.AddRow();
            spacer.Height = 12;
            cellSpacer = spacer.Cells[0];
            cellSpacer.MergeRight = 12;
            cellSpacer.AddParagraph("");

            var clothingThirdPress1 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "3rd press felt" && c.PM_Number == 1);
            var clothingThirdPress2 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "3rd press felt" && c.PM_Number == 2);
            var clothingThirdPress3 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "3rd press felt" && c.PM_Number == 3);
            var clothingThirdPress4 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "3rd press felt" && c.PM_Number == 4);
            var rowThirdPress = weeklyPMTable.AddRow();
            rowThirdPress.Height = 10;
            var rowThirdPress2 = weeklyPMTable.AddRow();
            rowThirdPress2.Height = 10;

            var cell3Press = rowThirdPress.Cells[0];
            cell3Press.Format.Font.Size = 8;
            cell3Press.VerticalAlignment = VerticalAlignment.Center;
            cell3Press.MergeDown = 1;
            cell3Press.AddParagraph("3rd PRESS");

            #region 3rd Press PM 1
            // ----- 3rd Press Position PM 1 ----- //
            var cell3Press1Past = rowThirdPress.Cells[1];
            cell3Press1Past.Format.Font.Size = 8;
            cell3Press1Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(1, 5);
            if (pastAverage > 0)
                cell3Press1Past.AddParagraph($"{pastAverage}");

            var cell3Press1Current = rowThirdPress.Cells[2];
            cell3Press1Current.Format.Font.Size = 8;
            cell3Press1Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothingThirdPress1);
            positionGoal = GetPositionGoal(1, 5);
            if (currentAge > 0)
                cell3Press1Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cell3Press1Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cell3Press1Current.Shading.Color = Colors.Yellow;
            }
            

            var cell3Press1Goal = rowThirdPress.Cells[3];
            cell3Press1Goal.Format.Font.Size = 8;
            cell3Press1Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cell3Press1Goal.AddParagraph($"{positionGoal}");

            var cell3Press1Info = rowThirdPress2.Cells[1];
            cell3Press1Info.Format.Font.Size = 8;
            cell3Press1Info.MergeRight = 2;
            cell3Press1Info.Shading.Color = Colors.LightBlue;
            cell3Press1Info.AddParagraph($"{GetClothingInfo(clothingThirdPress1)}");
            #endregion

            #region 3rd Press PM 2
            // ----- 3rd Press Position PM 2 ----- //
            var cell3Press2Past = rowThirdPress.Cells[4];
            cell3Press2Past.Format.Font.Size = 8;
            cell3Press2Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(2, 5);
            if (pastAverage > 0)
                cell3Press2Past.AddParagraph($"{pastAverage}");

            var cell3Press2Current = rowThirdPress.Cells[5];
            cell3Press2Current.Format.Font.Size = 8;
            cell3Press2Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothingThirdPress2);
            positionGoal = GetPositionGoal(2, 5);
            if (currentAge > 0)
                cell3Press2Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cell3Press2Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cell3Press2Current.Shading.Color = Colors.Yellow;
            }
            

            var cell3Press2Goal = rowThirdPress.Cells[6];
            cell3Press2Goal.Format.Font.Size = 8;
            cell3Press2Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cell3Press2Goal.AddParagraph($"{positionGoal}");

            var cell3Press2Info = rowThirdPress2.Cells[4];
            cell3Press2Info.Format.Font.Size = 8;
            cell3Press2Info.MergeRight = 2;
            cell3Press2Info.Shading.Color = Colors.LightBlue;
            cell3Press2Info.AddParagraph($"{GetClothingInfo(clothingThirdPress2)}");
            #endregion

            #region 3rd Press PM 3
            // ----- 3rd Press Position PM 3 ----- //
            var cell3Press3Past = rowThirdPress.Cells[7];
            cell3Press3Past.Format.Font.Size = 8;
            cell3Press3Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(3, 5);
            if (pastAverage > 0)
                cell3Press3Past.AddParagraph($"{pastAverage}");

            var cell3Press3Current = rowThirdPress.Cells[8];
            cell3Press3Current.Format.Font.Size = 8;
            cell3Press3Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothingThirdPress3);
            positionGoal = GetPositionGoal(3, 5);
            if (currentAge > 0)
                cell3Press3Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cell3Press3Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cell3Press3Current.Shading.Color = Colors.Yellow;
            }
            

            var cell3Press3Goal = rowThirdPress.Cells[9];
            cell3Press3Goal.Format.Font.Size = 8;
            cell3Press3Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cell3Press3Goal.AddParagraph($"{positionGoal}");

            var cell3Press3Info = rowThirdPress2.Cells[7];
            cell3Press3Info.Format.Font.Size = 8;
            cell3Press3Info.MergeRight = 2;
            cell3Press3Info.Shading.Color = Colors.LightBlue;
            cell3Press3Info.AddParagraph($"{GetClothingInfo(clothingThirdPress3)}");
            #endregion

            #region 3rd Press PM 4
            // ----- 3rd Press Position PM 4 ----- //
            var cell3Press4Past = rowThirdPress.Cells[10];
            cell3Press4Past.Format.Font.Size = 8;
            cell3Press4Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(4, 5);
            if (pastAverage > 0)
                cell3Press4Past.AddParagraph($"{pastAverage}");

            var cell3Press4Current = rowThirdPress.Cells[11];
            cell3Press4Current.Format.Font.Size = 8;
            cell3Press4Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothingThirdPress4);
            positionGoal = GetPositionGoal(4, 5);
            if (currentAge > 0)
                cell3Press4Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cell3Press4Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cell3Press4Current.Shading.Color = Colors.Yellow;
            }
            

            var cell3Press4Goal = rowThirdPress.Cells[12];
            cell3Press4Goal.Format.Font.Size = 8;
            cell3Press4Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cell3Press4Goal.AddParagraph($"{positionGoal}");

            var cell3Press4Info = rowThirdPress2.Cells[10];
            cell3Press4Info.Format.Font.Size = 8;
            cell3Press4Info.MergeRight = 2;
            cell3Press4Info.Shading.Color = Colors.LightBlue;
            cell3Press4Info.AddParagraph($"{GetClothingInfo(clothingThirdPress4)}");
            #endregion

            #endregion // 3rd Press Position

            #region 1st Top Dryer Position
            // add spacer row
            spacer = weeklyPMTable.AddRow();
            spacer.Height = 12;
            cellSpacer = spacer.Cells[0];
            cellSpacer.MergeRight = 12;
            cellSpacer.AddParagraph("");

            var clothingFirstTopDryer1 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "1st top dryer felt" && c.PM_Number == 1);
            var clothingFirstTopDryer2 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "1st top dryer felt" && c.PM_Number == 2);
            var clothingFirstTopDryer3 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "1st top dryer felt" && c.PM_Number == 3);
            var clothingFirstTopDryer4 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "1st top dryer felt" && c.PM_Number == 4);
            var rowFirstTopDryer = weeklyPMTable.AddRow();
            rowFirstTopDryer.Height = 10;
            var rowFirstTopDryer2 = weeklyPMTable.AddRow();
            rowFirstTopDryer2.Height = 10;

            var cellFirstTopDryer = rowFirstTopDryer.Cells[0];
            cellFirstTopDryer.Format.Font.Size = 8;
            cellFirstTopDryer.VerticalAlignment = VerticalAlignment.Center;
            cellFirstTopDryer.MergeDown = 1;
            cellFirstTopDryer.AddParagraph("1ST TOP");

            #region 1st Top PM 1
            // ----- 1st Top Position PM 1 ----- //
            var cellFirstTopDryer1Past = rowFirstTopDryer.Cells[1];
            cellFirstTopDryer1Past.Format.Font.Size = 8;
            cellFirstTopDryer1Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(1, 6);
            if (pastAverage > 0)
                cellFirstTopDryer1Past.AddParagraph($"{pastAverage}");

            var cellFirstTopDryer1Current = rowFirstTopDryer.Cells[2];
            cellFirstTopDryer1Current.Format.Font.Size = 8;
            cellFirstTopDryer1Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothingFirstTopDryer1);
            positionGoal = GetPositionGoal(1, 6);
            if (currentAge > 0)
                cellFirstTopDryer1Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cellFirstTopDryer1Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cellFirstTopDryer1Current.Shading.Color = Colors.Yellow;
            }
            

            var cellFirstTopDryer1Goal = rowFirstTopDryer.Cells[3];
            cellFirstTopDryer1Goal.Format.Font.Size = 8;
            cellFirstTopDryer1Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cellFirstTopDryer1Goal.AddParagraph($"{positionGoal}");

            var cellFirstTopDryer1Info = rowFirstTopDryer2.Cells[1];
            cellFirstTopDryer1Info.Format.Font.Size = 8;
            cellFirstTopDryer1Info.MergeRight = 2;
            cellFirstTopDryer1Info.Shading.Color = Colors.LightBlue;
            cellFirstTopDryer1Info.AddParagraph($"{GetClothingInfo(clothingFirstTopDryer1)}");
            #endregion

            #region 1st Top PM 2
            // ----- 1st Top Position PM 2 ----- //
            var cellFirstTopDryer2Past = rowFirstTopDryer.Cells[4];
            cellFirstTopDryer2Past.Format.Font.Size = 8;
            cellFirstTopDryer2Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(2, 6);
            if (pastAverage > 0)
                cellFirstTopDryer2Past.AddParagraph($"{pastAverage}");


            var cellFirstTopDryer2Current = rowFirstTopDryer.Cells[5];
            cellFirstTopDryer2Current.Format.Font.Size = 8;
            cellFirstTopDryer2Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothingFirstTopDryer2);
            positionGoal = GetPositionGoal(2, 6);
            if (currentAge > 0)
                cellFirstTopDryer2Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cellFirstTopDryer2Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cellFirstTopDryer2Current.Shading.Color = Colors.Yellow;
            }
            

            var cellFirstTopDryer2Goal = rowFirstTopDryer.Cells[6];
            cellFirstTopDryer2Goal.Format.Font.Size = 8;
            cellFirstTopDryer2Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cellFirstTopDryer2Goal.AddParagraph($"{positionGoal}");

            var cellFirstTopDryer2Info = rowFirstTopDryer2.Cells[4];
            cellFirstTopDryer2Info.Format.Font.Size = 8;
            cellFirstTopDryer2Info.MergeRight = 2;
            cellFirstTopDryer2Info.Shading.Color = Colors.LightBlue;
            cellFirstTopDryer2Info.AddParagraph($"{GetClothingInfo(clothingFirstTopDryer2)}");
            #endregion

            #region 1st Top PM 3
            // ----- 1st Top Position PM 3 ----- //
            var cellFirstTopDryer3Past = rowFirstTopDryer.Cells[7];
            cellFirstTopDryer3Past.Format.Font.Size = 8;
            cellFirstTopDryer3Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(3, 6);
            if (pastAverage > 0)
                cellFirstTopDryer3Past.AddParagraph($"{pastAverage}");

            var cellFirstTopDryer3Current = rowFirstTopDryer.Cells[8];
            cellFirstTopDryer3Current.Format.Font.Size = 8;
            cellFirstTopDryer3Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothingFirstTopDryer3);
            positionGoal = GetPositionGoal(3, 6);
            if (currentAge > 0)
                cellFirstTopDryer3Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cellFirstTopDryer3Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cellFirstTopDryer3Current.Shading.Color = Colors.Yellow;
            }
            

            var cellFirstTopDryer3Goal = rowFirstTopDryer.Cells[9];
            cellFirstTopDryer3Goal.Format.Font.Size = 8;
            cellFirstTopDryer3Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cellFirstTopDryer3Goal.AddParagraph($"{positionGoal}");

            var cellFirstTopDryer3Info = rowFirstTopDryer2.Cells[7];
            cellFirstTopDryer3Info.Format.Font.Size = 8;
            cellFirstTopDryer3Info.MergeRight = 2;
            cellFirstTopDryer3Info.Shading.Color = Colors.LightBlue;
            cellFirstTopDryer3Info.AddParagraph($"{GetClothingInfo(clothingFirstTopDryer3)}");

            #endregion

            #region 1st Top PM 4
            // ----- 1st Top Position PM 4 ----- //
            var cellFirstTopDryer4Past = rowFirstTopDryer.Cells[10];
            cellFirstTopDryer4Past.Format.Font.Size = 8;
            cellFirstTopDryer4Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(4, 6);
            if (pastAverage > 0)
                cellFirstTopDryer4Past.AddParagraph($"{pastAverage}");

            var cellFirstTopDryer4Current = rowFirstTopDryer.Cells[11];
            cellFirstTopDryer4Current.Format.Font.Size = 8;
            cellFirstTopDryer4Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothingFirstTopDryer4);
            positionGoal = GetPositionGoal(4, 6);
            if (currentAge > 0)
                cellFirstTopDryer4Current.AddParagraph($"{currentAge}");
            if (currentAge  > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cellFirstTopDryer4Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cellFirstTopDryer4Current.Shading.Color = Colors.Yellow;
            }
            

            var cellFirstTopDryer4Goal = rowFirstTopDryer.Cells[12];
            cellFirstTopDryer4Goal.Format.Font.Size = 8;
            cellFirstTopDryer4Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cellFirstTopDryer4Goal.AddParagraph($"{positionGoal}");

            var cellFirstTopDryer4Info = rowFirstTopDryer2.Cells[10];
            cellFirstTopDryer4Info.Format.Font.Size = 8;
            cellFirstTopDryer4Info.MergeRight = 2;
            cellFirstTopDryer4Info.Shading.Color = Colors.LightBlue;
            cellFirstTopDryer4Info.AddParagraph($"{GetClothingInfo(clothingFirstTopDryer4)}");
            #endregion

            #endregion // 1st Top Dryer Position

            #region 2nd Top Dryer Position
            // add spacer row
            spacer = weeklyPMTable.AddRow();
            spacer.Height = 12;
            cellSpacer = spacer.Cells[0];
            cellSpacer.MergeRight = 12;
            cellSpacer.AddParagraph("");

            var clothing2ndTopDryer1 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "2nd top dryer felt" && c.PM_Number == 1);
            var clothing2ndTopDryer2 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "2nd top dryer felt" && c.PM_Number == 2);
            var clothing2ndTopDryer3 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "2nd top dryer felt" && c.PM_Number == 3);
            var clothing2ndTopDryer4 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "2nd top dryer felt" && c.PM_Number == 4);
            var row2ndTopDryer = weeklyPMTable.AddRow();
            row2ndTopDryer.Height = 10;
            var row2ndTopDryer2 = weeklyPMTable.AddRow();
            row2ndTopDryer2.Height = 10;

            var cell2ndTopDryer = row2ndTopDryer.Cells[0];
            cell2ndTopDryer.Format.Font.Size = 8;
            cell2ndTopDryer.VerticalAlignment = VerticalAlignment.Center;
            cell2ndTopDryer.MergeDown = 1;
            cell2ndTopDryer.AddParagraph("2ND TOP");

            #region 2nd Top PM 1
            // ----- 2nd Top Position PM 1 ----- //
            var cell2ndTopDryer1Past = row2ndTopDryer.Cells[1];
            cell2ndTopDryer1Past.Format.Font.Size = 8;
            cell2ndTopDryer1Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(1, 7);
            if (pastAverage > 0)
                cell2ndTopDryer1Past.AddParagraph($"{pastAverage}");

            var cell2ndTopDryer1Current = row2ndTopDryer.Cells[2];
            cell2ndTopDryer1Current.Format.Font.Size = 8;
            cell2ndTopDryer1Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothing2ndTopDryer1);
            positionGoal = GetPositionGoal(1, 7);
            if (currentAge > 0)
                cell2ndTopDryer1Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cell2ndTopDryer1Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cell2ndTopDryer1Current.Shading.Color = Colors.Yellow;
            }
            

            var cell2ndTopDryer1Goal = row2ndTopDryer.Cells[3];
            cell2ndTopDryer1Goal.Format.Font.Size = 8;
            cell2ndTopDryer1Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cell2ndTopDryer1Goal.AddParagraph($"{positionGoal}");

            var cell2ndTopDryer1Info = row2ndTopDryer2.Cells[1];
            cell2ndTopDryer1Info.Format.Font.Size = 8;
            cell2ndTopDryer1Info.MergeRight = 2;
            cell2ndTopDryer1Info.VerticalAlignment = VerticalAlignment.Center;
            cell2ndTopDryer1Info.Shading.Color = Colors.LightBlue;
            cell2ndTopDryer1Info.AddParagraph($"{GetClothingInfo(clothing2ndTopDryer1)}");
            #endregion

            #region 2nd Top PM 2
            // ----- 2nd Top Position PM 2 ----- //
            var cell2ndTopDryer2Past = row2ndTopDryer.Cells[4];
            cell2ndTopDryer2Past.Format.Font.Size = 8;
            cell2ndTopDryer2Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(2, 7);
            if (pastAverage > 0)
                cell2ndTopDryer2Past.AddParagraph($"{pastAverage}");

            var cell2ndTopDryer2Current = row2ndTopDryer.Cells[5];
            cell2ndTopDryer2Current.Format.Font.Size = 8;
            cell2ndTopDryer2Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothing2ndTopDryer2);
            positionGoal = GetPositionGoal(2, 7);
            if (currentAge > 0)
                cell2ndTopDryer2Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cell2ndTopDryer2Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cell2ndTopDryer2Current.Shading.Color = Colors.Yellow;
            }
            

            var cell2ndTopDryer2Goal = row2ndTopDryer.Cells[6];
            cell2ndTopDryer2Goal.Format.Font.Size = 8;
            cell2ndTopDryer2Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cell2ndTopDryer2Goal.AddParagraph($"{positionGoal}");

            var cell2ndTopDryer2Info = row2ndTopDryer2.Cells[4];
            cell2ndTopDryer2Info.Format.Font.Size = 8;
            cell2ndTopDryer2Info.MergeRight = 2;
            cell2ndTopDryer2Info.VerticalAlignment = VerticalAlignment.Center;
            cell2ndTopDryer2Info.Shading.Color = Colors.LightBlue;
            cell2ndTopDryer2Info.AddParagraph($"{GetClothingInfo(clothing2ndTopDryer2)}");
            #endregion

            #region 2nd Top PM 3
            // ----- 2nd Top Position PM 3 ----- //
            var cell2ndTopDryer3Past = row2ndTopDryer.Cells[7];
            cell2ndTopDryer3Past.Format.Font.Size = 8;
            cell2ndTopDryer3Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(3, 7);
            if (pastAverage > 0)
                cell2ndTopDryer3Past.AddParagraph($"{pastAverage}");

            var cell2ndTopDryer3Current = row2ndTopDryer.Cells[8];
            cell2ndTopDryer3Current.Format.Font.Size = 8;
            cell2ndTopDryer3Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothing2ndTopDryer3);
            positionGoal = GetPositionGoal(3, 7);
            if (currentAge > 0)
                cell2ndTopDryer3Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cell2ndTopDryer3Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cell2ndTopDryer3Current.Shading.Color = Colors.Yellow;
            }
            

            var cell2ndTopDryer3Goal = row2ndTopDryer.Cells[9];
            cell2ndTopDryer3Goal.Format.Font.Size = 8;
            cell2ndTopDryer3Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cell2ndTopDryer3Goal.AddParagraph($"{positionGoal}");

            var cell2ndTopDryer3Info = row2ndTopDryer2.Cells[7];
            cell2ndTopDryer3Info.Format.Font.Size = 8;
            cell2ndTopDryer3Info.MergeRight = 2;
            cell2ndTopDryer3Info.VerticalAlignment = VerticalAlignment.Center;
            cell2ndTopDryer3Info.Shading.Color = Colors.LightBlue;
            cell2ndTopDryer3Info.AddParagraph($"{GetClothingInfo(clothing2ndTopDryer3)}");
            #endregion

            #region 2nd Top PM 4
            // ----- 2nd Top Position PM 4 ----- //
            var cell2ndTopDryer4Past = row2ndTopDryer.Cells[10];
            cell2ndTopDryer4Past.Format.Font.Size = 8;
            cell2ndTopDryer4Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(4, 7);
            if (pastAverage > 0)
                cell2ndTopDryer4Past.AddParagraph($"{pastAverage}");

            var cell2ndTopDryer4Current = row2ndTopDryer.Cells[11];
            cell2ndTopDryer4Current.Format.Font.Size = 8;
            cell2ndTopDryer4Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothing2ndTopDryer4);
            positionGoal = GetPositionGoal(4, 7);
            if (currentAge > 0)
                cell2ndTopDryer4Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cell2ndTopDryer4Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cell2ndTopDryer4Current.Shading.Color = Colors.Yellow;
            }
            

            var cell2ndTopDryer4Goal = row2ndTopDryer.Cells[12];
            cell2ndTopDryer4Goal.Format.Font.Size = 8;
            cell2ndTopDryer4Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cell2ndTopDryer4Goal.AddParagraph($"{positionGoal}");

            var cell2ndTopDryer4Info = row2ndTopDryer2.Cells[10];
            cell2ndTopDryer4Info.Format.Font.Size = 8;
            cell2ndTopDryer4Info.MergeRight = 2;
            cell2ndTopDryer4Info.VerticalAlignment = VerticalAlignment.Center;
            cell2ndTopDryer4Info.Shading.Color = Colors.LightBlue;
            cell2ndTopDryer4Info.AddParagraph($"{GetClothingInfo(clothing2ndTopDryer4)}");
            #endregion

            #endregion //2nd Top Dryer Position

            #region 3rd Top Dryer Position
            // add spacer row
            spacer = weeklyPMTable.AddRow();
            spacer.Height = 12;
            cellSpacer = spacer.Cells[0];
            cellSpacer.MergeRight = 12;
            cellSpacer.AddParagraph("");

            var clothing3rdTopDryer1 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "3rd top dryer felt" && c.PM_Number == 1);
            var clothing3rdTopDryer2 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "3rd top dryer felt" && c.PM_Number == 2);
            var clothing3rdTopDryer3 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "3rd top dryer felt" && c.PM_Number == 3);
            var clothing3rdTopDryer4 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "3rd top dryer felt" && c.PM_Number == 4);
            var row3rdTopDryer = weeklyPMTable.AddRow();
            row3rdTopDryer.Height = 10;
            var row3rdTopDryer2 = weeklyPMTable.AddRow();
            row3rdTopDryer2.Height = 10;

            var cell3rdTopDryer = row3rdTopDryer.Cells[0];
            cell3rdTopDryer.Format.Font.Size = 8;
            cell3rdTopDryer.VerticalAlignment = VerticalAlignment.Center;
            cell3rdTopDryer.MergeDown = 1;
            cell3rdTopDryer.AddParagraph("3RD TOP");

            #region 3rd Top PM 1
            // ----- 3rd Top Position PM 1 ----- //
            var cell3rdTopDryer1Past = row3rdTopDryer.Cells[1];
            cell3rdTopDryer1Past.Format.Font.Size = 8;
            cell3rdTopDryer1Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(1, 8);
            if (pastAverage > 0)
                cell3rdTopDryer1Past.AddParagraph($"{pastAverage}");

            var cell3rdTopDryer1Current = row3rdTopDryer.Cells[2];
            cell3rdTopDryer1Current.Format.Font.Size = 8;
            cell3rdTopDryer1Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothing3rdTopDryer1);
            positionGoal = GetPositionGoal(1, 8);
            if (currentAge > 0)
                cell3rdTopDryer1Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cell3rdTopDryer1Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cell3rdTopDryer1Current.Shading.Color = Colors.Yellow;
            }
            

            var cell3rdTopDryer1Goal = row3rdTopDryer.Cells[3];
            cell3rdTopDryer1Goal.Format.Font.Size = 8;
            cell3rdTopDryer1Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cell3rdTopDryer1Goal.AddParagraph($"{positionGoal}");

            var cell3rdTopDryer1Info = row3rdTopDryer2.Cells[1];
            cell3rdTopDryer1Info.Format.Font.Size = 8;
            cell3rdTopDryer1Info.MergeRight = 2;
            cell3rdTopDryer1Info.VerticalAlignment = VerticalAlignment.Center;
            cell3rdTopDryer1Info.Shading.Color = Colors.LightBlue;
            cell3rdTopDryer1Info.AddParagraph($"{GetClothingInfo(clothing3rdTopDryer1)}");
            #endregion

            #region 3rd Top PM 2
            // ----- 3rd Top Position PM 2 ----- //
            var cell3rdTopDryer2Past = row3rdTopDryer.Cells[4];
            cell3rdTopDryer2Past.Format.Font.Size = 8;
            cell3rdTopDryer2Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(2, 8);
            if (pastAverage > 0)
                cell3rdTopDryer2Past.AddParagraph($"{pastAverage}");

            var cell3rdTopDryer2Current = row3rdTopDryer.Cells[5];
            cell3rdTopDryer2Current.Format.Font.Size = 8;
            cell3rdTopDryer2Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothing3rdTopDryer2);
            positionGoal = GetPositionGoal(2, 8);
            if (currentAge > 0)
                cell3rdTopDryer2Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cell3rdTopDryer2Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cell3rdTopDryer2Current.Shading.Color = Colors.Yellow;
            }
            

            var cell3rdTopDryer2Goal = row3rdTopDryer.Cells[6];
            cell3rdTopDryer2Goal.Format.Font.Size = 8;
            cell3rdTopDryer2Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cell3rdTopDryer2Goal.AddParagraph($"{positionGoal}");

            var cell3rdTopDryer2Info = row3rdTopDryer2.Cells[4];
            cell3rdTopDryer2Info.Format.Font.Size = 8;
            cell3rdTopDryer2Info.MergeRight = 2;
            cell3rdTopDryer2Info.VerticalAlignment = VerticalAlignment.Center;
            cell3rdTopDryer2Info.Shading.Color = Colors.LightBlue;
            cell3rdTopDryer2Info.AddParagraph($"{GetClothingInfo(clothing3rdTopDryer2)}");
            #endregion

            #region 3rd Top PM 3
            // ----- 3rd Top Position PM 3 ----- //
            var cell3rdTopDryer3Past = row3rdTopDryer.Cells[7];
            cell3rdTopDryer3Past.Format.Font.Size = 8;
            cell3rdTopDryer3Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(3, 8);
            if (pastAverage > 0)
                cell3rdTopDryer3Past.AddParagraph($"{pastAverage}");

            var cell3rdTopDryer3Current = row3rdTopDryer.Cells[8];
            cell3rdTopDryer3Current.Format.Font.Size = 8;
            cell3rdTopDryer3Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothing3rdTopDryer3);
            positionGoal = GetPositionGoal(3, 8);
            if (currentAge > 0)
                cell3rdTopDryer3Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cell3rdTopDryer3Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cell3rdTopDryer3Current.Shading.Color = Colors.Yellow;
            }
            

            var cell3rdTopDryer3Goal = row3rdTopDryer.Cells[9];
            cell3rdTopDryer3Goal.Format.Font.Size = 8;
            cell3rdTopDryer3Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cell3rdTopDryer3Goal.AddParagraph($"{positionGoal}");

            var cell3rdTopDryer3Info = row3rdTopDryer2.Cells[7];
            cell3rdTopDryer3Info.Format.Font.Size = 8;
            cell3rdTopDryer3Info.MergeRight = 2;
            cell3rdTopDryer3Info.VerticalAlignment = VerticalAlignment.Center;
            cell3rdTopDryer3Info.Shading.Color = Colors.LightBlue;
            cell3rdTopDryer3Info.AddParagraph($"{GetClothingInfo(clothing3rdTopDryer3)}");
            #endregion

            #region 3rd Top PM 4
            // ----- 3rd Top Position PM 4 ----- //
            var cell3rdTopDryer4Past = row3rdTopDryer.Cells[10];
            cell3rdTopDryer4Past.Format.Font.Size = 8;
            cell3rdTopDryer4Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(4, 8);
            if (pastAverage > 0)
                cell3rdTopDryer4Past.AddParagraph($"{pastAverage}");

            var cell3rdTopDryer4Current = row3rdTopDryer.Cells[11];
            cell3rdTopDryer4Current.Format.Font.Size = 8;
            cell3rdTopDryer4Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothing3rdTopDryer4);
            positionGoal = GetPositionGoal(4, 8);
            if (currentAge > 0)
                cell3rdTopDryer4Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cell3rdTopDryer4Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cell3rdTopDryer4Current.Shading.Color = Colors.Yellow;
            }
            

            var cell3rdTopDryer4Goal = row3rdTopDryer.Cells[12];
            cell3rdTopDryer4Goal.Format.Font.Size = 8;
            cell3rdTopDryer4Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cell3rdTopDryer4Goal.AddParagraph($"{positionGoal}");

            var cell3rdTopDryer4Info = row3rdTopDryer2.Cells[10];
            cell3rdTopDryer4Info.Format.Font.Size = 8;
            cell3rdTopDryer4Info.MergeRight = 2;
            cell3rdTopDryer4Info.VerticalAlignment = VerticalAlignment.Center;
            cell3rdTopDryer4Info.Shading.Color = Colors.LightBlue;
            cell3rdTopDryer4Info.AddParagraph($"{GetClothingInfo(clothing3rdTopDryer4)}");
            #endregion

            #endregion // 3rd Top Dryer Position

            #region 4th Top Dryer Position
            // add spacer row
            spacer = weeklyPMTable.AddRow();
            spacer.Height = 12;
            cellSpacer = spacer.Cells[0];
            cellSpacer.MergeRight = 12;
            cellSpacer.AddParagraph("");

            var clothing4thTopDryer1 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "4th top dryer felt" && c.PM_Number == 1);
            var clothing4thTopDryer2 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "4th top dryer felt" && c.PM_Number == 2);
            var clothing4thTopDryer3 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "4th top dryer felt" && c.PM_Number == 3);
            var clothing4thTopDryer4 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "4th top dryer felt" && c.PM_Number == 4);
            var row4thTopDryer = weeklyPMTable.AddRow();
            row4thTopDryer.Height = 10;
            var row4thTopDryer2 = weeklyPMTable.AddRow();
            row4thTopDryer2.Height = 10;

            var cell4thTopDryer = row4thTopDryer.Cells[0];
            cell4thTopDryer.Format.Font.Size = 8;
            cell4thTopDryer.VerticalAlignment = VerticalAlignment.Center;
            cell4thTopDryer.MergeDown = 1;
            cell4thTopDryer.AddParagraph("4TH TOP");

            #region 4th Top PM 1
            // ----- 4th Top Position PM 1 ----- //
            var cell4thTopDryer1Past = row4thTopDryer.Cells[1];
            cell4thTopDryer1Past.Format.Font.Size = 8;
            cell4thTopDryer1Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(1, 9);
            if (pastAverage > 0)
                cell4thTopDryer1Past.AddParagraph($"{pastAverage}");

            var cell4thTopDryer1Current = row4thTopDryer.Cells[2];
            cell4thTopDryer1Current.Format.Font.Size = 8;
            cell4thTopDryer1Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothing4thTopDryer1);
            positionGoal = GetPositionGoal(1, 9);
            if (currentAge > 0)
                cell4thTopDryer1Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cell4thTopDryer1Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cell4thTopDryer1Current.Shading.Color = Colors.Yellow;
            }
            

            var cell4thTopDryer1Goal = row4thTopDryer.Cells[3];
            cell4thTopDryer1Goal.Format.Font.Size = 8;
            cell4thTopDryer1Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cell4thTopDryer1Goal.AddParagraph($"{positionGoal}");

            var cell4thTopDryer1Info = row4thTopDryer2.Cells[1];
            cell4thTopDryer1Info.Format.Font.Size = 8;
            cell4thTopDryer1Info.MergeRight = 2;
            cell4thTopDryer1Info.VerticalAlignment = VerticalAlignment.Center;
            cell4thTopDryer1Info.Shading.Color = Colors.LightBlue;
            cell4thTopDryer1Info.AddParagraph($"{GetClothingInfo(clothing4thTopDryer1)}");
            #endregion

            #region 4th Top PM 2
            // ----- 4th Top Position PM 2 ----- //
            var cell4thTopDryer2Past = row4thTopDryer.Cells[4];
            cell4thTopDryer2Past.Format.Font.Size = 8;
            cell4thTopDryer2Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(2, 9);
            if (pastAverage > 0)
                cell4thTopDryer2Past.AddParagraph($"{pastAverage}");

            var cell4thTopDryer2Current = row4thTopDryer.Cells[5];
            cell4thTopDryer2Current.Format.Font.Size = 8;
            cell4thTopDryer2Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothing4thTopDryer2);
            positionGoal = GetPositionGoal(2, 9);
            if (currentAge > 0)
                cell4thTopDryer2Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cell4thTopDryer2Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cell4thTopDryer2Current.Shading.Color = Colors.Yellow;
            }
            

            var cell4thTopDryer2Goal = row4thTopDryer.Cells[6];
            cell4thTopDryer2Goal.Format.Font.Size = 8;
            cell4thTopDryer2Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cell4thTopDryer2Goal.AddParagraph($"{positionGoal}");

            var cell4thTopDryer2Info = row4thTopDryer2.Cells[4];
            cell4thTopDryer2Info.Format.Font.Size = 8;
            cell4thTopDryer2Info.MergeRight = 2;
            cell4thTopDryer2Info.VerticalAlignment = VerticalAlignment.Center;
            cell4thTopDryer2Info.Shading.Color = Colors.LightBlue;
            cell4thTopDryer2Info.AddParagraph($"{GetClothingInfo(clothing4thTopDryer2)}");

            #endregion

            #region 4th Top PM 3
            // ----- 4th Top Position PM 3 ----- //
            var cell4thTopDryer3Past = row4thTopDryer.Cells[7];
            cell4thTopDryer3Past.Format.Font.Size = 8;
            cell4thTopDryer3Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(3, 9);
            if (pastAverage > 0)
                cell4thTopDryer3Past.AddParagraph($"{pastAverage}");

            var cell4thTopDryer3Current = row4thTopDryer.Cells[8];
            cell4thTopDryer3Current.Format.Font.Size = 8;
            cell4thTopDryer3Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothing4thTopDryer3);
            positionGoal = GetPositionGoal(3, 9);
            if (currentAge > 0)
                cell4thTopDryer3Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cell4thTopDryer3Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cell4thTopDryer3Current.Shading.Color = Colors.Yellow;
            }
            

            var cell4thTopDryer3Goal = row4thTopDryer.Cells[9];
            cell4thTopDryer3Goal.Format.Font.Size = 8;
            cell4thTopDryer3Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cell4thTopDryer3Goal.AddParagraph($"{positionGoal}");

            var cell4thTopDryer3Info = row4thTopDryer2.Cells[7];
            cell4thTopDryer3Info.Format.Font.Size = 8;
            cell4thTopDryer3Info.MergeRight = 2;
            cell4thTopDryer3Info.VerticalAlignment = VerticalAlignment.Center;
            cell4thTopDryer3Info.Shading.Color = Colors.LightBlue;
            cell4thTopDryer3Info.AddParagraph($"{GetClothingInfo(clothing4thTopDryer3)}");
            #endregion

            #region 4th Top PM 4
            // ----- 4th Top Position PM 4 ----- //
            var cell4thTopDryer4Past = row4thTopDryer.Cells[10];
            cell4thTopDryer4Past.Format.Font.Size = 8;
            cell4thTopDryer4Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(4, 9);
            if (pastAverage > 0)
                cell4thTopDryer4Past.AddParagraph($"{pastAverage}"); ;

            var cell4thTopDryer4Current = row4thTopDryer.Cells[11];
            cell4thTopDryer4Current.Format.Font.Size = 8;
            cell4thTopDryer4Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothing4thTopDryer4);
            positionGoal = GetPositionGoal(4, 9);
            if (currentAge > 0)
                cell4thTopDryer4Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cell4thTopDryer4Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cell4thTopDryer4Current.Shading.Color = Colors.Yellow;
            }
            

            var cell4thTopDryer4Goal = row4thTopDryer.Cells[12];
            cell4thTopDryer4Goal.Format.Font.Size = 8;
            cell4thTopDryer4Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cell4thTopDryer4Goal.AddParagraph($"{positionGoal}");

            var cell4thTopDryer4Info = row4thTopDryer2.Cells[10];
            cell4thTopDryer4Info.Format.Font.Size = 8;
            cell4thTopDryer4Info.MergeRight = 2;
            cell4thTopDryer4Info.VerticalAlignment = VerticalAlignment.Center;
            cell4thTopDryer4Info.Shading.Color = Colors.LightBlue;
            cell4thTopDryer4Info.AddParagraph($"{GetClothingInfo(clothing4thTopDryer4)}");
            #endregion

            #endregion // 4th Top Dryer Position

            #region 1st Bottom Dryer Position
            // add spacer row
            spacer = weeklyPMTable.AddRow();
            spacer.Height = 12;
            cellSpacer = spacer.Cells[0];
            cellSpacer.MergeRight = 12;
            cellSpacer.AddParagraph("");

            var clothing1stBottomDryer1 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "1st bottom dryer felt" && c.PM_Number == 1);
            var clothing1stBottomDryer2 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "1st bottom dryer felt" && c.PM_Number == 2);
            var clothing1stBottomDryer3 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "1st bottom dryer felt" && c.PM_Number == 3);
            var clothing1stBottomDryer4 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "1st bottom dryer felt" && c.PM_Number == 4);
            var row1stBottomDryer = weeklyPMTable.AddRow();
            row1stBottomDryer.Height = 10;
            var row1stBottomDryer2 = weeklyPMTable.AddRow();
            row1stBottomDryer2.Height = 10;

            var cell1stBottomDryer = row1stBottomDryer.Cells[0];
            cell1stBottomDryer.Format.Font.Size = 8;
            cell1stBottomDryer.VerticalAlignment = VerticalAlignment.Center;
            cell1stBottomDryer.MergeDown = 1;
            cell1stBottomDryer.AddParagraph("1ST BOTTOM");

            #region 1st Bottom PM 1
            // ----- 1st Bottom Position PM 1 ----- //
            var cell1stBottomDryer1Past = row1stBottomDryer.Cells[1];
            cell1stBottomDryer1Past.Format.Font.Size = 8;
            cell1stBottomDryer1Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(1, 10);
            if (pastAverage > 0)
                cell1stBottomDryer1Past.AddParagraph($"{pastAverage}");

            var cell1stBottomDryer1Current = row1stBottomDryer.Cells[2];
            cell1stBottomDryer1Current.Format.Font.Size = 8;
            cell1stBottomDryer1Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothing1stBottomDryer1);
            positionGoal = GetPositionGoal(1, 10);
            if (currentAge > 0)
                cell1stBottomDryer1Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cell1stBottomDryer1Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cell1stBottomDryer1Current.Shading.Color = Colors.Yellow;
            }
            

            var cell1stBottomDryer1Goal = row1stBottomDryer.Cells[3];
            cell1stBottomDryer1Goal.Format.Font.Size = 8;
            cell1stBottomDryer1Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cell1stBottomDryer1Goal.AddParagraph($"{positionGoal}");

            var cell1stBottomDryer1Info = row1stBottomDryer2.Cells[1];
            cell1stBottomDryer1Info.Format.Font.Size = 8;
            cell1stBottomDryer1Info.MergeRight = 2;
            cell1stBottomDryer1Info.VerticalAlignment = VerticalAlignment.Center;
            cell1stBottomDryer1Info.Shading.Color = Colors.LightBlue;
            cell1stBottomDryer1Info.AddParagraph($"{GetClothingInfo(clothing1stBottomDryer1)}");
            #endregion

            #region 1st Bottom PM 2
            // ----- 1st Bottom Position PM 2 ----- //
            var cell1stBottomDryer2Past = row1stBottomDryer.Cells[4];
            cell1stBottomDryer2Past.Format.Font.Size = 8;
            cell1stBottomDryer2Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(2, 10);
            if (pastAverage > 0)
                cell1stBottomDryer2Past.AddParagraph($"{GetPastAverage(2, 10)}");

            var cell1stBottomDryer2Current = row1stBottomDryer.Cells[5];
            cell1stBottomDryer2Current.Format.Font.Size = 8;
            cell1stBottomDryer2Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothing1stBottomDryer2);
            positionGoal = GetPositionGoal(2, 10);
            if (currentAge > 0)
                cell1stBottomDryer2Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cell1stBottomDryer2Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cell1stBottomDryer2Current.Shading.Color = Colors.Yellow;
            }
           

            var cell1stBottomDryer2Goal = row1stBottomDryer.Cells[6];
            cell1stBottomDryer2Goal.Format.Font.Size = 8;
            cell1stBottomDryer2Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cell1stBottomDryer2Goal.AddParagraph($"{positionGoal}");

            var cell1stBottomDryer2Info = row1stBottomDryer2.Cells[4];
            cell1stBottomDryer2Info.Format.Font.Size = 8;
            cell1stBottomDryer2Info.MergeRight = 2;
            cell1stBottomDryer2Info.VerticalAlignment = VerticalAlignment.Center;
            cell1stBottomDryer2Info.Shading.Color = Colors.LightBlue;
            cell1stBottomDryer2Info.AddParagraph($"{GetClothingInfo(clothing1stBottomDryer2)}");
            #endregion

            #region 1st Bottom PM 3
            // ----- 1st Bottom Position PM 3 ----- //
            var cell1stBottomDryer3Past = row1stBottomDryer.Cells[7];
            cell1stBottomDryer3Past.Format.Font.Size = 8;
            cell1stBottomDryer3Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(3, 10);
            if (pastAverage > 0)
                cell1stBottomDryer3Past.AddParagraph($"{pastAverage}");

            var cell1stBottomDryer3Current = row1stBottomDryer.Cells[8];
            cell1stBottomDryer3Current.Format.Font.Size = 8;
            cell1stBottomDryer3Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothing1stBottomDryer3);
            positionGoal = GetPositionGoal(3, 10);
            if (currentAge > 0)
                cell1stBottomDryer3Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cell1stBottomDryer3Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cell1stBottomDryer3Current.Shading.Color = Colors.Yellow;
            }
            

            var cell1stBottomDryer3Goal = row1stBottomDryer.Cells[9];
            cell1stBottomDryer3Goal.Format.Font.Size = 8;
            cell1stBottomDryer3Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cell1stBottomDryer3Goal.AddParagraph($"{positionGoal}");

            var cell1stBottomDryer3Info = row1stBottomDryer2.Cells[7];
            cell1stBottomDryer3Info.Format.Font.Size = 8;
            cell1stBottomDryer3Info.MergeRight = 2;
            cell1stBottomDryer3Info.VerticalAlignment = VerticalAlignment.Center;
            cell1stBottomDryer3Info.Shading.Color = Colors.LightBlue;
            cell1stBottomDryer3Info.AddParagraph($"{GetClothingInfo(clothing1stBottomDryer3)}");
            #endregion

            #region 1st Bottom PM 4
            // ----- 1st Bottom Position PM 4 ----- //
            var cell1stBottomDryer4Past = row1stBottomDryer.Cells[10];
            cell1stBottomDryer4Past.Format.Font.Size = 8;
            cell1stBottomDryer4Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(4, 10);
            if (pastAverage > 0)
                cell1stBottomDryer4Past.AddParagraph($"{pastAverage}");

            var cell1stBottomDryer4Current = row1stBottomDryer.Cells[11];
            cell1stBottomDryer4Current.Format.Font.Size = 8;
            cell1stBottomDryer4Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothing1stBottomDryer4);
            positionGoal = GetPositionGoal(4, 10);
            if (currentAge > 0)
                cell1stBottomDryer4Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cell1stBottomDryer4Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cell1stBottomDryer4Current.Shading.Color = Colors.Yellow;
            }
            

            var cell1stBottomDryer4Goal = row1stBottomDryer.Cells[12];
            cell1stBottomDryer4Goal.Format.Font.Size = 8;
            cell1stBottomDryer4Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cell1stBottomDryer4Goal.AddParagraph($"{positionGoal}");

            var cell1stBottomDryer4Info = row1stBottomDryer2.Cells[10];
            cell1stBottomDryer4Info.Format.Font.Size = 8;
            cell1stBottomDryer4Info.MergeRight = 2;
            cell1stBottomDryer4Info.VerticalAlignment = VerticalAlignment.Center;
            cell1stBottomDryer4Info.Shading.Color = Colors.LightBlue;
            cell1stBottomDryer4Info.AddParagraph($"{GetClothingInfo(clothing1stBottomDryer4)}");
            #endregion

            #endregion // 1st Bottom Dryer Position

            #region 2nd Bottom Dryer Position
            // add spacer row
            spacer = weeklyPMTable.AddRow();
            spacer.Height = 12;
            cellSpacer = spacer.Cells[0];
            cellSpacer.MergeRight = 12;
            cellSpacer.AddParagraph("");

            var clothing2ndBottomDryer1 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "2nd bottom dryer felt" && c.PM_Number == 1);
            var clothing2ndBottomDryer2 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "2nd bottom dryer felt" && c.PM_Number == 2);
            var clothing2ndBottomDryer3 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "2nd bottom dryer felt" && c.PM_Number == 3);
            var clothing2ndBottomDryer4 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "2nd bottom dryer felt" && c.PM_Number == 4);
            var row2ndBottomDryer = weeklyPMTable.AddRow();
            row2ndBottomDryer.Height = 10;
            var row2ndBottomDryer2 = weeklyPMTable.AddRow();
            row2ndBottomDryer2.Height = 10;

            var cell2ndBottomDryer = row2ndBottomDryer.Cells[0];
            cell2ndBottomDryer.Format.Font.Size = 8;
            cell2ndBottomDryer.VerticalAlignment = VerticalAlignment.Center;
            cell2ndBottomDryer.MergeDown = 1;
            cell2ndBottomDryer.AddParagraph("2ND BOTTOM");

            #region 2nd Bottom PM 1
            // ----- 2nd Bottom Position PM 1 ----- //
            var cell2ndBottomDryer1Past = row2ndBottomDryer.Cells[1];
            cell2ndBottomDryer1Past.Format.Font.Size = 8;
            cell2ndBottomDryer1Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(1, 11);
            if (pastAverage > 0)
                cell2ndBottomDryer1Past.AddParagraph($"{pastAverage}");

            var cell2ndtBottomDryer1Current = row2ndBottomDryer.Cells[2];
            cell2ndtBottomDryer1Current.Format.Font.Size = 8;
            cell2ndtBottomDryer1Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothing2ndBottomDryer1);
            positionGoal = GetPositionGoal(1, 11);
            if (currentAge > 0)
                cell2ndtBottomDryer1Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cell2ndtBottomDryer1Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cell2ndtBottomDryer1Current.Shading.Color = Colors.Yellow;
            }
            

            var cell2ndBottomDryer1Goal = row2ndBottomDryer.Cells[3];
            cell2ndBottomDryer1Goal.Format.Font.Size = 8;
            cell2ndBottomDryer1Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cell2ndBottomDryer1Goal.AddParagraph($"{positionGoal}");

            var cell2ndBottomDryer1Info = row2ndBottomDryer2.Cells[1];
            cell2ndBottomDryer1Info.Format.Font.Size = 8;
            cell2ndBottomDryer1Info.MergeRight = 2;
            cell2ndBottomDryer1Info.VerticalAlignment = VerticalAlignment.Center;
            cell2ndBottomDryer1Info.Shading.Color = Colors.LightBlue;
            cell2ndBottomDryer1Info.AddParagraph($"{GetClothingInfo(clothing2ndBottomDryer1)}");
            #endregion

            #region 2nd Bottom PM 2
            // ----- 2nd Bottom Position PM 2 ----- //
            var cell2ndBottomDryer2Past = row2ndBottomDryer.Cells[4];
            cell2ndBottomDryer2Past.Format.Font.Size = 8;
            cell2ndBottomDryer2Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(2, 11);
            if (pastAverage > 0)
                cell2ndBottomDryer2Past.AddParagraph($"{pastAverage}");

            var cell2ndBottomDryer2Current = row2ndBottomDryer.Cells[5];
            cell2ndBottomDryer2Current.Format.Font.Size = 8;
            cell2ndBottomDryer2Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothing2ndBottomDryer2);
            positionGoal = GetPositionGoal(2, 11);
            if (currentAge > 0)
                cell2ndBottomDryer2Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cell2ndBottomDryer2Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cell2ndBottomDryer2Current.Shading.Color = Colors.Yellow;
            }
            

            var cell2ndBottomDryer2Goal = row2ndBottomDryer.Cells[6];
            cell2ndBottomDryer2Goal.Format.Font.Size = 8;
            cell2ndBottomDryer2Goal.VerticalAlignment = VerticalAlignment.Center;
            var secondBottomDryer2Goal = db.Goals.SingleOrDefault(g => g.PM_ID == 2 && g.PositionID == 11);
            if (positionGoal > 0)
                cell2ndBottomDryer2Goal.AddParagraph($"{positionGoal}");

            var cell2ndBottomDryer2Info = row2ndBottomDryer2.Cells[4];
            cell2ndBottomDryer2Info.Format.Font.Size = 8;
            cell2ndBottomDryer2Info.MergeRight = 2;
            cell2ndBottomDryer2Info.VerticalAlignment = VerticalAlignment.Center;
            cell2ndBottomDryer2Info.Shading.Color = Colors.LightBlue;
            cell2ndBottomDryer2Info.AddParagraph($"{GetClothingInfo(clothing2ndBottomDryer2)}");
            #endregion

            #region 2nd Bottom PM 3
            // ----- 2nd Bottom Position PM 3 ----- //
            var cell2ndBottomDryer3Past = row2ndBottomDryer.Cells[7];
            cell2ndBottomDryer3Past.Format.Font.Size = 8;
            cell2ndBottomDryer3Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(3, 11);
            if (pastAverage > 0)
                cell2ndBottomDryer3Past.AddParagraph($"{pastAverage}");

            var cell2ndBottomDryer3Current = row2ndBottomDryer.Cells[8];
            cell2ndBottomDryer3Current.Format.Font.Size = 8;
            cell2ndBottomDryer3Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothing2ndBottomDryer3);
            positionGoal = GetPositionGoal(3, 11);
            if (currentAge > 0)
                cell2ndBottomDryer3Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cell2ndBottomDryer3Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cell2ndBottomDryer3Current.Shading.Color = Colors.Yellow;
            }
            

            var cell2ndBottomDryer3Goal = row2ndBottomDryer.Cells[9];
            cell2ndBottomDryer3Goal.Format.Font.Size = 8;
            cell2ndBottomDryer3Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cell2ndBottomDryer3Goal.AddParagraph($"{positionGoal}");

            var cell2ndBottomDryer3Info = row2ndBottomDryer2.Cells[7];
            cell2ndBottomDryer3Info.Format.Font.Size = 8;
            cell2ndBottomDryer3Info.MergeRight = 2;
            cell2ndBottomDryer3Info.VerticalAlignment = VerticalAlignment.Center;
            cell2ndBottomDryer3Info.Shading.Color = Colors.LightBlue;
            cell2ndBottomDryer3Info.AddParagraph($"{GetClothingInfo(clothing2ndBottomDryer3)}");
            #endregion

            #region 2nd Bottom PM 4
            // ----- 2nd Bottom Position PM 4 ----- //
            var cell2ndBottomDryer4Past = row2ndBottomDryer.Cells[10];
            cell2ndBottomDryer4Past.Format.Font.Size = 8;
            cell2ndBottomDryer4Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(4, 11);
            if (pastAverage > 0)
                cell2ndBottomDryer4Past.AddParagraph($"{pastAverage}");

            var cell2ndBottomDryer4Current = row2ndBottomDryer.Cells[11];
            cell2ndBottomDryer4Current.Format.Font.Size = 8;
            cell2ndBottomDryer4Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothing2ndBottomDryer4);
            positionGoal = GetPositionGoal(4, 11);
            if (currentAge > 0)
                cell2ndBottomDryer4Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || pastAverage > 0)
            {
                if (currentAge >= positionGoal)
                    cell2ndBottomDryer4Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cell2ndBottomDryer4Current.Shading.Color = Colors.Yellow;
            }
            

            var cell2ndBottomDryer4Goal = row2ndBottomDryer.Cells[12];
            cell2ndBottomDryer4Goal.Format.Font.Size = 8;
            cell2ndBottomDryer4Goal.VerticalAlignment = VerticalAlignment.Center;
            cell2ndBottomDryer4Goal.AddParagraph($"{positionGoal}");

            var cell2ndBottomDryer4Info = row2ndBottomDryer2.Cells[10];
            cell2ndBottomDryer4Info.Format.Font.Size = 8;
            cell2ndBottomDryer4Info.MergeRight = 2;
            cell2ndBottomDryer4Info.VerticalAlignment = VerticalAlignment.Center;
            cell2ndBottomDryer4Info.Shading.Color = Colors.LightBlue;
            cell2ndBottomDryer4Info.AddParagraph($"{GetClothingInfo(clothing2ndBottomDryer4)}");
            #endregion

            #endregion // 2nd Bottom Dryer Position

            #region 3rd Bottom Dryer Position
            // add spacer row
            spacer = weeklyPMTable.AddRow();
            spacer.Height = 12;
            cellSpacer = spacer.Cells[0];
            cellSpacer.MergeRight = 12;
            cellSpacer.AddParagraph("");

            var clothing3rdBottomDryer1 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "3rd bottom dryer felt" && c.PM_Number == 1);
            var clothing3rdBottomDryer2 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "3rd bottom dryer felt" && c.PM_Number == 2);
            var clothing3rdBottomDryer3 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "3rd bottom dryer felt" && c.PM_Number == 3);
            var clothing3rdBottomDryer4 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "3rd bottom dryer felt" && c.PM_Number == 4);
            var row3rdBottomDryer = weeklyPMTable.AddRow();
            row3rdBottomDryer.Height = 10;
            var row3rdBottomDryer2 = weeklyPMTable.AddRow();
            row3rdBottomDryer2.Height = 10;

            var cell3rdBottomDryer = row3rdBottomDryer.Cells[0];
            cell3rdBottomDryer.Format.Font.Size = 8;
            cell3rdBottomDryer.VerticalAlignment = VerticalAlignment.Center;
            cell3rdBottomDryer.MergeDown = 1;
            cell3rdBottomDryer.AddParagraph("3RD BOTTOM");

            #region 3rd Bottom PM 1
            // ----- 3rd Bottom Position PM 1 ----- //
            var cell3rdBottomDryer1Past = row3rdBottomDryer.Cells[1];
            cell3rdBottomDryer1Past.Format.Font.Size = 8;
            cell3rdBottomDryer1Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(1, 12);
            if (pastAverage > 0)
                cell3rdBottomDryer1Past.AddParagraph($"{pastAverage}");

            var cell3rdBottomDryer1Current = row3rdBottomDryer.Cells[2];
            cell3rdBottomDryer1Current.Format.Font.Size = 8;
            cell3rdBottomDryer1Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothing3rdBottomDryer1);
            positionGoal = GetPositionGoal(1, 12);
            if (currentAge > 0)
                cell3rdBottomDryer1Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cell3rdBottomDryer1Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cell3rdBottomDryer1Current.Shading.Color = Colors.Yellow;
            }
            

            var cell3rdBottomDryer1Goal = row3rdBottomDryer.Cells[3];
            cell3rdBottomDryer1Goal.Format.Font.Size = 8;
            cell3rdBottomDryer1Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cell3rdBottomDryer1Goal.AddParagraph($"{positionGoal}");

            var cell3rdBottomDryer1Info = row3rdBottomDryer2.Cells[1];
            cell3rdBottomDryer1Info.Format.Font.Size = 8;
            cell3rdBottomDryer1Info.MergeRight = 2;
            cell3rdBottomDryer1Info.VerticalAlignment = VerticalAlignment.Center;
            cell3rdBottomDryer1Info.Shading.Color = Colors.LightBlue;
            cell3rdBottomDryer1Info.AddParagraph($"{GetClothingInfo(clothing3rdBottomDryer1)}");
            #endregion

            #region 3rd Bottom PM 2
            // ----- 3rd Bottom Position PM 2 ----- //
            var cell3rdBottomDryer2Past = row3rdBottomDryer.Cells[4];
            cell3rdBottomDryer2Past.Format.Font.Size = 8;
            cell3rdBottomDryer2Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(2, 12);
            if (pastAverage > 0)
                cell3rdBottomDryer2Past.AddParagraph($"{pastAverage}");

            var cell3rdBottomDryer2Current = row3rdBottomDryer.Cells[5];
            cell3rdBottomDryer2Current.Format.Font.Size = 8;
            cell3rdBottomDryer2Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothing3rdBottomDryer2);
            positionGoal = GetPositionGoal(2, 12);
            if (currentAge > 0)
                cell3rdBottomDryer2Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cell3rdBottomDryer2Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cell3rdBottomDryer2Current.Shading.Color = Colors.Yellow;
            }
            

            var cell3rdBottomDryer2Goal = row3rdBottomDryer.Cells[6];
            cell3rdBottomDryer2Goal.Format.Font.Size = 8;
            cell3rdBottomDryer2Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cell3rdBottomDryer2Goal.AddParagraph($"{positionGoal}");

            var cell3rdBottomDryer2Info = row3rdBottomDryer2.Cells[4];
            cell3rdBottomDryer2Info.Format.Font.Size = 8;
            cell3rdBottomDryer2Info.MergeRight = 2;
            cell3rdBottomDryer2Info.VerticalAlignment = VerticalAlignment.Center;
            cell3rdBottomDryer2Info.Shading.Color = Colors.LightBlue;
            cell3rdBottomDryer2Info.AddParagraph($"{GetClothingInfo(clothing3rdBottomDryer2)}");
            #endregion

            #region 3rd Bottom PM 3
            // ----- 3rd Bottom Position PM 3 ----- //
            var cell3rdBottomDryer3Past = row3rdBottomDryer.Cells[7];
            cell3rdBottomDryer3Past.Format.Font.Size = 8;
            cell3rdBottomDryer3Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(3, 12);
            if (pastAverage > 0)
                cell3rdBottomDryer3Past.AddParagraph($"{pastAverage}");

            var cell3rdBottomDryer3Current = row3rdBottomDryer.Cells[8];
            cell3rdBottomDryer3Current.Format.Font.Size = 8;
            cell3rdBottomDryer3Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothing3rdBottomDryer3);
            positionGoal = GetPositionGoal(3, 12);
            if (currentAge > 0)
                cell3rdBottomDryer3Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cell3rdBottomDryer3Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cell3rdBottomDryer3Current.Shading.Color = Colors.Yellow;
            }
            

            var cell3rdBottomDryer3Goal = row3rdBottomDryer.Cells[9];
            cell3rdBottomDryer3Goal.Format.Font.Size = 8;
            cell3rdBottomDryer3Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cell3rdBottomDryer3Goal.AddParagraph($"{positionGoal}");

            var cell3rdBottomDryer3Info = row3rdBottomDryer2.Cells[7];
            cell3rdBottomDryer3Info.Format.Font.Size = 8;
            cell3rdBottomDryer3Info.MergeRight = 2;
            cell3rdBottomDryer3Info.VerticalAlignment = VerticalAlignment.Center;
            cell3rdBottomDryer3Info.Shading.Color = Colors.LightBlue;
            cell3rdBottomDryer3Info.AddParagraph($"{GetClothingInfo(clothing3rdBottomDryer3)}");
            #endregion

            #region 3rd Bottom PM 4
            // ----- 3rd Bottom Position PM 4 ----- //
            var cell3rdBottomDryer4Past = row3rdBottomDryer.Cells[10];
            cell3rdBottomDryer4Past.Format.Font.Size = 8;
            cell3rdBottomDryer4Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(4, 12);
            if (pastAverage > 0)
                cell3rdBottomDryer4Past.AddParagraph($"{pastAverage}");

            var cell3rdBottomDryer4Current = row3rdBottomDryer.Cells[11];
            cell3rdBottomDryer4Current.Format.Font.Size = 8;
            cell3rdBottomDryer4Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothing3rdBottomDryer4);
            positionGoal = GetPositionGoal(4, 12);
            if (currentAge > 0)
                cell3rdBottomDryer4Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cell3rdBottomDryer4Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cell3rdBottomDryer4Current.Shading.Color = Colors.Yellow;
            }
            

            var cell3rdBottomDryer4Goal = row3rdBottomDryer.Cells[12];
            cell3rdBottomDryer4Goal.Format.Font.Size = 8;
            cell3rdBottomDryer4Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cell3rdBottomDryer4Goal.AddParagraph($"{positionGoal}");

            var cell3rdBottomDryer4Info = row3rdBottomDryer2.Cells[10];
            cell3rdBottomDryer4Info.Format.Font.Size = 8;
            cell3rdBottomDryer4Info.MergeRight = 2;
            cell3rdBottomDryer4Info.VerticalAlignment = VerticalAlignment.Center;
            cell3rdBottomDryer4Info.Shading.Color = Colors.LightBlue;
            cell3rdBottomDryer4Info.AddParagraph($"{GetClothingInfo(clothing3rdBottomDryer4)}");
            #endregion

            #endregion // 3rd Bottom Dryer Position

            #region 4th Bottom Dryer Position
            // add spacer row
            spacer = weeklyPMTable.AddRow();
            spacer.Height = 12;
            cellSpacer = spacer.Cells[0];
            cellSpacer.MergeRight = 12;
            cellSpacer.AddParagraph("");

            var clothing4thBottomDryer1 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "4th bottom dryer felt" && c.PM_Number == 1);
            var clothing4thBottomDryer2 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "4th bottom dryer felt" && c.PM_Number == 2);
            var clothing4thBottomDryer3 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "4th bottom dryer felt" && c.PM_Number == 3);
            var clothing4thBottomDryer4 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "4th bottom dryer felt" && c.PM_Number == 4);
            var row4thBottomDryer = weeklyPMTable.AddRow();
            row4thBottomDryer.Height = 10;
            var row4thBottomDryer2 = weeklyPMTable.AddRow();
            row4thBottomDryer2.Height = 10;

            var cell4thBottomDryer = row4thBottomDryer.Cells[0];
            cell4thBottomDryer.Format.Font.Size = 8;
            cell4thBottomDryer.VerticalAlignment = VerticalAlignment.Center;
            cell4thBottomDryer.MergeDown = 1;
            cell4thBottomDryer.AddParagraph("4TH BOTTOM");

            #region 4th Bottom PM 1
            // ----- 4th Bottom Position PM 1 ----- //
            var cell4thBottomDryer1Past = row4thBottomDryer.Cells[1];
            cell4thBottomDryer1Past.Format.Font.Size = 8;
            cell4thBottomDryer1Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(1, 13);
            if (pastAverage > 0)
                cell4thBottomDryer1Past.AddParagraph($"{pastAverage}");

            var cell4thBottomDryer1Current = row4thBottomDryer.Cells[2];
            cell4thBottomDryer1Current.Format.Font.Size = 8;
            cell4thBottomDryer1Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothing4thBottomDryer1);
            positionGoal = GetPositionGoal(1, 13);
            if (currentAge > 0)
                cell4thBottomDryer1Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cell4thBottomDryer1Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cell4thBottomDryer1Current.Shading.Color = Colors.Yellow;
            }
            

            var cell4thBottomDryer1Goal = row4thBottomDryer.Cells[3];
            cell4thBottomDryer1Goal.Format.Font.Size = 8;
            cell4thBottomDryer1Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cell4thBottomDryer1Goal.AddParagraph($"{positionGoal}");

            var cell4thBottomDryer1Info = row4thBottomDryer2.Cells[1];
            cell4thBottomDryer1Info.Format.Font.Size = 8;
            cell4thBottomDryer1Info.MergeRight = 2;
            cell4thBottomDryer1Info.VerticalAlignment = VerticalAlignment.Center;
            cell4thBottomDryer1Info.Shading.Color = Colors.LightBlue;
            cell4thBottomDryer1Info.AddParagraph($"{GetClothingInfo(clothing4thBottomDryer1)}");
            #endregion

            #region 4th Bottom PM 2
            // ----- 4th Bottom Position PM 2 ----- //
            var cell4thBottomDryer2Past = row4thBottomDryer.Cells[4];
            cell4thBottomDryer2Past.Format.Font.Size = 8;
            cell4thBottomDryer2Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(2, 13);
            if (pastAverage > 0)
                cell4thBottomDryer2Past.AddParagraph($"{pastAverage}");

            var cell4thBottomDryer2Current = row4thBottomDryer.Cells[5];
            cell4thBottomDryer2Current.Format.Font.Size = 8;
            cell4thBottomDryer2Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothing4thBottomDryer2);
            positionGoal = GetPositionGoal(2, 13);
            if (currentAge > 0)
                cell4thBottomDryer2Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cell4thBottomDryer2Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cell4thBottomDryer2Current.Shading.Color = Colors.Yellow;
            }
            

            var cell4thBottomDryer2Goal = row4thBottomDryer.Cells[6];
            cell4thBottomDryer2Goal.Format.Font.Size = 8;
            cell4thBottomDryer2Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cell4thBottomDryer2Goal.AddParagraph($"{positionGoal}");

            var cell4thBottomDryer2Info = row4thBottomDryer2.Cells[4];
            cell4thBottomDryer2Info.Format.Font.Size = 8;
            cell4thBottomDryer2Info.MergeRight = 2;
            cell4thBottomDryer2Info.VerticalAlignment = VerticalAlignment.Center;
            cell4thBottomDryer2Info.Shading.Color = Colors.LightBlue;
            cell4thBottomDryer2Info.AddParagraph($"{GetClothingInfo(clothing4thBottomDryer2)}");
            #endregion

            #region 4th Bottom PM 3
            // ----- 4th Bottom Position PM 3 ----- //
            var cell4thBottomDryer3Past = row4thBottomDryer.Cells[7];
            cell4thBottomDryer3Past.Format.Font.Size = 8;
            cell4thBottomDryer3Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(3, 13);
            if (pastAverage > 0)
                cell4thBottomDryer3Past.AddParagraph($"{pastAverage}");

            var cell4thBottomDryer3Current = row4thBottomDryer.Cells[8];
            cell4thBottomDryer3Current.Format.Font.Size = 8;
            cell4thBottomDryer3Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothing4thBottomDryer3);
            positionGoal = GetPositionGoal(3, 13);
            if (currentAge > 0)
                cell4thBottomDryer3Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cell4thBottomDryer3Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cell4thBottomDryer3Current.Shading.Color = Colors.Yellow;
            }
            

            var cell4thBottomDryer3Goal = row4thBottomDryer.Cells[9];
            cell4thBottomDryer3Goal.Format.Font.Size = 8;
            cell4thBottomDryer3Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cell4thBottomDryer3Goal.AddParagraph($"{positionGoal}");

            var cell4thBottomDryer3Info = row4thBottomDryer2.Cells[7];
            cell4thBottomDryer3Info.Format.Font.Size = 8;
            cell4thBottomDryer3Info.MergeRight = 2;
            cell4thBottomDryer3Info.VerticalAlignment = VerticalAlignment.Center;
            cell4thBottomDryer3Info.Shading.Color = Colors.LightBlue;
            cell4thBottomDryer3Info.AddParagraph($"{GetClothingInfo(clothing4thBottomDryer3)}");
            #endregion

            #region 4th Bottom PM 4
            // ----- 4th Bottom Position PM 4 ----- //
            var cell4thBottomDryer4Past = row4thBottomDryer.Cells[10];
            cell4thBottomDryer4Past.Format.Font.Size = 8;
            cell4thBottomDryer4Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(4, 13);
            if (pastAverage > 0)
                cell4thBottomDryer4Past.AddParagraph($"{pastAverage}");

            var cell4thBottomDryer4Current = row4thBottomDryer.Cells[11];
            cell4thBottomDryer4Current.Format.Font.Size = 8;
            cell4thBottomDryer4Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothing4thBottomDryer4);
            positionGoal = GetPositionGoal(4, 13);
            if (currentAge > 0)
                cell4thBottomDryer4Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cell4thBottomDryer4Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cell4thBottomDryer4Current.Shading.Color = Colors.Yellow;
            }
            

            var cell4thBottomDryer4Goal = row4thBottomDryer.Cells[12];
            cell4thBottomDryer4Goal.Format.Font.Size = 8;
            cell4thBottomDryer4Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cell4thBottomDryer4Goal.AddParagraph($"{positionGoal}");

            var cell4thBottomDryer4Info = row4thBottomDryer2.Cells[10];
            cell4thBottomDryer4Info.Format.Font.Size = 8;
            cell4thBottomDryer4Info.MergeRight = 2;
            cell4thBottomDryer4Info.VerticalAlignment = VerticalAlignment.Center;
            cell4thBottomDryer4Info.Shading.Color = Colors.LightBlue;
            cell4thBottomDryer4Info.AddParagraph($"{GetClothingInfo(clothing4thBottomDryer4)}");
            #endregion

            #endregion // 4th Bottom Dryer Position

            return weeklyPMTable;
        }

        internal static Table BuildWeeklyPMRollsTable(List<Clothing> clothings)
        {
            var weeklyPMTable = new Table
            {
                Style = "Table"
            };
            weeklyPMTable.Borders.Width = .075;
            weeklyPMTable.Rows.Alignment = RowAlignment.Left;
            weeklyPMTable.Rows.LeftIndent = 20;

            #region Header Columns
            //define header columns
            var posCol = weeklyPMTable.AddColumn(Unit.FromCentimeter(3.5));
            posCol.Format.Alignment = ParagraphAlignment.Center;
            posCol.Format.Font.Size = 10;

            var pm1Col = weeklyPMTable.AddColumn(Unit.FromCentimeter(2));
            pm1Col.Format.Alignment = ParagraphAlignment.Center;
            pm1Col.Format.Font.Size = 10;

            var pm1Colb = weeklyPMTable.AddColumn(Unit.FromCentimeter(2));
            pm1Colb.Format.Alignment = ParagraphAlignment.Center;
            pm1Colb.Format.Font.Size = 10;

            var pm1Colc = weeklyPMTable.AddColumn(Unit.FromCentimeter(2));
            pm1Colc.Format.Alignment = ParagraphAlignment.Center;
            pm1Colc.Format.Font.Size = 10;

            var pm2col = weeklyPMTable.AddColumn(Unit.FromCentimeter(2));
            pm2col.Format.Alignment = ParagraphAlignment.Center;
            pm2col.Format.Font.Size = 10;

            var pm2colb = weeklyPMTable.AddColumn(Unit.FromCentimeter(2));
            pm2colb.Format.Alignment = ParagraphAlignment.Center;
            pm2colb.Format.Font.Size = 10;

            var pm2colc = weeklyPMTable.AddColumn(Unit.FromCentimeter(2));
            pm2colc.Format.Alignment = ParagraphAlignment.Center;
            pm2colc.Format.Font.Size = 10;

            var pm3col = weeklyPMTable.AddColumn(Unit.FromCentimeter(2));
            pm3col.Format.Alignment = ParagraphAlignment.Center;
            pm3col.Format.Font.Size = 10;

            var pm3colb = weeklyPMTable.AddColumn(Unit.FromCentimeter(2));
            pm3colb.Format.Alignment = ParagraphAlignment.Center;
            pm3colb.Format.Font.Size = 10;

            var pm3colc = weeklyPMTable.AddColumn(Unit.FromCentimeter(2));
            pm3colc.Format.Alignment = ParagraphAlignment.Center;
            pm3colc.Format.Font.Size = 10;

            var pm4col = weeklyPMTable.AddColumn(Unit.FromCentimeter(2));
            pm4col.Format.Alignment = ParagraphAlignment.Center;
            pm4col.Format.Font.Size = 10;

            var pm4colb = weeklyPMTable.AddColumn(Unit.FromCentimeter(2));
            pm4colb.Format.Alignment = ParagraphAlignment.Center;
            pm4colb.Format.Font.Size = 10;

            var pm4colc = weeklyPMTable.AddColumn(Unit.FromCentimeter(2));
            pm4colc.Format.Alignment = ParagraphAlignment.Center;
            pm4colc.Format.Font.Size = 10;

            #endregion

            #region Header Rows

            //define headers row
            var headers = weeklyPMTable.AddRow();
            headers.Shading.Color = Colors.LightGray;
            headers.Height = 15;

            //define header row cells
            var posHeader = headers.Cells[0];
            posHeader.Format.Font.Size = 10;
            posHeader.VerticalAlignment = VerticalAlignment.Center;
            posHeader.MergeDown = 1;
            posHeader.AddParagraph("Position");

            var pm1Header = headers.Cells[1];
            pm1Header.Format.Font.Size = 10;
            pm1Header.MergeRight = 2;
            pm1Header.VerticalAlignment = VerticalAlignment.Center;
            pm1Header.AddParagraph("#1 PM Days Run");

            var pm2Header = headers.Cells[4];
            pm2Header.Format.Font.Size = 10;
            pm2Header.MergeRight = 2;
            pm2Header.VerticalAlignment = VerticalAlignment.Center;
            pm2Header.AddParagraph("#2 PM Days Run");

            var pm3Header = headers.Cells[7];
            pm3Header.Format.Font.Size = 10;
            pm3Header.MergeRight = 2;
            pm3Header.VerticalAlignment = VerticalAlignment.Center;
            pm3Header.AddParagraph("#3 PM Days Run");

            var pm4Header = headers.Cells[10];
            pm4Header.Format.Font.Size = 10;
            pm4Header.MergeRight = 2;
            pm4Header.VerticalAlignment = VerticalAlignment.Center;
            pm4Header.AddParagraph("#4 PM Days Run");

            //define header row 2
            var headers2 = weeklyPMTable.AddRow();
            headers2.Height = 10;

            //define 2nd header row cells
            var pm1PastHeader = headers2.Cells[1];
            pm1PastHeader.Format.Font.Size = 8;
            pm1PastHeader.VerticalAlignment = VerticalAlignment.Center;
            pm1PastHeader.AddParagraph("PAST");

            var pm1CurrentHeader = headers2.Cells[2];
            pm1CurrentHeader.Format.Font.Size = 8;
            pm1CurrentHeader.VerticalAlignment = VerticalAlignment.Center;
            pm1CurrentHeader.AddParagraph("CURRENT");

            var pm1GoalHeader = headers2.Cells[3];
            pm1GoalHeader.Format.Font.Size = 8;
            pm1GoalHeader.VerticalAlignment = VerticalAlignment.Center;
            pm1GoalHeader.AddParagraph("GOAL");

            var pm2PastHeader = headers2.Cells[4];
            pm2PastHeader.Format.Font.Size = 8;
            pm2PastHeader.VerticalAlignment = VerticalAlignment.Center;
            pm2PastHeader.AddParagraph("PAST");

            var pm2CurrentHeader = headers2.Cells[5];
            pm2CurrentHeader.Format.Font.Size = 8;
            pm2CurrentHeader.VerticalAlignment = VerticalAlignment.Center;
            pm2CurrentHeader.AddParagraph("CURRENT");

            var pm2GoalHeader = headers2.Cells[6];
            pm2GoalHeader.Format.Font.Size = 8;
            pm2GoalHeader.VerticalAlignment = VerticalAlignment.Center;
            pm2GoalHeader.AddParagraph("GOAL");

            var pm3PastHeader = headers2.Cells[7];
            pm3PastHeader.Format.Font.Size = 8;
            pm3PastHeader.VerticalAlignment = VerticalAlignment.Center;
            pm3PastHeader.AddParagraph("PAST");

            var pm3CurrentHeader = headers2.Cells[8];
            pm3CurrentHeader.Format.Font.Size = 8;
            pm3CurrentHeader.VerticalAlignment = VerticalAlignment.Center;
            pm3CurrentHeader.AddParagraph("CURRENT");

            var pm3GoalHeader = headers2.Cells[9];
            pm3GoalHeader.Format.Font.Size = 8;
            pm3GoalHeader.VerticalAlignment = VerticalAlignment.Center;
            pm3GoalHeader.AddParagraph("GOAL");

            var pm4PastHeader = headers2.Cells[10];
            pm4PastHeader.Format.Font.Size = 8;
            pm4PastHeader.VerticalAlignment = VerticalAlignment.Center;
            pm4PastHeader.AddParagraph("PAST");

            var pm4CurrentHeader = headers2.Cells[11];
            pm4CurrentHeader.Format.Font.Size = 8;
            pm4CurrentHeader.VerticalAlignment = VerticalAlignment.Center;
            pm4CurrentHeader.AddParagraph("CURRENT");

            var pm4GoalHeader = headers2.Cells[12];
            pm4GoalHeader.Format.Font.Size = 8;
            pm4GoalHeader.VerticalAlignment = VerticalAlignment.Center;
            pm4GoalHeader.AddParagraph("GOAL");
            #endregion

            //Rolls start here

            #region Breast Roll Position
            // add spacer row
            var spacer = weeklyPMTable.AddRow();
            spacer.Height = 12;
            var cellSpacer = spacer.Cells[0];
            cellSpacer.MergeRight = 12;
            cellSpacer.AddParagraph("");

            var clothingBreast1 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "breast roll" && c.PM_Number == 1);
            var clothingBreast2 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "breast roll" && c.PM_Number == 2);
            var clothingBreast3 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "breast roll" && c.PM_Number == 3);
            var clothingBreast4 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "breast roll" && c.PM_Number == 4);
            var rowBreast = weeklyPMTable.AddRow();
            rowBreast.Height = 10;
            var rowBreast2 = weeklyPMTable.AddRow();
            rowBreast2.Height = 10;

            var cellBreast = rowBreast.Cells[0];
            cellBreast.Format.Font.Size = 8;
            cellBreast.VerticalAlignment = VerticalAlignment.Center;
            cellBreast.MergeDown = 1;
            cellBreast.AddParagraph("BREAST ROLL");

            #region Breast PM 1
            // ----- Breast Position PM 1 ----- //
            var cellBreast1Past = rowBreast.Cells[1];
            cellBreast1Past.Format.Font.Size = 8;
            cellBreast1Past.VerticalAlignment = VerticalAlignment.Center;
            var pastAverage = GetPastAverage(1, 14);
            if (pastAverage > 0)
                cellBreast1Past.AddParagraph($"{pastAverage}");

            var cellBreast1Current = rowBreast.Cells[2];
            cellBreast1Current.Format.Font.Size = 8;
            cellBreast1Current.VerticalAlignment = VerticalAlignment.Center;
            var currentAge = GetCurrentAge(clothingBreast1);
            var positionGoal = GetPositionGoal(1, 14);
            if (currentAge > 0)
                cellBreast1Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cellBreast1Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cellBreast1Current.Shading.Color = Colors.Yellow;
            }


            var cellBreast1Goal = rowBreast.Cells[3];
            cellBreast1Goal.Format.Font.Size = 8;
            cellBreast1Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cellBreast1Goal.AddParagraph($"{positionGoal}");

            var cellBreast1Info = rowBreast2.Cells[1];
            cellBreast1Info.Format.Font.Size = 8;
            cellBreast1Info.MergeRight = 2;
            cellBreast1Info.VerticalAlignment = VerticalAlignment.Center;
            cellBreast1Info.Shading.Color = Colors.LightBlue;
            cellBreast1Info.AddParagraph($"{GetClothingInfo(clothingBreast1)}");
            #endregion

            #region Breast PM 2
            // ----- Breast Position PM 2 ----- //
            var cellBreast2Past = rowBreast.Cells[4];
            cellBreast2Past.Format.Font.Size = 8;
            cellBreast2Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(2, 14);
            if (pastAverage > 0)
                cellBreast2Past.AddParagraph($"{pastAverage}");

            var cellBreast2Current = rowBreast.Cells[5];
            cellBreast2Current.Format.Font.Size = 8;
            cellBreast2Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothingBreast2);
            positionGoal = GetPositionGoal(2, 14);
            if (currentAge > 0)
                cellBreast2Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cellBreast2Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cellBreast2Current.Shading.Color = Colors.Yellow;
            }


            var cellBreast2Goal = rowBreast.Cells[6];
            cellBreast2Goal.Format.Font.Size = 8;
            cellBreast2Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cellBreast2Goal.AddParagraph($"{positionGoal}");

            var cellBreast2Info = rowBreast2.Cells[4];
            cellBreast2Info.Format.Font.Size = 8;
            cellBreast2Info.MergeRight = 2;
            cellBreast2Info.VerticalAlignment = VerticalAlignment.Center;
            cellBreast2Info.Shading.Color = Colors.LightBlue;
            cellBreast2Info.AddParagraph($"{GetClothingInfo(clothingBreast2)}");
            #endregion

            #region Breast PM 3
            // ----- Breast Position PM 3 ----- //
            var cellBreast3Past = rowBreast.Cells[7];
            cellBreast3Past.Format.Font.Size = 8;
            cellBreast3Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(3, 13);
            if (pastAverage > 0)
                cellBreast3Past.AddParagraph($"{pastAverage}");

            var cellBreast3Current = rowBreast.Cells[8];
            cellBreast3Current.Format.Font.Size = 8;
            cellBreast3Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothingBreast3);
            positionGoal = GetPositionGoal(3, 14);
            if (currentAge > 0)
                cellBreast3Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cellBreast3Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cellBreast3Current.Shading.Color = Colors.Yellow;
            }


            var cellBreast3Goal = rowBreast.Cells[9];
            cellBreast3Goal.Format.Font.Size = 8;
            cellBreast3Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cellBreast3Goal.AddParagraph($"{positionGoal}");

            var cellBreast3Info = rowBreast2.Cells[7];
            cellBreast3Info.Format.Font.Size = 8;
            cellBreast3Info.MergeRight = 2;
            cellBreast3Info.VerticalAlignment = VerticalAlignment.Center;
            cellBreast3Info.Shading.Color = Colors.LightBlue;
            cellBreast3Info.AddParagraph($"{GetClothingInfo(clothingBreast3)}");
            #endregion

            #region Breast PM 4
            // ----- Breast Position PM 4 ----- //
            var cellBreast4Past = rowBreast.Cells[10];
            cellBreast4Past.Format.Font.Size = 8;
            cellBreast4Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(4, 14);
            if (pastAverage > 0)
                cellBreast4Past.AddParagraph($"{pastAverage}");

            var cellBreast4Current = rowBreast.Cells[11];
            cellBreast4Current.Format.Font.Size = 8;
            cellBreast4Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothingBreast4);
            positionGoal = GetPositionGoal(4, 14);
            if (currentAge > 0)
                cellBreast4Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cellBreast4Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cellBreast4Current.Shading.Color = Colors.Yellow;
            }


            var cellBreast4Goal = rowBreast.Cells[12];
            cellBreast4Goal.Format.Font.Size = 8;
            cellBreast4Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cellBreast4Goal.AddParagraph($"{positionGoal}");

            var cellBreast4Info = rowBreast2.Cells[10];
            cellBreast4Info.Format.Font.Size = 8;
            cellBreast4Info.MergeRight = 2;
            cellBreast4Info.VerticalAlignment = VerticalAlignment.Center;
            cellBreast4Info.Shading.Color = Colors.LightBlue;
            cellBreast4Info.AddParagraph($"{GetClothingInfo(clothingBreast4)}");
            #endregion

            #endregion // Breast Roll Position

            #region Dandy Position
            // add spacer row
            spacer = weeklyPMTable.AddRow();
            spacer.Height = 12;
            cellSpacer = spacer.Cells[0];
            cellSpacer = spacer.Cells[0];
            cellSpacer.MergeRight = 12;
            cellSpacer.AddParagraph("");

            var clothingDandy1 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "dandy roll" && c.PM_Number == 1);
            var clothingDandy2 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "dandy roll" && c.PM_Number == 2);
            var clothingDandy3 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "dandy roll" && c.PM_Number == 3);
            var clothingDandy4 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "dandy roll" && c.PM_Number == 4);
            var rowDandy = weeklyPMTable.AddRow();
            rowDandy.Height = 10;
            var rowDandy2 = weeklyPMTable.AddRow();
            rowDandy2.Height = 10;

            var cellDandy = rowDandy.Cells[0];
            cellDandy.Format.Font.Size = 8;
            cellDandy.VerticalAlignment = VerticalAlignment.Center;
            cellDandy.MergeDown = 1;
            cellDandy.AddParagraph("DANDY");

            #region Dandy PM 1
            // ----- Dandy Position PM 1 ----- //
            var cellDandy1Past = rowDandy.Cells[1];
            cellDandy1Past.Format.Font.Size = 8;
            cellDandy1Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(1, 15);
            if (pastAverage > 0)
                cellDandy1Past.AddParagraph($"{pastAverage}");

            var cellDandy1Current = rowDandy.Cells[2];
            cellDandy1Current.Format.Font.Size = 8;
            cellDandy1Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothingDandy1);
            positionGoal = GetPositionGoal(1, 15);
            if (currentAge > 0)
                cellDandy1Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cellDandy1Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cellDandy1Current.Shading.Color = Colors.Yellow;
            }


            var cellDandy1Goal = rowDandy.Cells[3];
            cellDandy1Goal.Format.Font.Size = 8;
            cellDandy1Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cellDandy1Goal.AddParagraph($"{positionGoal}");

            var cellDandy1Info = rowDandy2.Cells[1];
            cellDandy1Info.Format.Font.Size = 8;
            cellDandy1Info.MergeRight = 2;
            cellDandy1Info.VerticalAlignment = VerticalAlignment.Center;
            cellDandy1Info.Shading.Color = Colors.LightBlue;
            cellDandy1Info.AddParagraph($"{GetClothingInfo(clothingDandy1)}");
            #endregion

            #region Dandy PM 2
            // ----- Dandy Position PM 2 ----- //
            var cellDandy2Past = rowDandy.Cells[4];
            cellDandy2Past.Format.Font.Size = 8;
            cellDandy2Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(2, 15);
            if (pastAverage > 0)
                cellDandy2Past.AddParagraph($"{pastAverage}");

            var cellDandy2Current = rowDandy.Cells[5];
            cellDandy2Current.Format.Font.Size = 8;
            cellDandy2Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothingDandy2);
            positionGoal = GetPositionGoal(2, 15);
            if (currentAge > 0)
                cellDandy2Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cellDandy2Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cellDandy2Current.Shading.Color = Colors.Yellow;
            }


            var cellDandy2Goal = rowDandy.Cells[6];
            cellDandy2Goal.Format.Font.Size = 8;
            cellDandy2Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cellDandy2Goal.AddParagraph($"{positionGoal}");

            var cellDandy2Info = rowDandy2.Cells[4];
            cellDandy2Info.Format.Font.Size = 8;
            cellDandy2Info.MergeRight = 2;
            cellDandy2Info.VerticalAlignment = VerticalAlignment.Center;
            cellDandy2Info.Shading.Color = Colors.LightBlue;
            cellDandy2Info.AddParagraph($"{GetClothingInfo(clothingDandy2)}");
            #endregion

            #region Dandy PM 3
            // ----- Dandy Position PM 3 ----- //
            var cellDandy3Past = rowDandy.Cells[7];
            cellDandy3Past.Format.Font.Size = 8;
            cellDandy3Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(3, 15);
            if (pastAverage > 0)
                cellDandy3Past.AddParagraph($"{pastAverage}");

            var cellDandy3Current = rowDandy.Cells[8];
            cellDandy3Current.Format.Font.Size = 8;
            cellDandy3Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothingDandy3);
            positionGoal = GetPositionGoal(3, 15);
            if (currentAge > 0)
                cellDandy3Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cellDandy3Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cellDandy3Current.Shading.Color = Colors.Yellow;
            }


            var cellDandy3Goal = rowDandy.Cells[9];
            cellDandy3Goal.Format.Font.Size = 8;
            cellDandy3Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cellDandy3Goal.AddParagraph($"{positionGoal}");

            var cellDandy3Info = rowDandy2.Cells[7];
            cellDandy3Info.Format.Font.Size = 8;
            cellDandy3Info.MergeRight = 2;
            cellDandy3Info.VerticalAlignment = VerticalAlignment.Center;
            cellDandy3Info.Shading.Color = Colors.LightBlue;
            cellDandy3Info.AddParagraph($"{GetClothingInfo(clothingDandy3)}");
            #endregion

            #region Dandy PM 4
            // ----- Dandy Position PM 4 ----- //
            var cellDandy4Past = rowDandy.Cells[10];
            cellDandy4Past.Format.Font.Size = 8;
            cellDandy4Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(4, 15);
            if (pastAverage > 0)
                cellDandy4Past.AddParagraph($"{pastAverage}");

            var cellDandy4Current = rowDandy.Cells[11];
            cellDandy4Current.Format.Font.Size = 8;
            cellDandy4Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothingDandy4);
            positionGoal = GetPositionGoal(4, 15);
            if (currentAge > 0)
                cellDandy4Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cellDandy4Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cellDandy4Current.Shading.Color = Colors.Yellow;
            }


            var cellDandy4Goal = rowDandy.Cells[12];
            cellDandy4Goal.Format.Font.Size = 8;
            cellDandy4Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cellDandy4Goal.AddParagraph($"{positionGoal}");

            var cellDandy4Info = rowDandy2.Cells[10];
            cellDandy4Info.Format.Font.Size = 8;
            cellDandy4Info.MergeRight = 2;
            cellDandy4Info.VerticalAlignment = VerticalAlignment.Center;
            cellDandy4Info.Shading.Color = Colors.LightBlue;
            cellDandy4Info.AddParagraph($"{GetClothingInfo(clothingDandy4)}");
            #endregion

            #endregion // Dandy Roll Position

            #region Lumpbreaker Position

            var clothingLumpbreaker1 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "lumpbreaker roll" && c.PM_Number == 1);
            var clothingLumpbreaker2 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "lumpbreaker roll" && c.PM_Number == 2);
            var clothingLumpbreaker3 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "lumpbreaker roll" && c.PM_Number == 3);
            var clothingLumpbreaker4 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "lumpbreaker roll" && c.PM_Number == 4);
            var rowLumpbreaker = weeklyPMTable.AddRow();
            rowLumpbreaker.Height = 10;
            var rowLumpbreaker2 = weeklyPMTable.AddRow();
            rowLumpbreaker2.Height = 10;

            var cellLumpbreaker = rowLumpbreaker.Cells[0];
            cellLumpbreaker.Format.Font.Size = 8;
            cellLumpbreaker.VerticalAlignment = VerticalAlignment.Center;
            cellLumpbreaker.MergeDown = 1;
            cellLumpbreaker.AddParagraph("LUMPBREAKER");

            #region Lumpbreaker PM 1
            // ----- Lumpbreaker Position PM 1 ----- //
            var cellLumpbreaker1Past = rowLumpbreaker.Cells[1];
            cellLumpbreaker1Past.Format.Font.Size = 8;
            cellLumpbreaker1Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(1, 16);
            if (pastAverage > 0)
                cellLumpbreaker1Past.AddParagraph($"{pastAverage}");

            var cellLumpbreaker1Current = rowLumpbreaker.Cells[2];
            cellLumpbreaker1Current.Format.Font.Size = 8;
            cellLumpbreaker1Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothingLumpbreaker1);
            positionGoal = GetPositionGoal(1, 16);
            if (currentAge > 0)
                cellLumpbreaker1Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cellLumpbreaker1Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cellLumpbreaker1Current.Shading.Color = Colors.Yellow;
            }


            var cellLumpbreaker1Goal = rowLumpbreaker.Cells[3];
            cellLumpbreaker1Goal.Format.Font.Size = 8;
            cellLumpbreaker1Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cellLumpbreaker1Goal.AddParagraph($"{positionGoal}");

            var cellLumpbreaker1Info = rowLumpbreaker2.Cells[1];
            cellLumpbreaker1Info.Format.Font.Size = 8;
            cellLumpbreaker1Info.MergeRight = 2;
            cellLumpbreaker1Info.VerticalAlignment = VerticalAlignment.Center;
            cellLumpbreaker1Info.Shading.Color = Colors.LightBlue;
            cellLumpbreaker1Info.AddParagraph($"{GetClothingInfo(clothingLumpbreaker1)}");
            #endregion

            #region Lumpbreaker PM 2
            // ----- Lumpbreaker Position PM 2 ----- //
            var cellLumpbreaker2Past = rowLumpbreaker.Cells[4];
            cellLumpbreaker2Past.Format.Font.Size = 8;
            cellLumpbreaker2Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(2, 16);
            if (pastAverage > 0)
                cellLumpbreaker2Past.AddParagraph($"{pastAverage}");

            var cellLumpbreaker2Current = rowLumpbreaker.Cells[5];
            cellLumpbreaker2Current.Format.Font.Size = 8;
            cellLumpbreaker2Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothingLumpbreaker2);
            positionGoal = GetPositionGoal(2, 16);
            if (currentAge > 0)
                cellLumpbreaker2Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cellLumpbreaker2Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cellLumpbreaker2Current.Shading.Color = Colors.Yellow;
            }


            var cellLumpbreaker2Goal = rowLumpbreaker.Cells[6];
            cellLumpbreaker2Goal.Format.Font.Size = 8;
            cellLumpbreaker2Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cellLumpbreaker2Goal.AddParagraph($"{positionGoal}");

            var cellLumpbreaker2Info = rowLumpbreaker2.Cells[4];
            cellLumpbreaker2Info.Format.Font.Size = 8;
            cellLumpbreaker2Info.MergeRight = 2;
            cellLumpbreaker2Info.VerticalAlignment = VerticalAlignment.Center;
            cellLumpbreaker2Info.Shading.Color = Colors.LightBlue;
            cellLumpbreaker2Info.AddParagraph($"{GetClothingInfo(clothingLumpbreaker2)}");
            #endregion

            #region Lumpbreaker PM 3
            // ----- Lumpbreaker Position PM 3 ----- //
            var cellLumpbreaker3Past = rowLumpbreaker.Cells[7];
            cellLumpbreaker3Past.Format.Font.Size = 8;
            cellLumpbreaker3Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(3, 16);
            if (pastAverage > 0)
                cellLumpbreaker3Past.AddParagraph($"{pastAverage}");

            var cellLumpbreaker3Current = rowLumpbreaker.Cells[8];
            cellLumpbreaker3Current.Format.Font.Size = 8;
            cellLumpbreaker3Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothingLumpbreaker3);
            positionGoal = GetPositionGoal(3, 16);
            if (currentAge > 0)
                cellLumpbreaker3Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cellLumpbreaker3Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cellLumpbreaker3Current.Shading.Color = Colors.Yellow;
            }


            var cellLumpbreaker3Goal = rowLumpbreaker.Cells[9];
            cellLumpbreaker3Goal.Format.Font.Size = 8;
            cellLumpbreaker3Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cellLumpbreaker3Goal.AddParagraph($"{positionGoal}");

            var cellLumpbreaker3Info = rowLumpbreaker2.Cells[7];
            cellLumpbreaker3Info.Format.Font.Size = 8;
            cellLumpbreaker3Info.MergeRight = 2;
            cellLumpbreaker3Info.VerticalAlignment = VerticalAlignment.Center;
            cellLumpbreaker3Info.Shading.Color = Colors.LightBlue;
            cellLumpbreaker3Info.AddParagraph($"{GetClothingInfo(clothingLumpbreaker3)}");
            #endregion

            #region Lumpbreaker PM 4
            // ----- Lumpbreaker Position PM 4 ----- //
            var cellLumpbreaker4Past = rowLumpbreaker.Cells[10];
            cellLumpbreaker4Past.Format.Font.Size = 8;
            cellLumpbreaker4Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(4, 16);
            if (pastAverage > 0)
                cellLumpbreaker4Past.AddParagraph($"{pastAverage}");

            var cellLumpbreaker4Current = rowLumpbreaker.Cells[11];
            cellLumpbreaker4Current.Format.Font.Size = 8;
            cellLumpbreaker4Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothingLumpbreaker4);
            positionGoal = GetPositionGoal(4, 16);
            if (currentAge > 0)
                cellLumpbreaker4Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cellLumpbreaker4Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cellLumpbreaker4Current.Shading.Color = Colors.Yellow;
            }


            var cellLumpbreaker4Goal = rowLumpbreaker.Cells[12];
            cellLumpbreaker4Goal.Format.Font.Size = 8;
            cellLumpbreaker4Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cellLumpbreaker4Goal.AddParagraph($"{positionGoal}");

            var cellLumpbreaker4Info = rowLumpbreaker2.Cells[10];
            cellLumpbreaker4Info.Format.Font.Size = 8;
            cellLumpbreaker4Info.MergeRight = 2;
            cellLumpbreaker4Info.VerticalAlignment = VerticalAlignment.Center;
            cellLumpbreaker4Info.Shading.Color = Colors.LightBlue;
            cellLumpbreaker4Info.AddParagraph($"{GetClothingInfo(clothingLumpbreaker4)}");
            #endregion

            #endregion // Lumpbreaker Roll Position

            #region Suction Pickup Position
            // add spacer row
            spacer = weeklyPMTable.AddRow();
            spacer.Height = 12;
            cellSpacer = spacer.Cells[0];
            cellSpacer.MergeRight = 12;
            cellSpacer.AddParagraph("");

            var clothingSuction1 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "suction pickup roll" && c.PM_Number == 1);
            var clothingSuction2 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "suction pickup roll" && c.PM_Number == 2);
            var clothingSuction3 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "suction pickup roll" && c.PM_Number == 3);
            var clothingSuction4 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "suction pickup roll" && c.PM_Number == 4);
            var rowSuction = weeklyPMTable.AddRow();
            rowSuction.Height = 10;
            var rowSuction2 = weeklyPMTable.AddRow();
            rowSuction2.Height = 10;

            var cellSuction = rowSuction.Cells[0];
            cellSuction.Format.Font.Size = 8;
            cellSuction.VerticalAlignment = VerticalAlignment.Center;
            cellSuction.MergeDown = 1;
            cellSuction.AddParagraph("SUCTION PICKUP");

            #region Suction Pickup PM 1
            // ----- Suction Position PM 1 ----- //
            var cellSuction1Past = rowSuction.Cells[1];
            cellSuction1Past.Format.Font.Size = 8;
            cellSuction1Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(1, 17);
            if (pastAverage > 0)
                cellSuction1Past.AddParagraph($"{pastAverage}");

            var cellSuction1Current = rowSuction.Cells[2];
            cellSuction1Current.Format.Font.Size = 8;
            cellSuction1Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothingSuction1);
            positionGoal = GetPositionGoal(1, 17);
            if (currentAge > 0)
                cellSuction1Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cellSuction1Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cellSuction1Current.Shading.Color = Colors.Yellow;
            }


            var cellSuction1Goal = rowSuction.Cells[3];
            cellSuction1Goal.Format.Font.Size = 8;
            cellSuction1Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cellSuction1Goal.AddParagraph($"{positionGoal}");

            var cellSuction1Info = rowSuction2.Cells[1];
            cellSuction1Info.Format.Font.Size = 8;
            cellSuction1Info.MergeRight = 2;
            cellSuction1Info.VerticalAlignment = VerticalAlignment.Center;
            cellSuction1Info.Shading.Color = Colors.LightBlue;
            cellSuction1Info.AddParagraph($"{GetClothingInfo(clothingSuction1)}");
            #endregion

            #region Suction Pickup PM 2
            // ----- Suction Position PM 2 ----- //
            var cellSuction2Past = rowSuction.Cells[4];
            cellSuction2Past.Format.Font.Size = 8;
            cellSuction2Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(2, 17);
            if (pastAverage > 0)
                cellSuction2Past.AddParagraph($"{pastAverage}");

            var cellSuction2Current = rowSuction.Cells[5];
            cellSuction2Current.Format.Font.Size = 8;
            cellSuction2Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothingSuction2);
            positionGoal = GetPositionGoal(2, 17);
            if (currentAge > 0)
                cellSuction2Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cellSuction2Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cellSuction2Current.Shading.Color = Colors.Yellow;
            }


            var cellSuction2Goal = rowSuction.Cells[6];
            cellSuction2Goal.Format.Font.Size = 8;
            cellSuction2Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cellSuction2Goal.AddParagraph($"{positionGoal}");

            var cellSuction2Info = rowSuction2.Cells[4];
            cellSuction2Info.Format.Font.Size = 8;
            cellSuction2Info.MergeRight = 2;
            cellSuction2Info.VerticalAlignment = VerticalAlignment.Center;
            cellSuction2Info.Shading.Color = Colors.LightBlue;
            cellSuction2Info.AddParagraph($"{GetClothingInfo(clothingSuction2)}");
            #endregion

            #region  Suction Pickup PM 3
            // ----- Suction Position PM 3 ----- //
            var cellSuction3Past = rowSuction.Cells[7];
            cellSuction3Past.Format.Font.Size = 8;
            cellSuction3Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(3, 17);
            if (pastAverage > 0)
                cellSuction3Past.AddParagraph($"{pastAverage}");

            var cellSuction3Current = rowSuction.Cells[8];
            cellSuction3Current.Format.Font.Size = 8;
            cellSuction3Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothingSuction3);
            positionGoal = GetPositionGoal(3, 17);
            if (currentAge > 0)
                cellSuction3Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cellSuction3Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cellSuction3Current.Shading.Color = Colors.Yellow;
            }


            var cellSuction3Goal = rowSuction.Cells[9];
            cellSuction3Goal.Format.Font.Size = 8;
            cellSuction3Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cellSuction3Goal.AddParagraph($"{positionGoal}");

            var cellSuction3Info = rowSuction2.Cells[7];
            cellSuction3Info.Format.Font.Size = 8;
            cellSuction3Info.MergeRight = 2;
            cellSuction3Info.VerticalAlignment = VerticalAlignment.Center;
            cellSuction3Info.Shading.Color = Colors.LightBlue;
            cellSuction3Info.AddParagraph($"{GetClothingInfo(clothingSuction3)}");
            #endregion

            #region Suction Pickup PM 4
            // ----- Suction Position PM 4 ----- //
            var cellSuction4Past = rowSuction.Cells[10];
            cellSuction4Past.Format.Font.Size = 8;
            cellSuction4Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(4, 17);
            if (pastAverage > 0)
                cellSuction4Past.AddParagraph($"{pastAverage}");

            var cellSuction4Current = rowSuction.Cells[11];
            cellSuction4Current.Format.Font.Size = 8;
            cellSuction4Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothingSuction4);
            positionGoal = GetPositionGoal(4, 17);
            if (currentAge > 0)
                cellSuction4Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cellSuction4Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cellSuction4Current.Shading.Color = Colors.Yellow;
            }


            var cellSuction4Goal = rowSuction.Cells[12];
            cellSuction4Goal.Format.Font.Size = 8;
            cellSuction4Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cellSuction4Goal.AddParagraph($"{positionGoal}");

            var cellSuction4Info = rowSuction2.Cells[10];
            cellSuction4Info.Format.Font.Size = 8;
            cellSuction4Info.MergeRight = 2;
            cellSuction4Info.VerticalAlignment = VerticalAlignment.Center;
            cellSuction4Info.Shading.Color = Colors.LightBlue;
            cellSuction4Info.AddParagraph($"{GetClothingInfo(clothingSuction4)}");
            #endregion

            #endregion // Suction Pickup Position

            #region 1st Press Top Position
            // add spacer row
            spacer = weeklyPMTable.AddRow();
            spacer.Height = 12;
            cellSpacer = spacer.Cells[0];
            cellSpacer.MergeRight = 12;
            cellSpacer.AddParagraph("");

            var clothing1stPressTop1 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "1st press top roll" && c.PM_Number == 1);
            var clothing1stPressTop2 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "1st press top roll" && c.PM_Number == 2);
            var clothing1stPressTop3 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "1st press top roll" && c.PM_Number == 3);
            var clothing1stPressTop4 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "1st press top roll" && c.PM_Number == 4);
            var row1stPressTop = weeklyPMTable.AddRow();
            row1stPressTop.Height = 10;
            var row1stPressTop2 = weeklyPMTable.AddRow();
            row1stPressTop2.Height = 10;

            var cell1stPressTop = row1stPressTop.Cells[0];
            cell1stPressTop.Format.Font.Size = 8;
            cell1stPressTop.VerticalAlignment = VerticalAlignment.Center;
            cell1stPressTop.MergeDown = 1;
            cell1stPressTop.AddParagraph("1ST PRESS TOP");

            #region 1st Press Top PM 1
            // ----- 1st Press Top Position PM 1 ----- //
            var cell1stPressTop1Past = row1stPressTop.Cells[1];
            cell1stPressTop1Past.Format.Font.Size = 8;
            cell1stPressTop1Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(1, 18);
            if (pastAverage > 0)
                cell1stPressTop1Past.AddParagraph($"{pastAverage}");

            var cell1stPressTop1Current = row1stPressTop.Cells[2];
            cell1stPressTop1Current.Format.Font.Size = 8;
            cell1stPressTop1Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothing1stPressTop1);
            positionGoal = GetPositionGoal(1, 18);
            if (currentAge > 0)
                cell1stPressTop1Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cell1stPressTop1Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cell1stPressTop1Current.Shading.Color = Colors.Yellow;
            }


            var cell1stPressTop1Goal = row1stPressTop.Cells[3];
            cell1stPressTop1Goal.Format.Font.Size = 8;
            cell1stPressTop1Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cell1stPressTop1Goal.AddParagraph($"{positionGoal}");

            var cell1stPressTop1Info = row1stPressTop2.Cells[1];
            cell1stPressTop1Info.Format.Font.Size = 8;
            cell1stPressTop1Info.MergeRight = 2;
            cell1stPressTop1Info.VerticalAlignment = VerticalAlignment.Center;
            cell1stPressTop1Info.Shading.Color = Colors.LightBlue;
            cell1stPressTop1Info.AddParagraph($"{GetClothingInfo(clothing1stPressTop1)}");
            #endregion

            #region 1st Press Top PM 2
            // ----- 1st Press Top Position PM 2 ----- //
            var cell1stPressTop2Past = row1stPressTop.Cells[4];
            cell1stPressTop2Past.Format.Font.Size = 8;
            cell1stPressTop2Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(2, 18);
            if (pastAverage > 0)
                cell1stPressTop2Past.AddParagraph($"{pastAverage}");

            var cell1stPressTop2Current = row1stPressTop.Cells[5];
            cell1stPressTop2Current.Format.Font.Size = 8;
            cell1stPressTop2Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothing1stPressTop2);
            positionGoal = GetPositionGoal(2, 18);
            if (currentAge > 0)
                cell1stPressTop2Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cell1stPressTop2Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cell1stPressTop2Current.Shading.Color = Colors.Yellow;
            }


            var cell1stPressTop2Goal = row1stPressTop.Cells[6];
            cell1stPressTop2Goal.Format.Font.Size = 8;
            cell1stPressTop2Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cell1stPressTop2Goal.AddParagraph($"{positionGoal}");

            var cell1stPressTop2Info = row1stPressTop2.Cells[4];
            cell1stPressTop2Info.Format.Font.Size = 8;
            cell1stPressTop2Info.MergeRight = 2;
            cell1stPressTop2Info.VerticalAlignment = VerticalAlignment.Center;
            cell1stPressTop2Info.Shading.Color = Colors.LightBlue;
            cell1stPressTop2Info.AddParagraph($"{GetClothingInfo(clothing1stPressTop2)}");
            #endregion

            #region 1st Press Top PM 3
            // ----- 1st Press Top Position PM 3 ----- //
            var cell1stPressTop3Past = row1stPressTop.Cells[7];
            cell1stPressTop3Past.Format.Font.Size = 8;
            cell1stPressTop3Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(3, 18);
            if (pastAverage > 0)
                cell1stPressTop3Past.AddParagraph($"{pastAverage}");

            var cell1stPressTop3Current = row1stPressTop.Cells[8];
            cell1stPressTop3Current.Format.Font.Size = 8;
            cell1stPressTop3Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothing1stPressTop3);
            positionGoal = GetPositionGoal(3, 18);
            if (currentAge > 0)
                cell1stPressTop3Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cell1stPressTop3Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cell1stPressTop3Current.Shading.Color = Colors.Yellow;
            }


            var cell1stPressTop3Goal = row1stPressTop.Cells[9];
            cell1stPressTop3Goal.Format.Font.Size = 8;
            cell1stPressTop3Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cell1stPressTop3Goal.AddParagraph($"{positionGoal}");

            var cell1stPressTop3Info = row1stPressTop2.Cells[7];
            cell1stPressTop3Info.Format.Font.Size = 8;
            cell1stPressTop3Info.MergeRight = 2;
            cell1stPressTop3Info.VerticalAlignment = VerticalAlignment.Center;
            cell1stPressTop3Info.Shading.Color = Colors.LightBlue;
            cell1stPressTop3Info.AddParagraph($"{GetClothingInfo(clothing1stPressTop3)}");
            #endregion

            #region 1st Press Top PM 4
            // ----- 1st Press Top Position PM 4 ----- //
            var cell1stPressTop4Past = row1stPressTop.Cells[10];
            cell1stPressTop4Past.Format.Font.Size = 8;
            cell1stPressTop4Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(4, 18);
            if (pastAverage > 0)
                cell1stPressTop4Past.AddParagraph($"{pastAverage}");

            var cell1stPressTop4Current = row1stPressTop.Cells[11];
            cell1stPressTop4Current.Format.Font.Size = 8;
            cell1stPressTop4Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothing1stPressTop1);
            positionGoal = GetPositionGoal(4, 18);
            if (currentAge > 0)
                cell1stPressTop4Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cell1stPressTop4Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cell1stPressTop4Current.Shading.Color = Colors.Yellow;
            }


            var cell1stPressTop4Goal = row1stPressTop.Cells[12];
            cell1stPressTop4Goal.Format.Font.Size = 8;
            cell1stPressTop4Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cell1stPressTop4Goal.AddParagraph($"{positionGoal}");

            var cell1stPressTop4Info = row1stPressTop2.Cells[10];
            cell1stPressTop4Info.Format.Font.Size = 8;
            cell1stPressTop4Info.MergeRight = 2;
            cell1stPressTop4Info.VerticalAlignment = VerticalAlignment.Center;
            cell1stPressTop4Info.Shading.Color = Colors.LightBlue;
            cell1stPressTop4Info.AddParagraph($"{GetClothingInfo(clothing1stPressTop4)}");
            #endregion

            #endregion // 1st Press Top Position

            #region 1st Press Bottom Position
            // add spacer row
            spacer = weeklyPMTable.AddRow();
            spacer.Height = 12;
            cellSpacer = spacer.Cells[0];
            cellSpacer.MergeRight = 12;
            cellSpacer.AddParagraph("");

            var clothing1stPressBottom1 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "1st press bottom roll" && c.PM_Number == 1);
            var clothing1stPressBottom2 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "1st press bottom roll" && c.PM_Number == 2);
            var clothing1stPressBottom3 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "1st press bottom roll" && c.PM_Number == 3);
            var clothing1stPressBottom4 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "1st press bottom roll" && c.PM_Number == 4);
            var row1stPressBottom = weeklyPMTable.AddRow();
            row1stPressBottom.Height = 10;
            var row1stPressBottom2 = weeklyPMTable.AddRow();
            row1stPressBottom2.Height = 10;

            var cell1stPressBottom = row1stPressBottom.Cells[0];
            cell1stPressBottom.Format.Font.Size = 8;
            cell1stPressBottom.VerticalAlignment = VerticalAlignment.Center;
            cell1stPressBottom.MergeDown = 1;
            cell1stPressBottom.AddParagraph("1ST PRESS BOTTOM");

            #region 1st Press Bottom PM 1
            // ----- 1st Press Bottom Position PM 1 ----- //
            var cell1stPressBottom1Past = row1stPressBottom.Cells[1];
            cell1stPressBottom1Past.Format.Font.Size = 8;
            cell1stPressBottom1Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(1, 19);
            if (pastAverage > 0)
                cell1stPressBottom1Past.AddParagraph($"{pastAverage}");

            var cell1stPressBottom1Current = row1stPressBottom.Cells[2];
            cell1stPressBottom1Current.Format.Font.Size = 8;
            cell1stPressBottom1Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothing1stPressBottom1);
            positionGoal = GetPositionGoal(1, 19);
            if (currentAge > 0)
                cell1stPressBottom1Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cell1stPressBottom1Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cell1stPressBottom1Current.Shading.Color = Colors.Yellow;
            }


            var cell1stPressBottom1Goal = row1stPressBottom.Cells[3];
            cell1stPressBottom1Goal.Format.Font.Size = 8;
            cell1stPressBottom1Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cell1stPressBottom1Goal.AddParagraph($"{positionGoal}");

            var cell1stPressBottom1Info = row1stPressBottom2.Cells[1];
            cell1stPressBottom1Info.Format.Font.Size = 8;
            cell1stPressBottom1Info.MergeRight = 2;
            cell1stPressBottom1Info.VerticalAlignment = VerticalAlignment.Center;
            cell1stPressBottom1Info.Shading.Color = Colors.LightBlue;
            cell1stPressBottom1Info.AddParagraph($"{GetClothingInfo(clothing1stPressBottom1)}");
            #endregion

            #region 1st Press Bottom PM 2
            // ----- 1st Press Bottom Position PM 2 ----- //
            var cell1stPressBottom2Past = row1stPressBottom.Cells[4];
            cell1stPressBottom2Past.Format.Font.Size = 8;
            cell1stPressBottom2Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(2, 19);
            if (pastAverage > 0)
                cell1stPressBottom2Past.AddParagraph($"{pastAverage}");

            var cell1stPressBottom2Current = row1stPressBottom.Cells[5];
            cell1stPressBottom2Current.Format.Font.Size = 8;
            cell1stPressBottom2Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothing1stPressBottom2);
            positionGoal = GetPositionGoal(2, 19);
            if (currentAge > 0)
                cell1stPressBottom2Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cell1stPressBottom2Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cell1stPressBottom2Current.Shading.Color = Colors.Yellow;
            }


            var cell1stPressBottom2Goal = row1stPressBottom.Cells[6];
            cell1stPressBottom2Goal.Format.Font.Size = 8;
            cell1stPressBottom2Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cell1stPressBottom2Goal.AddParagraph($"{positionGoal}");

            var cell1stPressBottom2Info = row1stPressBottom2.Cells[4];
            cell1stPressBottom2Info.Format.Font.Size = 8;
            cell1stPressBottom2Info.MergeRight = 2;
            cell1stPressBottom2Info.VerticalAlignment = VerticalAlignment.Center;
            cell1stPressBottom2Info.Shading.Color = Colors.LightBlue;
            cell1stPressBottom2Info.AddParagraph($"{GetClothingInfo(clothing1stPressBottom2)}");
            #endregion

            #region 1st Press Bottom PM 3
            // ----- 1st Press Bottom Position PM 3 ----- //
            var cell1stPressBottom3Past = row1stPressBottom.Cells[7];
            cell1stPressBottom3Past.Format.Font.Size = 8;
            cell1stPressBottom3Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(3, 19);
            if (pastAverage > 0)
                cell1stPressBottom3Past.AddParagraph($"{pastAverage}");

            var cell1stPressBottom3Current = row1stPressBottom.Cells[8];
            cell1stPressBottom3Current.Format.Font.Size = 8;
            cell1stPressBottom3Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothing1stPressBottom3);
            positionGoal = GetPositionGoal(3, 19);
            if (currentAge > 0)
                cell1stPressBottom3Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cell1stPressBottom3Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cell1stPressBottom3Current.Shading.Color = Colors.Yellow;
            }


            var cell1stPressBottom3Goal = row1stPressBottom.Cells[9];
            cell1stPressBottom3Goal.Format.Font.Size = 8;
            cell1stPressBottom3Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cell1stPressBottom3Goal.AddParagraph($"{positionGoal}");

            var cell1stPressBottom3Info = row1stPressBottom2.Cells[7];
            cell1stPressBottom3Info.Format.Font.Size = 8;
            cell1stPressBottom3Info.MergeRight = 2;
            cell1stPressBottom3Info.VerticalAlignment = VerticalAlignment.Center;
            cell1stPressBottom3Info.Shading.Color = Colors.LightBlue;
            cell1stPressBottom3Info.AddParagraph($"{GetClothingInfo(clothing1stPressBottom3)}");
            #endregion

            #region 1st Press Bottom PM 4
            // ----- 1st Press Bottom Position PM 4 ----- //
            var cell1stPressBottom4Past = row1stPressBottom.Cells[10];
            cell1stPressBottom4Past.Format.Font.Size = 8;
            cell1stPressBottom4Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(4, 19);
            if (pastAverage > 0)
                cell1stPressBottom4Past.AddParagraph($"{pastAverage}");

            var cell1stPressBottom4Current = row1stPressBottom.Cells[11];
            cell1stPressBottom4Current.Format.Font.Size = 8;
            cell1stPressBottom4Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothing1stPressBottom4);
            positionGoal = GetPositionGoal(4, 19);
            if (currentAge > 0)
                cell1stPressBottom4Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cell1stPressBottom4Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cell1stPressBottom4Current.Shading.Color = Colors.Yellow;
            }


            var cell1stPressBottom4Goal = row1stPressBottom.Cells[12];
            cell1stPressBottom4Goal.Format.Font.Size = 8;
            cell1stPressBottom4Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cell1stPressBottom4Goal.AddParagraph($"{positionGoal}");

            var cell1stPressBottom4Info = row1stPressBottom2.Cells[10];
            cell1stPressBottom4Info.Format.Font.Size = 8;
            cell1stPressBottom4Info.MergeRight = 2;
            cell1stPressBottom4Info.VerticalAlignment = VerticalAlignment.Center;
            cell1stPressBottom4Info.Shading.Color = Colors.LightBlue;
            cell1stPressBottom4Info.AddParagraph($"{GetClothingInfo(clothing1stPressBottom4)}");
            #endregion

            #endregion // 1st Press Bottom Position

            #region 2nd Press Top Position
            // add spacer row
            spacer = weeklyPMTable.AddRow();
            spacer.Height = 12;
            cellSpacer = spacer.Cells[0];
            cellSpacer.MergeRight = 12;
            cellSpacer.AddParagraph("");

            var clothing2ndPressTop1 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "2nd press top roll" && c.PM_Number == 1);
            var clothing2ndPressTop2 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "2nd press top roll" && c.PM_Number == 2);
            var clothing2ndPressTop3 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "2nd press top roll" && c.PM_Number == 3);
            var clothing2ndPressTop4 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "2nd press top roll" && c.PM_Number == 4);
            var row2ndPressTop = weeklyPMTable.AddRow();
            row2ndPressTop.Height = 10;
            var row2ndPressTop2 = weeklyPMTable.AddRow();
            row2ndPressTop2.Height = 10;

            var cell2ndPressTop = row2ndPressTop.Cells[0];
            cell2ndPressTop.Format.Font.Size = 8;
            cell2ndPressTop.VerticalAlignment = VerticalAlignment.Center;
            cell2ndPressTop.MergeDown = 1;
            cell2ndPressTop.AddParagraph("2ND PRESS TOP");

            #region 2nd Press Top PM 1
            // ----- 2nd Press Top Position PM 1 ----- //
            var cell2ndPressTop1Past = row2ndPressTop.Cells[1];
            cell2ndPressTop1Past.Format.Font.Size = 8;
            cell2ndPressTop1Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(1, 20);
            if (pastAverage > 0)
                cell2ndPressTop1Past.AddParagraph($"{pastAverage}");

            var cell2ndPressTop1Current = row2ndPressTop.Cells[2];
            cell2ndPressTop1Current.Format.Font.Size = 8;
            cell2ndPressTop1Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothing2ndPressTop1);
            positionGoal = GetPositionGoal(1, 20);
            if (currentAge > 0)
                cell2ndPressTop1Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cell2ndPressTop1Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cell2ndPressTop1Current.Shading.Color = Colors.Yellow;
            }


            var cell2ndPressTop1Goal = row2ndPressTop.Cells[3];
            cell2ndPressTop1Goal.Format.Font.Size = 8;
            cell2ndPressTop1Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cell2ndPressTop1Goal.AddParagraph($"{positionGoal}");

            var cell2ndPressTop1Info = row2ndPressTop2.Cells[1];
            cell2ndPressTop1Info.Format.Font.Size = 8;
            cell2ndPressTop1Info.MergeRight = 2;
            cell2ndPressTop1Info.VerticalAlignment = VerticalAlignment.Center;
            cell2ndPressTop1Info.Shading.Color = Colors.LightBlue;
            cell2ndPressTop1Info.AddParagraph($"{GetClothingInfo(clothing2ndPressTop1)}");
            #endregion

            #region 2nd Press Top PM 2
            // ----- 2nd Press Top Position PM 2 ----- //
            var cell2ndPressTop2Past = row2ndPressTop.Cells[4];
            cell2ndPressTop2Past.Format.Font.Size = 8;
            cell2ndPressTop2Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(2, 20);
            if (pastAverage > 0)
                cell2ndPressTop2Past.AddParagraph($"{pastAverage}");

            var cell2ndPressTop2Current = row2ndPressTop.Cells[5];
            cell2ndPressTop2Current.Format.Font.Size = 8;
            cell2ndPressTop2Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothing2ndPressTop2);
            positionGoal = GetPositionGoal(2, 20);
            if (currentAge > 0)
                cell2ndPressTop2Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cell2ndPressTop2Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cell2ndPressTop2Current.Shading.Color = Colors.Yellow;
            }


            var cell2ndPressTop2Goal = row2ndPressTop.Cells[6];
            cell2ndPressTop2Goal.Format.Font.Size = 8;
            cell2ndPressTop2Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cell2ndPressTop2Goal.AddParagraph($"{positionGoal}");

            var cell2ndPressTop2Info = row2ndPressTop2.Cells[4];
            cell2ndPressTop2Info.Format.Font.Size = 8;
            cell2ndPressTop2Info.MergeRight = 2;
            cell2ndPressTop2Info.VerticalAlignment = VerticalAlignment.Center;
            cell2ndPressTop2Info.Shading.Color = Colors.LightBlue;
            cell2ndPressTop2Info.AddParagraph($"{GetClothingInfo(clothing2ndPressTop2)}");
            #endregion

            #region 2nd Press Top PM 3
            // -----2nd Press Top Position PM 3---- - //
            var cell2ndPressTop3Past = row2ndPressTop.Cells[7];
            cell2ndPressTop3Past.Format.Font.Size = 8;
            cell2ndPressTop3Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(3, 20);
            if (pastAverage > 0)
                cell2ndPressTop3Past.AddParagraph($"{pastAverage}");

            var cell2ndPressTop3Current = row2ndPressTop.Cells[8];
            cell2ndPressTop3Current.Format.Font.Size = 8;
            cell2ndPressTop3Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothing2ndPressTop3);
            positionGoal = GetPositionGoal(3, 20);
            if (currentAge > 0)
                cell2ndPressTop3Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cell2ndPressTop3Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cell2ndPressTop3Current.Shading.Color = Colors.Yellow;
            }


            var cell2ndPressTop3Goal = row2ndPressTop.Cells[9];
            cell2ndPressTop3Goal.Format.Font.Size = 8;
            cell2ndPressTop3Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cell2ndPressTop3Goal.AddParagraph($"{positionGoal}");

            var cell2ndPressTop3Info = row2ndPressTop2.Cells[7];
            cell2ndPressTop3Info.Format.Font.Size = 8;
            cell2ndPressTop3Info.MergeRight = 2;
            cell2ndPressTop3Info.VerticalAlignment = VerticalAlignment.Center;
            cell2ndPressTop3Info.Shading.Color = Colors.LightBlue;
            cell2ndPressTop3Info.AddParagraph($"{GetClothingInfo(clothing2ndPressTop3)}");
            #endregion

            #region 2nd Press Top PM 4
            // -----2nd Press Top Position PM 4---- - //
            var cell2ndPressTop4Past = row2ndPressTop.Cells[10];
            cell2ndPressTop4Past.Format.Font.Size = 8;
            cell2ndPressTop4Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(4, 20);
            if (pastAverage > 0)
                cell2ndPressTop4Past.AddParagraph($"{pastAverage}");

            var cell2ndPressTop4Current = row2ndPressTop.Cells[11];
            cell2ndPressTop4Current.Format.Font.Size = 8;
            cell2ndPressTop4Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothing2ndPressTop4);
            positionGoal = GetPositionGoal(4, 20);
            if (currentAge > 0)
                cell2ndPressTop4Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cell2ndPressTop4Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cell2ndPressTop4Current.Shading.Color = Colors.Yellow;
            }


            var cell2ndPressTop4Goal = row2ndPressTop.Cells[12];
            cell2ndPressTop4Goal.Format.Font.Size = 8;
            cell2ndPressTop4Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cell2ndPressTop4Goal.AddParagraph($"{positionGoal}");

            var cell2ndPressTop4Info = row2ndPressTop2.Cells[10];
            cell2ndPressTop4Info.Format.Font.Size = 8;
            cell2ndPressTop4Info.MergeRight = 2;
            cell2ndPressTop4Info.VerticalAlignment = VerticalAlignment.Center;
            cell2ndPressTop4Info.Shading.Color = Colors.LightBlue;
            cell2ndPressTop4Info.AddParagraph($"{GetClothingInfo(clothing2ndPressTop4)}");
            #endregion

            #endregion // 2nd Press Top Position

            #region 2nd Press Bottom Position
            // add spacer row
            spacer = weeklyPMTable.AddRow();
            spacer.Height = 12;
            cellSpacer = spacer.Cells[0];
            cellSpacer.MergeRight = 12;
            cellSpacer.AddParagraph("");

            var clothing2ndPressBottom1 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "2nd press bottom roll" && c.PM_Number == 1);
            var clothing2ndPressBottom2 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "2nd press bottom roll" && c.PM_Number == 2);
            var clothing2ndPressBottom3 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "2nd press bottom roll" && c.PM_Number == 3);
            var clothing2ndPressBottom4 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "2nd press bottom roll" && c.PM_Number == 4);
            var row2ndPressBottom = weeklyPMTable.AddRow();
            row2ndPressBottom.Height = 10;
            var row2ndPressBottom2 = weeklyPMTable.AddRow();
            row2ndPressBottom2.Height = 10;

            var cell2ndPressBottom = row2ndPressBottom.Cells[0];
            cell2ndPressBottom.Format.Font.Size = 8;
            cell2ndPressBottom.VerticalAlignment = VerticalAlignment.Center;
            cell2ndPressBottom.MergeDown = 1;
            cell2ndPressBottom.AddParagraph("2ND PRESS BOTTOM");

            #region 2nd Press Bottom PM 1
            // ----- 2nd Press Bottom Position PM 1 ----- //
            var cell2ndPressBottom1Past = row2ndPressBottom.Cells[1];
            cell2ndPressBottom1Past.Format.Font.Size = 8;
            cell2ndPressBottom1Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(1, 21);
            if (pastAverage > 0)
                cell2ndPressBottom1Past.AddParagraph($"{pastAverage}");

            var cell2ndPressBottom1Current = row2ndPressBottom.Cells[2];
            cell2ndPressBottom1Current.Format.Font.Size = 8;
            cell2ndPressBottom1Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothing2ndPressBottom1);
            positionGoal = GetPositionGoal(1, 21);
            if (currentAge > 0)
                cell2ndPressBottom1Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cell2ndPressBottom1Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cell2ndPressBottom1Current.Shading.Color = Colors.Yellow;
            }


            var cell2ndPressBottom1Goal = row2ndPressBottom.Cells[3];
            cell2ndPressBottom1Goal.Format.Font.Size = 8;
            cell2ndPressBottom1Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cell2ndPressBottom1Goal.AddParagraph($"{positionGoal}");

            var cell2ndPressBottom1Info = row2ndPressBottom2.Cells[1];
            cell2ndPressBottom1Info.Format.Font.Size = 8;
            cell2ndPressBottom1Info.MergeRight = 2;
            cell2ndPressBottom1Info.VerticalAlignment = VerticalAlignment.Center;
            cell2ndPressBottom1Info.Shading.Color = Colors.LightBlue;
            cell2ndPressBottom1Info.AddParagraph($"{GetClothingInfo(clothing2ndPressBottom1)}");
            #endregion

            #region 2nd Press Bottom PM 2
            // ----- 2nd Press Bottom Position PM 2 ----- //
            var cell2ndPressBottom2Past = row2ndPressBottom.Cells[4];
            cell2ndPressBottom2Past.Format.Font.Size = 8;
            cell2ndPressBottom2Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(2, 21);
            if (pastAverage > 0)
                cell2ndPressBottom2Past.AddParagraph($"{pastAverage}");

            var cell2ndPressBottom2Current = row2ndPressBottom.Cells[5];
            cell2ndPressBottom2Current.Format.Font.Size = 8;
            cell2ndPressBottom2Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothing2ndPressBottom2);
            positionGoal = GetPositionGoal(2, 21);
            if (currentAge > 0)
                cell2ndPressBottom2Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cell2ndPressBottom2Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cell2ndPressBottom2Current.Shading.Color = Colors.Yellow;
            }


            var cell2ndPressBottom2Goal = row2ndPressBottom.Cells[6];
            cell2ndPressBottom2Goal.Format.Font.Size = 8;
            cell2ndPressBottom2Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cell2ndPressBottom2Goal.AddParagraph($"{positionGoal}");

            var cell2ndPressBottom2Info = row2ndPressBottom2.Cells[4];
            cell2ndPressBottom2Info.Format.Font.Size = 8;
            cell2ndPressBottom2Info.MergeRight = 2;
            cell2ndPressBottom2Info.VerticalAlignment = VerticalAlignment.Center;
            cell2ndPressBottom2Info.Shading.Color = Colors.LightBlue;
            cell2ndPressBottom2Info.AddParagraph($"{GetClothingInfo(clothing2ndPressBottom2)}");
            #endregion

            #region 2nd Press Bottom PM 3
            // ----- 2nd Press Bottom Position PM 3 ----- //
            var cell2ndPressBottom3Past = row2ndPressBottom.Cells[7];
            cell2ndPressBottom3Past.Format.Font.Size = 8;
            cell2ndPressBottom3Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(3, 21);
            if (pastAverage > 0)
                cell2ndPressBottom3Past.AddParagraph($"{pastAverage}");

            var cell2ndPressBottom3Current = row2ndPressBottom.Cells[8];
            cell2ndPressBottom3Current.Format.Font.Size = 8;
            cell2ndPressBottom3Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothing2ndPressBottom3);
            positionGoal = GetPositionGoal(3, 21);
            if (currentAge > 0)
                cell2ndPressBottom3Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cell2ndPressBottom3Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cell2ndPressBottom3Current.Shading.Color = Colors.Yellow;
            }


            var cell2ndPressBottom3Goal = row2ndPressBottom.Cells[9];
            cell2ndPressBottom3Goal.Format.Font.Size = 8;
            cell2ndPressBottom3Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cell2ndPressBottom3Goal.AddParagraph($"{positionGoal}");

            var cell2ndPressBottom3Info = row2ndPressBottom2.Cells[7];
            cell2ndPressBottom3Info.Format.Font.Size = 8;
            cell2ndPressBottom3Info.MergeRight = 2;
            cell2ndPressBottom3Info.VerticalAlignment = VerticalAlignment.Center;
            cell2ndPressBottom3Info.Shading.Color = Colors.LightBlue;
            cell2ndPressBottom3Info.AddParagraph($"{GetClothingInfo(clothing2ndPressBottom3)}");
            #endregion

            #region 2nd Press Bottom PM 4
            // ----- 2nd Press Bottom Position PM 4 ----- //
            var cell2ndPressBottom4Past = row2ndPressBottom.Cells[10];
            cell2ndPressBottom4Past.Format.Font.Size = 8;
            cell2ndPressBottom4Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(4, 21);
            if (pastAverage > 0)
                cell2ndPressBottom4Past.AddParagraph($"{pastAverage}");

            var cell2ndPressBottom4Current = row2ndPressBottom.Cells[11];
            cell2ndPressBottom4Current.Format.Font.Size = 8;
            cell2ndPressBottom4Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothing2ndPressBottom4);
            positionGoal = GetPositionGoal(4, 21);
            if (currentAge > 0)
                cell2ndPressBottom4Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cell2ndPressBottom4Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cell2ndPressBottom4Current.Shading.Color = Colors.Yellow;
            }


            var cell2ndPressBottom4Goal = row2ndPressBottom.Cells[12];
            cell2ndPressBottom4Goal.Format.Font.Size = 8;
            cell2ndPressBottom4Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cell2ndPressBottom4Goal.AddParagraph($"{positionGoal}");

            var cell2ndPressBottom4Info = row2ndPressBottom2.Cells[10];
            cell2ndPressBottom4Info.Format.Font.Size = 8;
            cell2ndPressBottom4Info.MergeRight = 2;
            cell2ndPressBottom4Info.VerticalAlignment = VerticalAlignment.Center;
            cell2ndPressBottom4Info.Shading.Color = Colors.LightBlue;
            cell2ndPressBottom4Info.AddParagraph($"{GetClothingInfo(clothing2ndPressBottom4)}");
            #endregion

            #endregion // 2nd Press Bottom Position

            #region 3rd Press Top Position
            // add spacer row
            spacer = weeklyPMTable.AddRow();
            spacer.Height = 12;
            cellSpacer = spacer.Cells[0];
            cellSpacer.MergeRight = 12;
            cellSpacer.AddParagraph("");

            var clothing3rdPressTop1 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "3rd press top roll" && c.PM_Number == 1);
            var clothing3rdPressTop2 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "3rd press top roll" && c.PM_Number == 2);
            var clothing3rdPressTop3 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "3rd press top roll" && c.PM_Number == 3);
            var clothing3rdPressTop4 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "3rd press top roll" && c.PM_Number == 4);
            var row3rdPressTop = weeklyPMTable.AddRow();
            row3rdPressTop.Height = 10;
            var row3rdPressTop2 = weeklyPMTable.AddRow();
            row3rdPressTop2.Height = 10;

            var cell3rdPressTop = row3rdPressTop.Cells[0];
            cell3rdPressTop.Format.Font.Size = 8;
            cell3rdPressTop.VerticalAlignment = VerticalAlignment.Center;
            cell3rdPressTop.MergeDown = 1;
            cell3rdPressTop.AddParagraph("3RD PRESS TOP");

            #region 3rd Press Top PM 1
            // ----- 3rd Press Top Position PM 1 ----- //
            var cell3rdPressTop1Past = row3rdPressTop.Cells[1];
            cell3rdPressTop1Past.Format.Font.Size = 8;
            cell3rdPressTop1Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(1, 22);
            if (pastAverage > 0)
                cell3rdPressTop1Past.AddParagraph($"{pastAverage}");

            var cell3rdPressTop1Current = row3rdPressTop.Cells[2];
            cell3rdPressTop1Current.Format.Font.Size = 8;
            cell3rdPressTop1Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothing3rdPressTop1);
            positionGoal = GetPositionGoal(1, 22);
            if (currentAge > 0)
                cell3rdPressTop1Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cell3rdPressTop1Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cell3rdPressTop1Current.Shading.Color = Colors.Yellow;
            }


            var cell3rdPressTop1Goal = row3rdPressTop.Cells[3];
            cell3rdPressTop1Goal.Format.Font.Size = 8;
            cell3rdPressTop1Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cell3rdPressTop1Goal.AddParagraph($"{positionGoal}");

            var cell3rdPressTop1Info = row3rdPressTop2.Cells[1];
            cell3rdPressTop1Info.Format.Font.Size = 8;
            cell3rdPressTop1Info.MergeRight = 2;
            cell3rdPressTop1Info.VerticalAlignment = VerticalAlignment.Center;
            cell3rdPressTop1Info.Shading.Color = Colors.LightBlue;
            cell3rdPressTop1Info.AddParagraph($"{GetClothingInfo(clothing3rdPressTop1)}");
            #endregion

            #region 3rd Press Top PM 2
            // ----- 3rd Press Top Position PM 2 ----- //
            var cell3rdPressTop2Past = row3rdPressTop.Cells[4];
            cell3rdPressTop2Past.Format.Font.Size = 8;
            cell3rdPressTop2Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(2, 22);
            if (pastAverage > 0)
                cell3rdPressTop2Past.AddParagraph($"{pastAverage}");

            var cell3rdPressTop2Current = row3rdPressTop.Cells[5];
            cell3rdPressTop2Current.Format.Font.Size = 8;
            cell3rdPressTop2Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothing3rdPressTop2);
            positionGoal = GetPositionGoal(2, 22);
            if (currentAge > 0)
                cell3rdPressTop2Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cell3rdPressTop2Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cell3rdPressTop2Current.Shading.Color = Colors.Yellow;
            }


            var cell3rdPressTop2Goal = row3rdPressTop.Cells[6];
            cell3rdPressTop2Goal.Format.Font.Size = 8;
            cell3rdPressTop2Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cell3rdPressTop2Goal.AddParagraph($"{positionGoal}");

            var cell3rdPressTop2Info = row3rdPressTop2.Cells[4];
            cell3rdPressTop2Info.Format.Font.Size = 8;
            cell3rdPressTop2Info.MergeRight = 2;
            cell3rdPressTop2Info.VerticalAlignment = VerticalAlignment.Center;
            cell3rdPressTop2Info.Shading.Color = Colors.LightBlue;
            cell3rdPressTop2Info.AddParagraph($"{GetClothingInfo(clothing3rdPressTop2)}");
            #endregion

            #region 3rd Press Top PM 3
            // ----- 3rd Press Top Position PM 3 ----- //
            var cell3rdPressTop3Past = row3rdPressTop.Cells[7];
            cell3rdPressTop3Past.Format.Font.Size = 8;
            cell3rdPressTop3Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(3, 22);
            if (pastAverage > 0)
                cell3rdPressTop3Past.AddParagraph($"{pastAverage}");

            var cell3rdPressTop3Current = row3rdPressTop.Cells[8];
            cell3rdPressTop3Current.Format.Font.Size = 8;
            cell3rdPressTop3Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothing3rdPressTop1);
            positionGoal = GetPositionGoal(3, 22);
            if (currentAge > 0)
                cell3rdPressTop3Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cell3rdPressTop3Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cell3rdPressTop3Current.Shading.Color = Colors.Yellow;
            }


            var cell3rdPressTop3Goal = row3rdPressTop.Cells[9];
            cell3rdPressTop3Goal.Format.Font.Size = 8;
            cell3rdPressTop3Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cell3rdPressTop3Goal.AddParagraph($"{positionGoal}");

            var cell3rdPressTop3Info = row3rdPressTop2.Cells[7];
            cell3rdPressTop3Info.Format.Font.Size = 8;
            cell3rdPressTop3Info.MergeRight = 2;
            cell3rdPressTop3Info.VerticalAlignment = VerticalAlignment.Center;
            cell3rdPressTop3Info.Shading.Color = Colors.LightBlue;
            cell3rdPressTop3Info.AddParagraph($"{GetClothingInfo(clothing3rdPressTop3)}");
            #endregion

            #region 3rd Press Top PM 4
            // ----- 3rd Press Top Position PM 4 ----- //
            var cell3rdPressTop4Past = row3rdPressTop.Cells[10];
            cell3rdPressTop4Past.Format.Font.Size = 8;
            cell3rdPressTop4Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(4, 22);
            if (pastAverage > 0)
                cell3rdPressTop4Past.AddParagraph($"{pastAverage}");

            var cell3rdPressTop4Current = row3rdPressTop.Cells[11];
            cell3rdPressTop4Current.Format.Font.Size = 8;
            cell3rdPressTop4Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothing3rdPressTop4);
            positionGoal = GetPositionGoal(4, 22);
            if (currentAge > 0)
                cell3rdPressTop4Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cell3rdPressTop4Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cell3rdPressTop4Current.Shading.Color = Colors.Yellow;
            }


            var cell3rdPressTop4Goal = row3rdPressTop.Cells[12];
            cell3rdPressTop4Goal.Format.Font.Size = 8;
            cell3rdPressTop4Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cell3rdPressTop4Goal.AddParagraph($"{positionGoal}");

            var cell3rdPressTop4Info = row3rdPressTop2.Cells[10];
            cell3rdPressTop4Info.Format.Font.Size = 8;
            cell3rdPressTop4Info.MergeRight = 2;
            cell3rdPressTop4Info.VerticalAlignment = VerticalAlignment.Center;
            cell3rdPressTop4Info.Shading.Color = Colors.LightBlue;
            cell3rdPressTop4Info.AddParagraph($"{GetClothingInfo(clothing3rdPressTop4)}");
            #endregion

            #endregion // 3rd Press Top Position

            #region 3rd Press Bottom Position
            // add spacer row
            spacer = weeklyPMTable.AddRow();
            spacer.Height = 12;
            cellSpacer = spacer.Cells[0];
            cellSpacer.MergeRight = 12;
            cellSpacer.AddParagraph("");

            var clothing3rdPressBottom1 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "3rd press bottom roll" && c.PM_Number == 1);
            var clothing3rdPressBottom2 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "3rd press bottom roll" && c.PM_Number == 2);
            var clothing3rdPressBottom3 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "3rd press bottom roll" && c.PM_Number == 3);
            var clothing3rdPressBottom4 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "3rd press bottom roll" && c.PM_Number == 4);
            var row3rdPressBottom = weeklyPMTable.AddRow();
            row3rdPressBottom.Height = 10;
            var row3rdPressBottom2 = weeklyPMTable.AddRow();
            row3rdPressBottom2.Height = 10;

            var cell3rdPressBottom = row3rdPressBottom.Cells[0];
            cell3rdPressBottom.Format.Font.Size = 8;
            cell3rdPressBottom.VerticalAlignment = VerticalAlignment.Center;
            cell3rdPressBottom.MergeDown = 1;
            cell3rdPressBottom.AddParagraph("3RD PRESS BOTTOM");

            #region 3rd Press Bottom PM 1
            // ----- 3rd Press Bottom Position PM 1 ----- //
            var cell3rdPressBottom1Past = row3rdPressBottom.Cells[1];
            cell3rdPressBottom1Past.Format.Font.Size = 8;
            cell3rdPressBottom1Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(1, 23);
            if (pastAverage > 0)
                cell3rdPressBottom1Past.AddParagraph($"{pastAverage}");

            var cell3rdPressBottom1Current = row3rdPressBottom.Cells[2];
            cell3rdPressBottom1Current.Format.Font.Size = 8;
            cell3rdPressBottom1Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothing3rdPressBottom1);
            positionGoal = GetPositionGoal(1, 23);
            if (currentAge > 0)
                cell3rdPressBottom1Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cell3rdPressBottom1Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cell3rdPressBottom1Current.Shading.Color = Colors.Yellow;
            }


            var cell3rdPressBottom1Goal = row3rdPressBottom.Cells[3];
            cell3rdPressBottom1Goal.Format.Font.Size = 8;
            cell3rdPressBottom1Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cell3rdPressBottom1Goal.AddParagraph($"{positionGoal}");

            var cell3rdPressBottom1Info = row3rdPressBottom2.Cells[1];
            cell3rdPressBottom1Info.Format.Font.Size = 8;
            cell3rdPressBottom1Info.MergeRight = 2;
            cell3rdPressBottom1Info.VerticalAlignment = VerticalAlignment.Center;
            cell3rdPressBottom1Info.Shading.Color = Colors.LightBlue;
            cell3rdPressBottom1Info.AddParagraph($"{GetClothingInfo(clothing3rdPressBottom1)}");
            #endregion

            #region 3rd Press Bottom PM 2
            // ----- 3rd Press Bottom Position PM 2 ----- //
            var cell3rdPressBottom2Past = row3rdPressBottom.Cells[4];
            cell3rdPressBottom2Past.Format.Font.Size = 8;
            cell3rdPressBottom2Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(2, 23);
            if (pastAverage > 0)
                cell3rdPressBottom2Past.AddParagraph($"{GetPastAverage(2, 23)}");

            var cell3rdPressBottom2Current = row3rdPressBottom.Cells[5];
            cell3rdPressBottom2Current.Format.Font.Size = 8;
            cell3rdPressBottom2Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothing3rdPressBottom2);
            positionGoal = GetPositionGoal(2, 23);
            if (currentAge > 0)
                cell3rdPressBottom2Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cell3rdPressBottom2Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cell3rdPressBottom2Current.Shading.Color = Colors.Yellow;
            }


            var cell3rdPressBottom2Goal = row3rdPressBottom.Cells[6];
            cell3rdPressBottom2Goal.Format.Font.Size = 8;
            cell3rdPressBottom2Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cell3rdPressBottom2Goal.AddParagraph($"{positionGoal}");

            var cell3rdPressBottom2Info = row3rdPressBottom2.Cells[4];
            cell3rdPressBottom2Info.Format.Font.Size = 8;
            cell3rdPressBottom2Info.MergeRight = 2;
            cell3rdPressBottom2Info.VerticalAlignment = VerticalAlignment.Center;
            cell3rdPressBottom2Info.Shading.Color = Colors.LightBlue;
            cell3rdPressBottom2Info.AddParagraph($"{GetClothingInfo(clothing3rdPressBottom2)}");
            #endregion

            #region 3rd Press Bottom PM 3
            // ----- 3rd Press Bottom Position PM 3 ----- //
            var cell3rdPressBottom3Past = row3rdPressBottom.Cells[7];
            cell3rdPressBottom3Past.Format.Font.Size = 8;
            cell3rdPressBottom3Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(3, 23);
            if (pastAverage > 0)
                cell3rdPressBottom3Past.AddParagraph($"{pastAverage}");

            var cell3rdPressBottom3Current = row3rdPressBottom.Cells[8];
            cell3rdPressBottom3Current.Format.Font.Size = 8;
            cell3rdPressBottom3Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothing3rdPressBottom3);
            positionGoal = GetPositionGoal(3, 23);
            if (currentAge > 0)
                cell3rdPressBottom3Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cell3rdPressBottom3Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cell3rdPressBottom3Current.Shading.Color = Colors.Yellow;
            }


            var cell3rdPressBottom3Goal = row3rdPressBottom.Cells[9];
            cell3rdPressBottom3Goal.Format.Font.Size = 8;
            cell3rdPressBottom3Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cell3rdPressBottom3Goal.AddParagraph($"{positionGoal}");

            var cell3rdPressBottom3Info = row3rdPressBottom2.Cells[7];
            cell3rdPressBottom3Info.Format.Font.Size = 8;
            cell3rdPressBottom3Info.MergeRight = 2;
            cell3rdPressBottom3Info.VerticalAlignment = VerticalAlignment.Center;
            cell3rdPressBottom3Info.Shading.Color = Colors.LightBlue;
            cell3rdPressBottom3Info.AddParagraph($"{GetClothingInfo(clothing3rdPressBottom3)}");
            #endregion

            #region 3rd Press Bottom PM 4
            // ----- 3rd Press Bottom Position PM 4 ----- //
            var cell3rdPressBottom4Past = row3rdPressBottom.Cells[10];
            cell3rdPressBottom4Past.Format.Font.Size = 8;
            cell3rdPressBottom4Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(4, 23);
            if (pastAverage > 0)
                cell3rdPressBottom4Past.AddParagraph($"{pastAverage}");

            var cell3rdPressBottom4Current = row3rdPressBottom.Cells[11];
            cell3rdPressBottom4Current.Format.Font.Size = 8;
            cell3rdPressBottom4Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothing3rdPressBottom4);
            positionGoal = GetPositionGoal(4, 23);
            if (currentAge > 0)
                cell3rdPressBottom4Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cell3rdPressBottom4Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cell3rdPressBottom4Current.Shading.Color = Colors.Yellow;
            }


            var cell3rdPressBottom4Goal = row3rdPressBottom.Cells[12];
            cell3rdPressBottom4Goal.Format.Font.Size = 8;
            cell3rdPressBottom4Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cell3rdPressBottom4Goal.AddParagraph($"{positionGoal}");

            var cell3rdPressBottom4Info = row3rdPressBottom2.Cells[10];
            cell3rdPressBottom4Info.Format.Font.Size = 8;
            cell3rdPressBottom4Info.MergeRight = 2;
            cell3rdPressBottom4Info.VerticalAlignment = VerticalAlignment.Center;
            cell3rdPressBottom4Info.Shading.Color = Colors.LightBlue;
            cell3rdPressBottom4Info.AddParagraph($"{GetClothingInfo(clothing3rdPressBottom4)}");
            #endregion

            #endregion // 3rd Press Bottom Position

            #region Smooth Top Position
            // add spacer row
            spacer = weeklyPMTable.AddRow();
            spacer.Height = 12;
            cellSpacer = spacer.Cells[0];
            cellSpacer.MergeRight = 12;
            cellSpacer.AddParagraph("");

            var clothingSmoothTop1 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "smoother top roll" && c.PM_Number == 1);
            var clothingSmoothTop2 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "smoother top roll" && c.PM_Number == 2);
            var clothingSmoothTop3 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "smoother top roll" && c.PM_Number == 3);
            var clothingSmoothTop4 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "smoother top roll" && c.PM_Number == 4);
            var rowSmoothTop = weeklyPMTable.AddRow();
            rowSmoothTop.Height = 10;
            var rowSmoothTop2 = weeklyPMTable.AddRow();
            rowSmoothTop2.Height = 10;

            var cellSmoothTop = rowSmoothTop.Cells[0];
            cellSmoothTop.Format.Font.Size = 8;
            cellSmoothTop.VerticalAlignment = VerticalAlignment.Center;
            cellSmoothTop.MergeDown = 1;
            cellSmoothTop.AddParagraph("SMOOTH TOP");

            #region Smooth Top PM 1
            // ----- Smooth Top Position PM 1 ----- //
            var cellSmoothTop1Past = rowSmoothTop.Cells[1];
            cellSmoothTop1Past.Format.Font.Size = 8;
            cellSmoothTop1Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(1, 24);
            if (pastAverage > 0)
                cellSmoothTop1Past.AddParagraph($"{pastAverage}");

            var cellSmoothTop1Current = rowSmoothTop.Cells[2];
            cellSmoothTop1Current.Format.Font.Size = 8;
            cellSmoothTop1Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothingSmoothTop1);
            positionGoal = GetPositionGoal(1, 24);
            if (currentAge > 0)
                cellSmoothTop1Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cellSmoothTop1Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cellSmoothTop1Current.Shading.Color = Colors.Yellow;
            }


            var cellSmoothTop1Goal = rowSmoothTop.Cells[3];
            cellSmoothTop1Goal.Format.Font.Size = 8;
            cellSmoothTop1Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cellSmoothTop1Goal.AddParagraph($"{positionGoal}");

            var cellSmoothTop1Info = rowSmoothTop2.Cells[1];
            cellSmoothTop1Info.Format.Font.Size = 8;
            cellSmoothTop1Info.MergeRight = 2;
            cellSmoothTop1Info.VerticalAlignment = VerticalAlignment.Center;
            cellSmoothTop1Info.Shading.Color = Colors.LightBlue;
            cellSmoothTop1Info.AddParagraph($"{GetClothingInfo(clothingSmoothTop1)}");
            #endregion

            #region Smooth Top PM 2
            // ----- Smooth Top Position PM 2 ----- //
            var cellSmoothTop2Past = rowSmoothTop.Cells[4];
            cellSmoothTop2Past.Format.Font.Size = 8;
            cellSmoothTop2Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(2, 24);
            if (pastAverage > 0)
                cellSmoothTop2Past.AddParagraph($"{pastAverage}");

            var cellSmoothTop2Current = rowSmoothTop.Cells[5];
            cellSmoothTop2Current.Format.Font.Size = 8;
            cellSmoothTop2Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothingSmoothTop2);
            positionGoal = GetPositionGoal(2, 24);
            if (currentAge > 0)
                cellSmoothTop2Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cellSmoothTop2Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cellSmoothTop2Current.Shading.Color = Colors.Yellow;
            }


            var cellSmoothTop2Goal = rowSmoothTop.Cells[6];
            cellSmoothTop2Goal.Format.Font.Size = 8;
            cellSmoothTop2Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cellSmoothTop2Goal.AddParagraph($"{positionGoal}");

            var cellSmoothTop2Info = rowSmoothTop2.Cells[4];
            cellSmoothTop2Info.Format.Font.Size = 8;
            cellSmoothTop2Info.MergeRight = 2;
            cellSmoothTop2Info.VerticalAlignment = VerticalAlignment.Center;
            cellSmoothTop2Info.Shading.Color = Colors.LightBlue;
            cellSmoothTop2Info.AddParagraph($"{GetClothingInfo(clothingSmoothTop2)}");
            #endregion

            #region Smooth Top PM 3
            // ----- Smooth Top Position PM 3 ----- //
            var cellSmoothTop3Past = rowSmoothTop.Cells[7];
            cellSmoothTop3Past.Format.Font.Size = 8;
            cellSmoothTop3Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(3, 24);
            if (pastAverage > 0)
                cellSmoothTop3Past.AddParagraph($"{pastAverage}");

            var cellSmoothTop3Current = rowSmoothTop.Cells[8];
            cellSmoothTop3Current.Format.Font.Size = 8;
            cellSmoothTop3Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothingSmoothTop3);
            positionGoal = GetPositionGoal(3, 24);
            if (currentAge > 0)
                cellSmoothTop3Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cellSmoothTop3Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cellSmoothTop3Current.Shading.Color = Colors.Yellow;
            }


            var cellSmoothTop3Goal = rowSmoothTop.Cells[9];
            cellSmoothTop3Goal.Format.Font.Size = 8;
            cellSmoothTop3Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cellSmoothTop3Goal.AddParagraph($"{positionGoal}");

            var cellSmoothTop3Info = rowSmoothTop2.Cells[7];
            cellSmoothTop3Info.Format.Font.Size = 8;
            cellSmoothTop3Info.MergeRight = 2;
            cellSmoothTop3Info.VerticalAlignment = VerticalAlignment.Center;
            cellSmoothTop3Info.Shading.Color = Colors.LightBlue;
            cellSmoothTop3Info.AddParagraph($"{GetClothingInfo(clothingSmoothTop3)}");
            #endregion

            #region Smooth Top PM 4
            // ----- Smooth Top Position PM 4 ----- //
            var cellSmoothTop4Past = rowSmoothTop.Cells[10];
            cellSmoothTop4Past.Format.Font.Size = 8;
            cellSmoothTop4Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(4, 24);
            if (pastAverage > 0)
                cellSmoothTop4Past.AddParagraph($"{pastAverage}");

            var cellSmoothTop4Current = rowSmoothTop.Cells[11];
            cellSmoothTop4Current.Format.Font.Size = 8;
            cellSmoothTop4Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothingSmoothTop4);
            positionGoal = GetPositionGoal(4, 24);
            if (currentAge > 0)
                cellSmoothTop4Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cellSmoothTop4Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cellSmoothTop4Current.Shading.Color = Colors.Yellow;
            }


            var cellSmoothTop4Goal = rowSmoothTop.Cells[12];
            cellSmoothTop4Goal.Format.Font.Size = 8;
            cellSmoothTop4Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cellSmoothTop4Goal.AddParagraph($"{positionGoal}");

            var cellSmoothTop4Info = rowSmoothTop2.Cells[10];
            cellSmoothTop4Info.Format.Font.Size = 8;
            cellSmoothTop4Info.MergeRight = 2;
            cellSmoothTop4Info.VerticalAlignment = VerticalAlignment.Center;
            cellSmoothTop4Info.Shading.Color = Colors.LightBlue;
            cellSmoothTop4Info.AddParagraph($"{GetClothingInfo(clothingSmoothTop4)}");
            #endregion

            #endregion // Smooth Top Position

            #region Smooth Bottom Position
            // add spacer row
            spacer = weeklyPMTable.AddRow();
            spacer.Height = 12;
            cellSpacer = spacer.Cells[0];
            cellSpacer.MergeRight = 12;
            cellSpacer.AddParagraph("");

            var clothingSmoothBottom1 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "smoother bottom roll" && c.PM_Number == 1);
            var clothingSmoothBottom2 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "smoother bottom roll" && c.PM_Number == 2);
            var clothingSmoothBottom3 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "smoother bottom roll" && c.PM_Number == 3);
            var clothingSmoothBottom4 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "smoother bottom roll" && c.PM_Number == 4);
            var rowSmoothBottom = weeklyPMTable.AddRow();
            rowSmoothBottom.Height = 10;
            var rowSmoothBottom2 = weeklyPMTable.AddRow();
            rowSmoothBottom2.Height = 10;

            var cellSmoothBottom = rowSmoothBottom.Cells[0];
            cellSmoothBottom.Format.Font.Size = 8;
            cellSmoothBottom.VerticalAlignment = VerticalAlignment.Center;
            cellSmoothBottom.MergeDown = 1;
            cellSmoothBottom.AddParagraph("SMOOTH BOTTOM");

            #region Smooth Bottom PM 1
            // ----- Smooth Bottom Position PM 1 ----- //
            var cellSmoothBottom1Past = rowSmoothBottom.Cells[1];
            cellSmoothBottom1Past.Format.Font.Size = 8;
            cellSmoothBottom1Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(1, 25);
            if (pastAverage > 0)
                cellSmoothBottom1Past.AddParagraph($"{pastAverage}");

            var cellSmoothBottom1Current = rowSmoothBottom.Cells[2];
            cellSmoothBottom1Current.Format.Font.Size = 8;
            cellSmoothBottom1Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothingSmoothBottom1);
            positionGoal = GetPositionGoal(1, 25);
            if (currentAge > 0)
                cellSmoothBottom1Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cellSmoothBottom1Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cellSmoothBottom1Current.Shading.Color = Colors.Yellow;
            }


            var cellSmoothBottom1Goal = rowSmoothBottom.Cells[3];
            cellSmoothBottom1Goal.Format.Font.Size = 8;
            cellSmoothBottom1Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cellSmoothBottom1Goal.AddParagraph($"{positionGoal}");

            var cellSmoothBottom1Info = rowSmoothBottom2.Cells[1];
            cellSmoothBottom1Info.Format.Font.Size = 8;
            cellSmoothBottom1Info.MergeRight = 2;
            cellSmoothBottom1Info.VerticalAlignment = VerticalAlignment.Center;
            cellSmoothBottom1Info.Shading.Color = Colors.LightBlue;
            cellSmoothBottom1Info.AddParagraph($"{GetClothingInfo(clothingSmoothBottom1)}");
            #endregion

            #region Smooth Bottom PM 2
            // ----- Smooth Bottom Position PM 2 ----- //
            var cellSmoothBottom2Past = rowSmoothBottom.Cells[4];
            cellSmoothBottom2Past.Format.Font.Size = 8;
            cellSmoothBottom2Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(2, 25);
            if (pastAverage > 0)
                cellSmoothBottom2Past.AddParagraph($"{pastAverage}");

            var cellSmoothBottom2Current = rowSmoothBottom.Cells[5];
            cellSmoothBottom2Current.Format.Font.Size = 8;
            cellSmoothBottom2Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothingSmoothBottom2);
            positionGoal = GetPositionGoal(2, 25);
            if (currentAge > 0)
                cellSmoothBottom2Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cellSmoothBottom2Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cellSmoothBottom2Current.Shading.Color = Colors.Yellow;
            }


            var cellSmoothBottom2Goal = rowSmoothBottom.Cells[6];
            cellSmoothBottom2Goal.Format.Font.Size = 8;
            cellSmoothBottom2Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cellSmoothBottom2Goal.AddParagraph($"{positionGoal}");

            var cellSmoothBottom2Info = rowSmoothBottom2.Cells[4];
            cellSmoothBottom2Info.Format.Font.Size = 8;
            cellSmoothBottom2Info.MergeRight = 2;
            cellSmoothBottom2Info.VerticalAlignment = VerticalAlignment.Center;
            cellSmoothBottom2Info.Shading.Color = Colors.LightBlue;
            cellSmoothBottom2Info.AddParagraph($"{GetClothingInfo(clothingSmoothBottom2)}");
            #endregion

            #region Smooth Bottom PM 3
            // ----- Smooth Bottom Position PM 3 ----- //
            var cellSmoothBottom3Past = rowSmoothBottom.Cells[7];
            cellSmoothBottom3Past.Format.Font.Size = 8;
            cellSmoothBottom3Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(3, 25);
            if (pastAverage > 0)
                cellSmoothBottom3Past.AddParagraph($"{pastAverage}");

            var cellSmoothBottom3Current = rowSmoothBottom.Cells[8];
            cellSmoothBottom3Current.Format.Font.Size = 8;
            cellSmoothBottom3Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothingSmoothBottom3);
            positionGoal = GetPositionGoal(3, 25);
            if (currentAge > 0)
                cellSmoothBottom3Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cellSmoothBottom3Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cellSmoothBottom3Current.Shading.Color = Colors.Yellow;
            }


            var cellSmoothBottom3Goal = rowSmoothBottom.Cells[9];
            cellSmoothBottom3Goal.Format.Font.Size = 8;
            cellSmoothBottom3Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cellSmoothBottom3Goal.AddParagraph($"{positionGoal}");

            var cellSmoothBottom3Info = rowSmoothBottom2.Cells[7];
            cellSmoothBottom3Info.Format.Font.Size = 8;
            cellSmoothBottom3Info.MergeRight = 2;
            cellSmoothBottom3Info.VerticalAlignment = VerticalAlignment.Center;
            cellSmoothBottom3Info.Shading.Color = Colors.LightBlue;
            cellSmoothBottom3Info.AddParagraph($"{GetClothingInfo(clothingSmoothBottom3)}");
            #endregion

            #region Smooth Bottom PM 4
            // ----- Smooth Bottom Position PM 4 ----- //
            var cellSmoothBottom4Past = rowSmoothBottom.Cells[10];
            cellSmoothBottom4Past.Format.Font.Size = 8;
            cellSmoothBottom4Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(4, 25);
            if (pastAverage > 0)
                cellSmoothBottom4Past.AddParagraph($"{pastAverage}");

            var cellSmoothBottom4Current = rowSmoothBottom.Cells[11];
            cellSmoothBottom4Current.Format.Font.Size = 8;
            cellSmoothBottom4Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothingSmoothBottom4);
            positionGoal = GetPositionGoal(4, 25);
            if (currentAge > 0)
                cellSmoothBottom4Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cellSmoothBottom4Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cellSmoothBottom4Current.Shading.Color = Colors.Yellow;
            }


            var cellSmoothBottom4Goal = rowSmoothBottom.Cells[12];
            cellSmoothBottom4Goal.Format.Font.Size = 8;
            cellSmoothBottom4Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cellSmoothBottom4Goal.AddParagraph($"{positionGoal}");

            var cellSmoothBottom4Info = rowSmoothBottom2.Cells[10];
            cellSmoothBottom4Info.Format.Font.Size = 8;
            cellSmoothBottom4Info.MergeRight = 2;
            cellSmoothBottom4Info.VerticalAlignment = VerticalAlignment.Center;
            cellSmoothBottom4Info.Shading.Color = Colors.LightBlue;
            cellSmoothBottom4Info.AddParagraph($"{GetClothingInfo(clothingSmoothBottom4)}");
            #endregion

            #endregion // Smooth Bottom Position

            #region Hard Size Press Position
            // add spacer row
            spacer = weeklyPMTable.AddRow();
            spacer.Height = 12;
            cellSpacer = spacer.Cells[0];
            cellSpacer.MergeRight = 12;
            cellSpacer.AddParagraph("");

            var clothingHardSizePress1 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "hard size roll" && c.PM_Number == 1);
            var clothingHardSizePress2 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "hard size roll" && c.PM_Number == 2);
            var clothingHardSizePress3 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "hard size roll" && c.PM_Number == 3);
            var clothingHardSizePress4 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "hard size roll" && c.PM_Number == 4);
            var rowHardSizePress = weeklyPMTable.AddRow();
            rowHardSizePress.Height = 10;
            var rowHardSizePress2 = weeklyPMTable.AddRow();
            rowHardSizePress2.Height = 10;

            var cellHardSizePress = rowHardSizePress.Cells[0];
            cellHardSizePress.Format.Font.Size = 8;
            cellHardSizePress.VerticalAlignment = VerticalAlignment.Center;
            cellHardSizePress.MergeDown = 1;
            cellHardSizePress.AddParagraph("HARD SIZE PRESS (TOP)");

            #region Hard Size Press PM 1
            // ----- Hard Size Press Position PM 1 ----- //
            var cellHardSizePress1Past = rowHardSizePress.Cells[1];
            cellHardSizePress1Past.Format.Font.Size = 8;
            cellHardSizePress1Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(1, 26);
            if (pastAverage > 0)
                cellHardSizePress1Past.AddParagraph($"{pastAverage}");

            var cellHardSizePress1Current = rowHardSizePress.Cells[2];
            cellHardSizePress1Current.Format.Font.Size = 8;
            cellHardSizePress1Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothingHardSizePress1);
            positionGoal = GetPositionGoal(1, 26);
            if (currentAge > 0)
                cellHardSizePress1Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cellHardSizePress1Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cellHardSizePress1Current.Shading.Color = Colors.Yellow;
            }


            var cellHardSizePress1Goal = rowHardSizePress.Cells[3];
            cellHardSizePress1Goal.Format.Font.Size = 8;
            cellHardSizePress1Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cellHardSizePress1Goal.AddParagraph($"{positionGoal}");

            var cellHardSizePress1Info = rowHardSizePress2.Cells[1];
            cellHardSizePress1Info.Format.Font.Size = 8;
            cellHardSizePress1Info.MergeRight = 2;
            cellHardSizePress1Info.VerticalAlignment = VerticalAlignment.Center;
            cellHardSizePress1Info.Shading.Color = Colors.LightBlue;
            cellHardSizePress1Info.AddParagraph($"{GetClothingInfo(clothingHardSizePress1)}");
            #endregion

            #region Hard Size Press PM 2
            // ----- Hard Size Press Position PM 2 ----- //
            var cellHardSizePress2Past = rowHardSizePress.Cells[4];
            cellHardSizePress2Past.Format.Font.Size = 8;
            cellHardSizePress2Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(2, 26);
            if (pastAverage > 0)
                cellHardSizePress2Past.AddParagraph($"{pastAverage}");

            var cellHardSizePress2Current = rowHardSizePress.Cells[5];
            cellHardSizePress2Current.Format.Font.Size = 8;
            cellHardSizePress2Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothingHardSizePress2);
            positionGoal = GetPositionGoal(2, 26);
            if (currentAge > 0)
                cellHardSizePress2Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cellHardSizePress2Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cellHardSizePress2Current.Shading.Color = Colors.Yellow;
            }


            var cellHardSizePress2Goal = rowHardSizePress.Cells[6];
            cellHardSizePress2Goal.Format.Font.Size = 8;
            cellHardSizePress2Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cellHardSizePress2Goal.AddParagraph($"{positionGoal}");

            var cellHardSizePress2Info = rowHardSizePress2.Cells[4];
            cellHardSizePress2Info.Format.Font.Size = 8;
            cellHardSizePress2Info.MergeRight = 2;
            cellHardSizePress2Info.VerticalAlignment = VerticalAlignment.Center;
            cellHardSizePress2Info.Shading.Color = Colors.LightBlue;
            cellHardSizePress2Info.AddParagraph($"{GetClothingInfo(clothingHardSizePress2)}");
            #endregion

            #region Hard Size Press PM 3
            // ----- Hard Size Press Position PM 3 ----- //
            var cellHardSizePress3Past = rowHardSizePress.Cells[7];
            cellHardSizePress3Past.Format.Font.Size = 8;
            cellHardSizePress3Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(3, 26);
            if (pastAverage > 0)
                cellHardSizePress3Past.AddParagraph($"{pastAverage}");

            var cellHardSizePress3Current = rowHardSizePress.Cells[8];
            cellHardSizePress3Current.Format.Font.Size = 8;
            cellHardSizePress3Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothingHardSizePress3);
            positionGoal = GetPositionGoal(3, 26);
            if (currentAge > 0)
                cellHardSizePress3Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cellHardSizePress3Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cellHardSizePress3Current.Shading.Color = Colors.Yellow;
            }


            var cellHardSizePress3Goal = rowHardSizePress.Cells[9];
            cellHardSizePress3Goal.Format.Font.Size = 8;
            cellHardSizePress3Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cellHardSizePress3Goal.AddParagraph($"{positionGoal}");

            var cellHardSizePress3Info = rowHardSizePress2.Cells[7];
            cellHardSizePress3Info.Format.Font.Size = 8;
            cellHardSizePress3Info.MergeRight = 2;
            cellHardSizePress3Info.VerticalAlignment = VerticalAlignment.Center;
            cellHardSizePress3Info.Shading.Color = Colors.LightBlue;
            cellHardSizePress3Info.AddParagraph($"{GetClothingInfo(clothingHardSizePress3)}");
            #endregion

            #region Hard Size Press PM 4
            // ----- Hard Size Press Position PM 4 ----- //
            var cellHardSizePress4Past = rowHardSizePress.Cells[10];
            cellHardSizePress4Past.Format.Font.Size = 8;
            cellHardSizePress4Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(4, 26);
            if (pastAverage > 0)
                cellHardSizePress4Past.AddParagraph($"{pastAverage}");

            var cellHardSizePress4Current = rowHardSizePress.Cells[11];
            cellHardSizePress4Current.Format.Font.Size = 8;
            cellHardSizePress4Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothingHardSizePress4);
            positionGoal = GetPositionGoal(4, 26);
            if (currentAge > 0)
                cellHardSizePress4Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cellHardSizePress4Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cellHardSizePress4Current.Shading.Color = Colors.Yellow;
            }


            var cellHardSizePress4Goal = rowHardSizePress.Cells[12];
            cellHardSizePress4Goal.Format.Font.Size = 8;
            cellHardSizePress4Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cellHardSizePress4Goal.AddParagraph($"{positionGoal}");

            var cellHardSizePress4Info = rowHardSizePress2.Cells[10];
            cellHardSizePress4Info.Format.Font.Size = 8;
            cellHardSizePress4Info.MergeRight = 2;
            cellHardSizePress4Info.VerticalAlignment = VerticalAlignment.Center;
            cellHardSizePress4Info.Shading.Color = Colors.LightBlue;
            cellHardSizePress4Info.AddParagraph($"{GetClothingInfo(clothingHardSizePress4)}");
            #endregion

            #endregion // Hard Size Press Position

            #region Soft Size Press Position
            // add spacer row
            spacer = weeklyPMTable.AddRow();
            spacer.Height = 12;
            cellSpacer = spacer.Cells[0];
            cellSpacer.MergeRight = 12;
            cellSpacer.AddParagraph("");

            var clothingSoftSizePress1 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "soft size roll" && c.PM_Number == 1);
            var clothingSoftSizePress2 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "soft size roll" && c.PM_Number == 2);
            var clothingSoftSizePress3 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "soft size roll" && c.PM_Number == 3);
            var clothingSoftSizePress4 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "soft size roll" && c.PM_Number == 4);
            var rowSoftSizePress = weeklyPMTable.AddRow();
            rowSoftSizePress.Height = 10;
            var rowSoftSizePress2 = weeklyPMTable.AddRow();
            rowSoftSizePress2.Height = 10;

            var cellSoftSizePress = rowSoftSizePress.Cells[0];
            cellSoftSizePress.Format.Font.Size = 8;
            cellSoftSizePress.VerticalAlignment = VerticalAlignment.Center;
            cellSoftSizePress.MergeDown = 1;
            cellSoftSizePress.AddParagraph("SOFT SIZE PRESS (BOTTOM)");

            #region Soft Size Press PM 1
            // ----- Soft Size Press Position PM 1 ----- //
            var cellSoftSizePress1Past = rowSoftSizePress.Cells[1];
            cellSoftSizePress1Past.Format.Font.Size = 8;
            cellSoftSizePress1Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(1, 27);
            if (pastAverage > 0)
                cellSoftSizePress1Past.AddParagraph($"{pastAverage}");

            var cellSoftSizePress1Current = rowSoftSizePress.Cells[2];
            cellSoftSizePress1Current.Format.Font.Size = 8;
            cellSoftSizePress1Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothingSoftSizePress1);
            positionGoal = GetPositionGoal(1, 27);
            if (currentAge > 0)
                cellSoftSizePress1Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cellSoftSizePress1Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cellSoftSizePress1Current.Shading.Color = Colors.Yellow;
            }


            var cellSoftSizePress1Goal = rowSoftSizePress.Cells[3];
            cellSoftSizePress1Goal.Format.Font.Size = 8;
            cellSoftSizePress1Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cellSoftSizePress1Goal.AddParagraph($"{positionGoal}");

            var cellSoftSizePress1Info = rowSoftSizePress2.Cells[1];
            cellSoftSizePress1Info.Format.Font.Size = 8;
            cellSoftSizePress1Info.MergeRight = 2;
            cellSoftSizePress1Info.VerticalAlignment = VerticalAlignment.Center;
            cellSoftSizePress1Info.Shading.Color = Colors.LightBlue;
            cellSoftSizePress1Info.AddParagraph($"{GetClothingInfo(clothingSoftSizePress1)}");
            #endregion

            #region Soft Size Press PM 2
            // ----- Soft Size Press Position PM 2 ----- //
            var cellSoftSizePress2Past = rowSoftSizePress.Cells[4];
            cellSoftSizePress2Past.Format.Font.Size = 8;
            cellSoftSizePress2Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(2, 27);
            if (pastAverage > 0)
                cellSoftSizePress2Past.AddParagraph($"{pastAverage}");

            var cellSoftSizePress2Current = rowSoftSizePress.Cells[5];
            cellSoftSizePress2Current.Format.Font.Size = 8;
            cellSoftSizePress2Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothingSoftSizePress2);
            positionGoal = GetPositionGoal(2, 27);
            if (currentAge > 0)
                cellSoftSizePress2Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cellSoftSizePress2Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cellSoftSizePress2Current.Shading.Color = Colors.Yellow;
            }


            var cellSoftSizePress2Goal = rowSoftSizePress.Cells[6];
            cellSoftSizePress2Goal.Format.Font.Size = 8;
            cellSoftSizePress2Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cellSoftSizePress2Goal.AddParagraph($"{positionGoal}");

            var cellSoftSizePress2Info = rowSoftSizePress2.Cells[4];
            cellSoftSizePress2Info.Format.Font.Size = 8;
            cellSoftSizePress2Info.MergeRight = 2;
            cellSoftSizePress2Info.VerticalAlignment = VerticalAlignment.Center;
            cellSoftSizePress2Info.Shading.Color = Colors.LightBlue;
            cellSoftSizePress2Info.AddParagraph($"{GetClothingInfo(clothingSoftSizePress2)}");
            #endregion

            #region Soft Size Press PM 3
            // ----- Soft Size Press Position PM 3 ----- //
            var cellSoftSizePress3Past = rowSoftSizePress.Cells[7];
            cellSoftSizePress3Past.Format.Font.Size = 8;
            cellSoftSizePress3Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(3, 27);
            if (pastAverage > 0)
                cellSoftSizePress3Past.AddParagraph($"{pastAverage}");

            var cellSoftSizePress3Current = rowSoftSizePress.Cells[8];
            cellSoftSizePress3Current.Format.Font.Size = 8;
            cellSoftSizePress3Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothingSoftSizePress3);
            positionGoal = GetPositionGoal(3, 27);
            if (currentAge > 0)
                cellSoftSizePress3Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cellSoftSizePress3Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cellSoftSizePress3Current.Shading.Color = Colors.Yellow;
            }


            var cellSoftSizePress3Goal = rowSoftSizePress.Cells[9];
            cellSoftSizePress3Goal.Format.Font.Size = 8;
            cellSoftSizePress3Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cellSoftSizePress3Goal.AddParagraph($"{positionGoal}");

            var cellSoftSizePress3Info = rowSoftSizePress2.Cells[7];
            cellSoftSizePress3Info.Format.Font.Size = 8;
            cellSoftSizePress3Info.MergeRight = 2;
            cellSoftSizePress3Info.VerticalAlignment = VerticalAlignment.Center;
            cellSoftSizePress3Info.Shading.Color = Colors.LightBlue;
            cellSoftSizePress3Info.AddParagraph($"{GetClothingInfo(clothingSoftSizePress3)}");
            #endregion

            #region Soft Size Press PM 4
            // ----- Soft Size Press Position PM 4 ----- //
            var cellSoftSizePress4Past = rowSoftSizePress.Cells[10];
            cellSoftSizePress4Past.Format.Font.Size = 8;
            cellSoftSizePress4Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(4, 27);
            if (pastAverage > 0)
                cellSoftSizePress4Past.AddParagraph($"{pastAverage}");

            var cellSoftSizePress4Current = rowSoftSizePress.Cells[11];
            cellSoftSizePress4Current.Format.Font.Size = 8;
            cellSoftSizePress4Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothingSoftSizePress4);
            positionGoal = GetPositionGoal(4, 27);
            if (currentAge > 0)
                cellSoftSizePress4Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cellSoftSizePress4Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cellSoftSizePress4Current.Shading.Color = Colors.Yellow;
            }


            var cellSoftSizePress4Goal = rowSoftSizePress.Cells[12];
            cellSoftSizePress4Goal.Format.Font.Size = 8;
            cellSoftSizePress4Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cellSoftSizePress4Goal.AddParagraph($"{positionGoal}");

            var cellSoftSizePress4Info = rowSoftSizePress2.Cells[10];
            cellSoftSizePress4Info.Format.Font.Size = 8;
            cellSoftSizePress4Info.MergeRight = 2;
            cellSoftSizePress4Info.VerticalAlignment = VerticalAlignment.Center;
            cellSoftSizePress4Info.Shading.Color = Colors.LightBlue;
            cellSoftSizePress4Info.AddParagraph($"{GetClothingInfo(clothingSoftSizePress4)}");
            #endregion

            #endregion // Soft Size Press Position

            #region Aquitherm Top Position

            var clothingAquithermTop1 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "aquatherm roll" && c.PM_Number == 1);
            var clothingAquithermTop2 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "aquatherm roll" && c.PM_Number == 2);
            var clothingAquithermTop3 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "aquatherm roll" && c.PM_Number == 3);
            var clothingAquithermTop4 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "aquatherm roll" && c.PM_Number == 4);
            var rowAquithermTop = weeklyPMTable.AddRow();
            rowAquithermTop.Height = 10;
            var rowAquithermTop2 = weeklyPMTable.AddRow();
            rowAquithermTop2.Height = 10;

            var cellAquithermTop = rowAquithermTop.Cells[0];
            cellAquithermTop.Format.Font.Size = 8;
            cellAquithermTop.VerticalAlignment = VerticalAlignment.Center;
            cellAquithermTop.MergeDown = 1;
            cellAquithermTop.AddParagraph("AQUATHERM TOP");

            #region Aquitherm Top PM 1
            // ----- Aquitherm Top Position PM 1 ----- //
            var cellAquithermTop1Past = rowAquithermTop.Cells[1];
            cellAquithermTop1Past.Format.Font.Size = 8;
            cellAquithermTop1Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(1, 28);
            if (pastAverage > 0)
                cellAquithermTop1Past.AddParagraph($"{pastAverage}");

            var cellAquithermTop1Current = rowAquithermTop.Cells[2];
            cellAquithermTop1Current.Format.Font.Size = 8;
            cellAquithermTop1Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothingAquithermTop1);
            positionGoal = GetPositionGoal(1, 28);
            if (currentAge > 0)
                cellAquithermTop1Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cellAquithermTop1Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cellAquithermTop1Current.Shading.Color = Colors.Yellow;
            }


            var cellAquithermTop1Goal = rowAquithermTop.Cells[3];
            cellAquithermTop1Goal.Format.Font.Size = 8;
            cellAquithermTop1Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cellAquithermTop1Goal.AddParagraph($"{positionGoal}");

            var cellAquithermTop1Info = rowAquithermTop2.Cells[1];
            cellAquithermTop1Info.Format.Font.Size = 8;
            cellAquithermTop1Info.MergeRight = 2;
            cellAquithermTop1Info.VerticalAlignment = VerticalAlignment.Center;
            cellAquithermTop1Info.Shading.Color = Colors.LightBlue;
            cellAquithermTop1Info.AddParagraph($"{GetClothingInfo(clothingAquithermTop1)}");
            #endregion

            #region Aquitherm Top PM 2
            // ----- Aquitherm Top Position PM 2 ----- //
            var cellAquithermTop2Past = rowAquithermTop.Cells[4];
            cellAquithermTop2Past.Format.Font.Size = 8;
            cellAquithermTop2Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(2, 28);
            if (pastAverage > 0)
                cellAquithermTop2Past.AddParagraph($"{pastAverage}");

            var cellAquithermTop2Current = rowAquithermTop.Cells[5];
            cellAquithermTop2Current.Format.Font.Size = 8;
            cellAquithermTop2Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothingAquithermTop2);
            positionGoal = GetPositionGoal(2, 28);
            if (currentAge > 0)
                cellAquithermTop2Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cellAquithermTop2Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cellAquithermTop2Current.Shading.Color = Colors.Yellow;
            }


            var cellAquithermTop2Goal = rowAquithermTop.Cells[6];
            cellAquithermTop2Goal.Format.Font.Size = 8;
            cellAquithermTop2Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cellAquithermTop2Goal.AddParagraph($"{positionGoal}");

            var cellAquithermTop2Info = rowAquithermTop2.Cells[4];
            cellAquithermTop2Info.Format.Font.Size = 8;
            cellAquithermTop2Info.MergeRight = 2;
            cellAquithermTop2Info.VerticalAlignment = VerticalAlignment.Center;
            cellAquithermTop2Info.Shading.Color = Colors.LightBlue;
            cellAquithermTop2Info.AddParagraph($"{GetClothingInfo(clothingAquithermTop2)}");
            #endregion

            #region Aquitherm Top PM 3
            // ----- Aquitherm Top Position PM 3 ----- //
            var cellAquithermTop3Past = rowAquithermTop.Cells[7];
            cellAquithermTop3Past.Format.Font.Size = 8;
            cellAquithermTop3Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(3, 28);
            if (pastAverage > 0)
                cellAquithermTop3Past.AddParagraph($"{pastAverage}");

            var cellAquithermTop3Current = rowAquithermTop.Cells[8];
            cellAquithermTop3Current.Format.Font.Size = 8;
            cellAquithermTop3Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothingAquithermTop3);
            positionGoal = GetPositionGoal(3, 28);
            if (currentAge > 0)
                cellAquithermTop3Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cellAquithermTop3Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cellAquithermTop3Current.Shading.Color = Colors.Yellow;
            }


            var cellAquithermTop3Goal = rowAquithermTop.Cells[9];
            cellAquithermTop3Goal.Format.Font.Size = 8;
            cellAquithermTop3Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cellAquithermTop3Goal.AddParagraph($"{positionGoal}");

            var cellAquithermTop3Info = rowAquithermTop2.Cells[7];
            cellAquithermTop3Info.Format.Font.Size = 8;
            cellAquithermTop3Info.MergeRight = 2;
            cellAquithermTop3Info.VerticalAlignment = VerticalAlignment.Center;
            cellAquithermTop3Info.Shading.Color = Colors.LightBlue;
            cellAquithermTop3Info.AddParagraph($"{GetClothingInfo(clothingAquithermTop3)}");
            #endregion

            #region Aquitherm Top PM 4
            // ----- Aquitherm Top Position PM 4 ----- //
            var cellAquithermTop4Past = rowAquithermTop.Cells[10];
            cellAquithermTop4Past.Format.Font.Size = 8;
            cellAquithermTop4Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(4, 28);
            if (pastAverage > 0)
                cellAquithermTop4Past.AddParagraph($"{pastAverage}");

            var cellAquithermTop4Current = rowAquithermTop.Cells[11];
            cellAquithermTop4Current.Format.Font.Size = 8;
            cellAquithermTop4Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothingAquithermTop4);
            positionGoal = GetPositionGoal(4, 28);
            if (currentAge > 0)
                cellAquithermTop4Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cellAquithermTop4Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cellAquithermTop4Current.Shading.Color = Colors.Yellow;
            }


            var cellAquithermTop4Goal = rowAquithermTop.Cells[12];
            cellAquithermTop4Goal.Format.Font.Size = 8;
            cellAquithermTop4Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cellAquithermTop4Goal.AddParagraph($"{positionGoal}");

            var cellAquithermTop4Info = rowAquithermTop2.Cells[10];
            cellAquithermTop4Info.Format.Font.Size = 8;
            cellAquithermTop4Info.MergeRight = 2;
            cellAquithermTop4Info.VerticalAlignment = VerticalAlignment.Center;
            cellAquithermTop4Info.Shading.Color = Colors.LightBlue;
            cellAquithermTop4Info.AddParagraph($"{GetClothingInfo(clothingAquithermTop4)}");
            #endregion

            #endregion // Aquitherm Top Position

            #region Nibco Bottom Position
            // add spacer row
            spacer = weeklyPMTable.AddRow();
            spacer.Height = 12;
            cellSpacer = spacer.Cells[0];
            cellSpacer.MergeRight = 12;
            cellSpacer.AddParagraph("");

            var clothingNibcoBottom1 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "nipco roll" && c.PM_Number == 1);
            var clothingNibcoBottom2 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "nipco roll" && c.PM_Number == 2);
            var clothingNibcoBottom3 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "nipco roll" && c.PM_Number == 3);
            var clothingNibcoBottom4 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "nipco roll" && c.PM_Number == 4);
            var rowNibcoBottom = weeklyPMTable.AddRow();
            rowNibcoBottom.Height = 10;
            var rowNibcoBottom2 = weeklyPMTable.AddRow();
            rowNibcoBottom2.Height = 10;

            var cellNibcoBottom = rowNibcoBottom.Cells[0];
            cellNibcoBottom.Format.Font.Size = 8;
            cellNibcoBottom.VerticalAlignment = VerticalAlignment.Center;
            cellNibcoBottom.MergeDown = 1;
            cellNibcoBottom.AddParagraph("NIPCO BOTTOM");

            #region Nibco Bottom PM 1
            // ----- Nibco Bottom Position PM 1 ----- //
            var cellNibcoBottom1Past = rowNibcoBottom.Cells[1];
            cellNibcoBottom1Past.Format.Font.Size = 8;
            cellNibcoBottom1Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(1, 29);
            if (pastAverage > 0)
                cellNibcoBottom1Past.AddParagraph($"{pastAverage}");

            var cellNibcoBottom1Current = rowNibcoBottom.Cells[2];
            cellNibcoBottom1Current.Format.Font.Size = 8;
            cellNibcoBottom1Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothingNibcoBottom1);
            positionGoal = GetPositionGoal(1, 29);
            if (currentAge > 0)
                cellNibcoBottom1Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cellNibcoBottom1Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cellNibcoBottom1Current.Shading.Color = Colors.Yellow;
            }


            var cellNibcoBottom1Goal = rowNibcoBottom.Cells[3];
            cellNibcoBottom1Goal.Format.Font.Size = 8;
            cellNibcoBottom1Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cellNibcoBottom1Goal.AddParagraph($"{positionGoal}");

            var cellNibcoBottom1Info = rowNibcoBottom2.Cells[1];
            cellNibcoBottom1Info.Format.Font.Size = 8;
            cellNibcoBottom1Info.MergeRight = 2;
            cellNibcoBottom1Info.VerticalAlignment = VerticalAlignment.Center;
            cellNibcoBottom1Info.Shading.Color = Colors.LightBlue;
            cellNibcoBottom1Info.AddParagraph($"{GetClothingInfo(clothingNibcoBottom1)}");
            #endregion

            #region Nibco Bottom PM 2
            // ----- Nibco Bottom Position PM 2 ----- //
            var cellNibcoBottom2Past = rowNibcoBottom.Cells[4];
            cellNibcoBottom2Past.Format.Font.Size = 8;
            cellNibcoBottom2Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(2, 29);
            if (pastAverage > 0)
                cellNibcoBottom2Past.AddParagraph($"{pastAverage}");

            var cellNibcoBottom2Current = rowNibcoBottom.Cells[5];
            cellNibcoBottom2Current.Format.Font.Size = 8;
            cellNibcoBottom2Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothingNibcoBottom2);
            positionGoal = GetPositionGoal(2, 29);
            if (currentAge > 0)
                cellNibcoBottom2Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cellNibcoBottom2Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cellNibcoBottom2Current.Shading.Color = Colors.Yellow;
            }


            var cellNibcoBottom2Goal = rowNibcoBottom.Cells[6];
            cellNibcoBottom2Goal.Format.Font.Size = 8;
            cellNibcoBottom2Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cellNibcoBottom2Goal.AddParagraph($"{positionGoal}");

            var cellNibcoBottom2Info = rowNibcoBottom2.Cells[4];
            cellNibcoBottom2Info.Format.Font.Size = 8;
            cellNibcoBottom2Info.MergeRight = 2;
            cellNibcoBottom2Info.VerticalAlignment = VerticalAlignment.Center;
            cellNibcoBottom2Info.Shading.Color = Colors.LightBlue;
            cellNibcoBottom2Info.AddParagraph($"{GetClothingInfo(clothingNibcoBottom2)}");
            #endregion

            #region Nibco Bottom PM 3
            // ----- Nibco Bottom Position PM 3 ----- //
            var cellNibcoBottom3Past = rowNibcoBottom.Cells[7];
            cellNibcoBottom3Past.Format.Font.Size = 8;
            cellNibcoBottom3Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(3, 29);
            if (pastAverage > 0)
                cellNibcoBottom3Past.AddParagraph($"{pastAverage}");

            var cellNibcoBottom3Current = rowNibcoBottom.Cells[8];
            cellNibcoBottom3Current.Format.Font.Size = 8;
            cellNibcoBottom3Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothingNibcoBottom3);
            positionGoal = GetPositionGoal(3, 29);
            if (currentAge > 0)
                cellNibcoBottom3Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cellNibcoBottom3Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cellNibcoBottom3Current.Shading.Color = Colors.Yellow;
            }


            var cellNibcoBottom3Goal = rowNibcoBottom.Cells[9];
            cellNibcoBottom3Goal.Format.Font.Size = 8;
            cellNibcoBottom3Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cellNibcoBottom3Goal.AddParagraph($"{positionGoal}");

            var cellNibcoBottom3Info = rowNibcoBottom2.Cells[7];
            cellNibcoBottom3Info.Format.Font.Size = 8;
            cellNibcoBottom3Info.MergeRight = 2;
            cellNibcoBottom3Info.VerticalAlignment = VerticalAlignment.Center;
            cellNibcoBottom3Info.Shading.Color = Colors.LightBlue;
            cellNibcoBottom3Info.AddParagraph($"{GetClothingInfo(clothingNibcoBottom3)}");
            #endregion

            #region Nibco Bottom PM 4
            // ----- Nibco Bottom Position PM 4 ----- //
            var cellNibcoBottom4Past = rowNibcoBottom.Cells[10];
            cellNibcoBottom4Past.Format.Font.Size = 8;
            cellNibcoBottom4Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(4, 29);
            if (pastAverage > 0)
                cellNibcoBottom4Past.AddParagraph($"{pastAverage}");

            var cellNibcoBottom4Current = rowNibcoBottom.Cells[11];
            cellNibcoBottom4Current.Format.Font.Size = 8;
            cellNibcoBottom4Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothingNibcoBottom4);
            positionGoal = GetPositionGoal(4, 29);
            if (currentAge > 0)
                cellNibcoBottom4Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cellNibcoBottom4Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cellNibcoBottom4Current.Shading.Color = Colors.Yellow;
            }


            var cellNibcoBottom4Goal = rowNibcoBottom.Cells[12];
            cellNibcoBottom4Goal.Format.Font.Size = 8;
            cellNibcoBottom4Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cellNibcoBottom4Goal.AddParagraph($"{positionGoal}");

            var cellNibcoBottom4Info = rowNibcoBottom2.Cells[10];
            cellNibcoBottom4Info.Format.Font.Size = 8;
            cellNibcoBottom4Info.MergeRight = 2;
            cellNibcoBottom4Info.VerticalAlignment = VerticalAlignment.Center;
            cellNibcoBottom4Info.Shading.Color = Colors.LightBlue;
            cellNibcoBottom4Info.AddParagraph($"{GetClothingInfo(clothingNibcoBottom4)}");
            #endregion

            #endregion // Nipco Bottom Position

            #region Couch Roll Position
            // add spacer row
            spacer = weeklyPMTable.AddRow();
            spacer.Height = 12;
            cellSpacer = spacer.Cells[0];
            cellSpacer.MergeRight = 12;
            cellSpacer.AddParagraph("");

            var clothingCouch1 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "couch roll" && c.PM_Number == 1);
            var clothingCouch2 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "couch roll" && c.PM_Number == 2);
            var clothingCouch3 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "couch roll" && c.PM_Number == 3);
            var clothingCouch4 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "couch roll" && c.PM_Number == 4);
            var rowCouch = weeklyPMTable.AddRow();
            rowCouch.Height = 10;
            var rowCouch2 = weeklyPMTable.AddRow();
            rowCouch2.Height = 10;

            var cellCouch = rowCouch.Cells[0];
            cellCouch.Format.Font.Size = 8;
            cellCouch.VerticalAlignment = VerticalAlignment.Center;
            cellCouch.MergeDown = 1;
            cellCouch.AddParagraph("COUCH ROLL");

            #region Couch Roll PM 1
            // ----- Couch Position PM 1 ----- //
            var cellCouch1Past = rowCouch.Cells[1];
            cellCouch1Past.Format.Font.Size = 8;
            cellCouch1Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(1, 30);
            if (pastAverage > 0)
                cellCouch1Past.AddParagraph($"{pastAverage}");

            var cellCouch1Current = rowCouch.Cells[2];
            cellCouch1Current.Format.Font.Size = 8;
            cellCouch1Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothingCouch1);
            positionGoal = GetPositionGoal(1, 30);
            if (currentAge > 0)
                cellCouch1Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cellCouch1Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cellCouch1Current.Shading.Color = Colors.Yellow;
            }


            var cellCouch1Goal = rowCouch.Cells[3];
            cellCouch1Goal.Format.Font.Size = 8;
            cellCouch1Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cellCouch1Goal.AddParagraph($"{positionGoal}");

            var cellCouch1Info = rowCouch2.Cells[1];
            cellCouch1Info.Format.Font.Size = 8;
            cellCouch1Info.MergeRight = 2;
            cellCouch1Info.VerticalAlignment = VerticalAlignment.Center;
            cellCouch1Info.Shading.Color = Colors.LightBlue;
            cellCouch1Info.AddParagraph($"{GetClothingInfo(clothingCouch1)}");
            #endregion

            #region Couch Roll PM 2
            // ----- Couch Position PM 2 ----- //
            var cellCouch2Past = rowCouch.Cells[4];
            cellCouch2Past.Format.Font.Size = 8;
            cellCouch2Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(2, 30);
            if (pastAverage > 0)
                cellCouch2Past.AddParagraph($"{pastAverage}");

            var cellCouch2Current = rowCouch.Cells[5];
            cellCouch2Current.Format.Font.Size = 8;
            cellCouch2Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothingCouch2);
            positionGoal = GetPositionGoal(2, 30);
            if (currentAge > 0)
                cellCouch2Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cellCouch2Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cellCouch2Current.Shading.Color = Colors.Yellow;
            }


            var cellCouch2Goal = rowCouch.Cells[6];
            cellCouch2Goal.Format.Font.Size = 8;
            cellCouch2Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cellCouch2Goal.AddParagraph($"{positionGoal}");

            var cellCouch2Info = rowCouch2.Cells[4];
            cellCouch2Info.Format.Font.Size = 8;
            cellCouch2Info.MergeRight = 2;
            cellCouch2Info.VerticalAlignment = VerticalAlignment.Center;
            cellCouch2Info.Shading.Color = Colors.LightBlue;
            cellCouch2Info.AddParagraph($"{GetClothingInfo(clothingCouch2)}");
            #endregion

            #region Couch Roll PM 3
            // ----- Couch Position PM 3 ----- //
            var cellCouch3Past = rowCouch.Cells[7];
            cellCouch3Past.Format.Font.Size = 8;
            cellCouch3Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(3, 30);
            if (pastAverage > 0)
                cellCouch3Past.AddParagraph($"{pastAverage}");

            var cellCouch3Current = rowCouch.Cells[8];
            cellCouch3Current.Format.Font.Size = 8;
            cellCouch3Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothingCouch3);
            positionGoal = GetPositionGoal(3, 30);
            if (currentAge > 0)
                cellCouch3Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cellCouch3Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cellCouch3Current.Shading.Color = Colors.Yellow;
            }


            var cellCouch3Goal = rowCouch.Cells[9];
            cellCouch3Goal.Format.Font.Size = 8;
            cellCouch3Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cellCouch3Goal.AddParagraph($"{positionGoal}");

            var cellCouch3Info = rowCouch2.Cells[7];
            cellCouch3Info.Format.Font.Size = 8;
            cellCouch3Info.MergeRight = 2;
            cellCouch3Info.VerticalAlignment = VerticalAlignment.Center;
            cellCouch3Info.Shading.Color = Colors.LightBlue;
            cellCouch3Info.AddParagraph($"{GetClothingInfo(clothingCouch3)}");
            #endregion

            #region Couch Roll PM 4
            // ----- Couch Position PM 4 ----- //
            var cellCouch4Past = rowCouch.Cells[10];
            cellCouch4Past.Format.Font.Size = 8;
            cellCouch4Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(4, 30);
            if (pastAverage > 0)
                cellCouch4Past.AddParagraph($"{pastAverage}");

            var cellCouch4Current = rowCouch.Cells[11];
            cellCouch4Current.Format.Font.Size = 8;
            cellCouch4Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothingCouch4);
            positionGoal = GetPositionGoal(4, 30);
            if (currentAge > 0)
                cellCouch4Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cellCouch4Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cellCouch4Current.Shading.Color = Colors.Yellow;
            }


            var cellCouch4Goal = rowCouch.Cells[12];
            cellCouch4Goal.Format.Font.Size = 8;
            cellCouch4Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cellCouch4Goal.AddParagraph($"{positionGoal}");

            var cellCouch4Info = rowCouch2.Cells[10];
            cellCouch4Info.Format.Font.Size = 8;
            cellCouch4Info.MergeRight = 2;
            cellCouch4Info.VerticalAlignment = VerticalAlignment.Center;
            cellCouch4Info.Shading.Color = Colors.LightBlue;
            cellCouch4Info.AddParagraph($"{GetClothingInfo(clothingCouch4)}");
            #endregion

            #endregion // Couch Roll Position

            #region L_In Hope Roll Position
            // add spacer row
            spacer = weeklyPMTable.AddRow();
            spacer.Height = 12;
            cellSpacer = spacer.Cells[0];
            cellSpacer.MergeRight = 12;
            cellSpacer.AddParagraph("");

            var clothingLInHope1 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "l_in hope roll" && c.PM_Number == 1);
            var clothingLInHope2 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "l_in hope roll" && c.PM_Number == 2);
            var clothingLInHope3 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "l_in hope roll" && c.PM_Number == 3);
            var clothingLInHope4 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "l_in hope roll" && c.PM_Number == 4);
            var rowLInHope = weeklyPMTable.AddRow();
            rowLInHope.Height = 10;
            var rowLInHope2 = weeklyPMTable.AddRow();
            rowLInHope2.Height = 10;

            var cellLInHope = rowLInHope.Cells[0];
            cellLInHope.Format.Font.Size = 8;
            cellLInHope.VerticalAlignment = VerticalAlignment.Center;
            cellLInHope.MergeDown = 1;
            cellLInHope.AddParagraph("L-IN HOPE ROLL");

            #region L_In Hope Roll PM 1
            // ----- L_In Hope Roll Position PM 1 ----- //
            var cellLInHope1Past = rowLInHope.Cells[1];
            cellLInHope1Past.Format.Font.Size = 8;
            cellLInHope1Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(1, 31);
            if (pastAverage > 0)
                cellLInHope1Past.AddParagraph($"{pastAverage}");

            var cellLInHope1Current = rowLInHope.Cells[2];
            cellLInHope1Current.Format.Font.Size = 8;
            cellLInHope1Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothingLInHope1);
            positionGoal = GetPositionGoal(1, 31);
            if (currentAge > 0)
                cellLInHope1Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cellLInHope1Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cellLInHope1Current.Shading.Color = Colors.Yellow;
            }


            var cellLInHope1Goal = rowLInHope.Cells[3];
            cellLInHope1Goal.Format.Font.Size = 8;
            cellLInHope1Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cellLInHope1Goal.AddParagraph($"{positionGoal}");

            var cellLInHope1Info = rowLInHope2.Cells[1];
            cellLInHope1Info.Format.Font.Size = 8;
            cellLInHope1Info.MergeRight = 2;
            cellLInHope1Info.VerticalAlignment = VerticalAlignment.Center;
            cellLInHope1Info.Shading.Color = Colors.LightBlue;
            cellLInHope1Info.AddParagraph($"{GetClothingInfo(clothingLInHope1)}");
            #endregion

            #region L_In Hope Roll PM 2
            // ----- L_In Hope Roll Position PM 2 ----- //
            var cellLInHope2Past = rowLInHope.Cells[4];
            cellLInHope2Past.Format.Font.Size = 8;
            cellLInHope2Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(2, 31);
            if (pastAverage > 0)
                cellLInHope2Past.AddParagraph($"{pastAverage}");

            var cellLInHope2Current = rowLInHope.Cells[5];
            cellLInHope2Current.Format.Font.Size = 8;
            cellLInHope2Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothingLInHope2);
            positionGoal = GetPositionGoal(2, 31);
            if (currentAge > 0)
                cellLInHope2Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cellLInHope2Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cellLInHope2Current.Shading.Color = Colors.Yellow;
            }


            var cellLInHope2Goal = rowLInHope.Cells[6];
            cellLInHope2Goal.Format.Font.Size = 8;
            cellLInHope2Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cellLInHope2Goal.AddParagraph($"{positionGoal}");

            var cellLInHope2Info = rowLInHope2.Cells[4];
            cellLInHope2Info.Format.Font.Size = 8;
            cellLInHope2Info.MergeRight = 2;
            cellLInHope2Info.VerticalAlignment = VerticalAlignment.Center;
            cellLInHope2Info.Shading.Color = Colors.LightBlue;
            cellLInHope2Info.AddParagraph($"{GetClothingInfo(clothingLInHope2)}");
            #endregion

            #region L_In Hope Roll PM 3
            // ----- L_In Hope Roll Position PM 3 ----- //
            var cellLInHope3Past = rowLInHope.Cells[7];
            cellLInHope3Past.Format.Font.Size = 8;
            cellLInHope3Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(3, 31);
            if (pastAverage > 0)
                cellLInHope3Past.AddParagraph($"{pastAverage}");

            var cellLInHope3Current = rowLInHope.Cells[8];
            cellLInHope3Current.Format.Font.Size = 8;
            cellLInHope3Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothingLInHope1);
            positionGoal = GetPositionGoal(3, 31);
            if (currentAge > 0)
                cellLInHope3Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cellLInHope3Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cellLInHope3Current.Shading.Color = Colors.Yellow;
            }


            var cellLInHope3Goal = rowLInHope.Cells[9];
            cellLInHope3Goal.Format.Font.Size = 8;
            cellLInHope3Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cellLInHope3Goal.AddParagraph($"{positionGoal}");

            var cellLInHope3Info = rowLInHope2.Cells[7];
            cellLInHope3Info.Format.Font.Size = 8;
            cellLInHope3Info.MergeRight = 2;
            cellLInHope3Info.VerticalAlignment = VerticalAlignment.Center;
            cellLInHope3Info.Shading.Color = Colors.LightBlue;
            cellLInHope3Info.AddParagraph($"{GetClothingInfo(clothingLInHope3)}");
            #endregion

            #region L_In Hope Roll PM 4
            // ----- L_In Hope Roll Position PM 4 ----- //
            var cellLInHope4Past = rowLInHope.Cells[10];
            cellLInHope4Past.Format.Font.Size = 8;
            cellLInHope4Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(4, 31);
            if (pastAverage > 0)
                cellLInHope4Past.AddParagraph($"{pastAverage}");

            var cellLInHope4Current = rowLInHope.Cells[11];
            cellLInHope4Current.Format.Font.Size = 8;
            cellLInHope4Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothingLInHope4);
            positionGoal = GetPositionGoal(4, 31);
            if (currentAge > 0)
                cellLInHope4Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cellLInHope4Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cellLInHope4Current.Shading.Color = Colors.Yellow;
            }


            var cellLInHope4Goal = rowLInHope.Cells[12];
            cellLInHope4Goal.Format.Font.Size = 8;
            cellLInHope4Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cellLInHope4Goal.AddParagraph($"{positionGoal}");

            var cellLInHope4Info = rowLInHope2.Cells[10];
            cellLInHope4Info.Format.Font.Size = 8;
            cellLInHope4Info.MergeRight = 2;
            cellLInHope4Info.VerticalAlignment = VerticalAlignment.Center;
            cellLInHope4Info.Shading.Color = Colors.LightBlue;
            cellLInHope4Info.AddParagraph($"{GetClothingInfo(clothingLInHope4)}");
            #endregion

            #endregion // L_In Hope Roll Position

            #region L_Out Hope Roll Position
            // add spacer row
            //spacer = weeklyPMTable.AddRow();
            //spacer.Height = 12;
            //cellSpacer = spacer.Cells[0];
            //cellSpacer.MergeRight = 12;
            //cellSpacer.AddParagraph("");

            var clothingLOutHope1 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "l_out hope roll" && c.PM_Number == 1);
            var clothingLOutHope2 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "l_out hope roll" && c.PM_Number == 2);
            var clothingLOutHope3 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "l_out hope roll" && c.PM_Number == 3);
            var clothingLOutHope4 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "l_out hope roll" && c.PM_Number == 4);
            var rowLOutHope = weeklyPMTable.AddRow();
            rowLInHope.Height = 10;
            var rowLOutHope2 = weeklyPMTable.AddRow();
            rowLInHope2.Height = 10;

            var cellLOutHope = rowLOutHope.Cells[0];
            cellLOutHope.Format.Font.Size = 8;
            cellLOutHope.VerticalAlignment = VerticalAlignment.Center;
            cellLOutHope.MergeDown = 1;
            cellLOutHope.AddParagraph("L-OUT HOPE ROLL");

            #region L_Out Hope PM 1
            // ----- L_Out Hope Roll Position PM 1 ----- //
            var cellLOutHope1Past = rowLOutHope.Cells[1];
            cellLOutHope1Past.Format.Font.Size = 8;
            cellLOutHope1Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(1, 32);
            if (pastAverage > 0)
                cellLOutHope1Past.AddParagraph($"{pastAverage}");

            var cellLOutHope1Current = rowLOutHope.Cells[2];
            cellLOutHope1Current.Format.Font.Size = 8;
            cellLOutHope1Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothingLOutHope1);
            positionGoal = GetPositionGoal(1, 32);
            if (currentAge > 0)
                cellLOutHope1Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cellLOutHope1Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cellLOutHope1Current.Shading.Color = Colors.Yellow;
            }


            var cellLOutHope1Goal = rowLOutHope.Cells[3];
            cellLOutHope1Goal.Format.Font.Size = 8;
            cellLOutHope1Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cellLOutHope1Goal.AddParagraph($"{positionGoal}");

            var cellLOutHope1Info = rowLOutHope2.Cells[1];
            cellLOutHope1Info.Format.Font.Size = 8;
            cellLOutHope1Info.MergeRight = 2;
            cellLOutHope1Info.VerticalAlignment = VerticalAlignment.Center;
            cellLOutHope1Info.Shading.Color = Colors.LightBlue;
            cellLOutHope1Info.AddParagraph($"{GetClothingInfo(clothingLOutHope1)}");
            #endregion

            #region L_Out Hope PM 2
            // ----- L_Out Hope Roll Position PM 2 ----- //
            var cellLOutHope2Past = rowLOutHope.Cells[4];
            cellLOutHope2Past.Format.Font.Size = 8;
            cellLOutHope2Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(2, 32);
            if (pastAverage > 0)
                cellLOutHope2Past.AddParagraph($"{pastAverage}");

            var cellLOutHope2Current = rowLOutHope.Cells[5];
            cellLOutHope2Current.Format.Font.Size = 8;
            cellLOutHope2Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothingLOutHope2);
            positionGoal = GetPositionGoal(2, 32);
            if (currentAge > 0)
                cellLOutHope2Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cellLOutHope2Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cellLOutHope2Current.Shading.Color = Colors.Yellow;
            }


            var cellLOutHope2Goal = rowLOutHope.Cells[6];
            cellLOutHope2Goal.Format.Font.Size = 8;
            cellLOutHope2Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cellLOutHope2Goal.AddParagraph($"{positionGoal}");

            var cellLOutHope2Info = rowLOutHope2.Cells[4];
            cellLOutHope2Info.Format.Font.Size = 8;
            cellLOutHope2Info.MergeRight = 2;
            cellLOutHope2Info.VerticalAlignment = VerticalAlignment.Center;
            cellLOutHope2Info.Shading.Color = Colors.LightBlue;
            cellLOutHope2Info.AddParagraph($"{GetClothingInfo(clothingLOutHope2)}");
            #endregion

            #region L_Out Hope PM 3
            // ----- L_Out Hope Roll Position PM 3 ----- //
            var cellLOutHope3Past = rowLOutHope.Cells[7];
            cellLOutHope3Past.Format.Font.Size = 8;
            cellLOutHope3Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(3, 32);
            if (pastAverage > 0)
                cellLOutHope3Past.AddParagraph($"{pastAverage}");

            var cellLOutHope3Current = rowLOutHope.Cells[8];
            cellLOutHope3Current.Format.Font.Size = 8;
            cellLOutHope3Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothingLOutHope3);
            positionGoal = GetPositionGoal(3, 32);
            if (currentAge > 0)
                cellLOutHope3Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cellLOutHope3Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cellLOutHope3Current.Shading.Color = Colors.Yellow;
            }


            var cellLOutHope3Goal = rowLOutHope.Cells[9];
            cellLOutHope3Goal.Format.Font.Size = 8;
            cellLOutHope3Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cellLOutHope3Goal.AddParagraph($"{positionGoal}");

            var cellLOutHope3Info = rowLOutHope2.Cells[7];
            cellLOutHope3Info.Format.Font.Size = 8;
            cellLOutHope3Info.MergeRight = 2;
            cellLOutHope3Info.VerticalAlignment = VerticalAlignment.Center;
            cellLOutHope3Info.Shading.Color = Colors.LightBlue;
            cellLOutHope3Info.AddParagraph($"{GetClothingInfo(clothingLOutHope3)}");
            #endregion

            #region L_Out Hope PM 4
            // ----- L_Out Hope Roll Position PM 4 ----- //
            var cellLOutHope4Past = rowLOutHope.Cells[10];
            cellLOutHope4Past.Format.Font.Size = 8;
            cellLOutHope4Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(4, 32);
            if (pastAverage > 0)
                cellLOutHope4Past.AddParagraph($"{pastAverage}");

            var cellLOutHope4Current = rowLOutHope.Cells[11];
            cellLOutHope4Current.Format.Font.Size = 8;
            cellLOutHope4Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothingLOutHope4);
            positionGoal = GetPositionGoal(4, 32);
            if (currentAge > 0)
                cellLOutHope4Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cellLOutHope4Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cellLOutHope4Current.Shading.Color = Colors.Yellow;
            }


            var cellLOutHope4Goal = rowLOutHope.Cells[12];
            cellLOutHope4Goal.Format.Font.Size = 8;
            cellLOutHope4Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cellLOutHope4Goal.AddParagraph($"{positionGoal}");

            var cellLOutHope4Info = rowLOutHope2.Cells[10];
            cellLOutHope4Info.Format.Font.Size = 8;
            cellLOutHope4Info.MergeRight = 2;
            cellLOutHope4Info.VerticalAlignment = VerticalAlignment.Center;
            cellLOutHope4Info.Shading.Color = Colors.LightBlue;
            cellLOutHope4Info.AddParagraph($"{GetClothingInfo(clothingLOutHope4)}");
            #endregion

            #endregion // L_Out Hope Roll Position

            #region Bottom Press Wringer Position
            // add spacer row
            spacer = weeklyPMTable.AddRow();
            spacer.Height = 12;
            cellSpacer = spacer.Cells[0];
            cellSpacer.MergeRight = 12;
            cellSpacer.AddParagraph("");

            var clothingBottomPressWringer1 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "bottom press wringer" && c.PM_Number == 1);
            var clothingBottomPressWringer2 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "bottom press wringer" && c.PM_Number == 2);
            var clothingBottomPressWringer3 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "bottom press wringer" && c.PM_Number == 3);
            var clothingBottomPressWringer4 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "bottom press wringer" && c.PM_Number == 4);
            var rowBottomPressWringer = weeklyPMTable.AddRow();
            rowBottomPressWringer.Height = 10;
            var rowBottomPressWringer2 = weeklyPMTable.AddRow();
            rowBottomPressWringer2.Height = 10;

            var cellBottomPressWringer = rowBottomPressWringer.Cells[0];
            cellBottomPressWringer.Format.Font.Size = 8;
            cellBottomPressWringer.VerticalAlignment = VerticalAlignment.Center;
            cellBottomPressWringer.MergeDown = 1;
            cellBottomPressWringer.AddParagraph("BOTTOM PRESS WRINGER");

            #region Bottom Press Wringer PM 1
            // ----- Bottom Press Wringer Position PM 1 ----- //
            var cellBottomPressWringer1Past = rowBottomPressWringer.Cells[1];
            cellBottomPressWringer1Past.Format.Font.Size = 8;
            cellBottomPressWringer1Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(1, 33);
            if (pastAverage > 0)
                cellBottomPressWringer1Past.AddParagraph($"{pastAverage}");

            var cellBottomPressWringer1Current = rowBottomPressWringer.Cells[2];
            cellBottomPressWringer1Current.Format.Font.Size = 8;
            cellBottomPressWringer1Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothingBottomPressWringer1);
            positionGoal = GetPositionGoal(1, 33);
            if (currentAge > 0)
                cellBottomPressWringer1Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cellBottomPressWringer1Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cellBottomPressWringer1Current.Shading.Color = Colors.Yellow;
            }


            var cellBottomPressWringer1Goal = rowBottomPressWringer.Cells[3];
            cellBottomPressWringer1Goal.Format.Font.Size = 8;
            cellBottomPressWringer1Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cellBottomPressWringer1Goal.AddParagraph($"{positionGoal}");

            var cellBottomPressWringer1Info = rowBottomPressWringer2.Cells[1];
            cellBottomPressWringer1Info.Format.Font.Size = 8;
            cellBottomPressWringer1Info.MergeRight = 2;
            cellBottomPressWringer1Info.VerticalAlignment = VerticalAlignment.Center;
            cellBottomPressWringer1Info.Shading.Color = Colors.LightBlue;
            cellBottomPressWringer1Info.AddParagraph($"{GetClothingInfo(clothingBottomPressWringer1)}");
            #endregion

            #region Bottom Press Wringer PM 2
            // ----- Bottom Press Wringer Position PM 2 ----- //
            var cellBottomPressWringer2Past = rowBottomPressWringer.Cells[4];
            cellBottomPressWringer2Past.Format.Font.Size = 8;
            cellBottomPressWringer2Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(2, 33);
            if (pastAverage > 0)
                cellBottomPressWringer2Past.AddParagraph($"{pastAverage}");

            var cellBottomPressWringer2Current = rowBottomPressWringer.Cells[5];
            cellBottomPressWringer2Current.Format.Font.Size = 8;
            cellBottomPressWringer2Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothingBottomPressWringer2);
            positionGoal = GetPositionGoal(2, 33);
            if (currentAge > 0)
                cellBottomPressWringer2Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cellBottomPressWringer2Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cellBottomPressWringer2Current.Shading.Color = Colors.Yellow;
            }


            var cellBottomPressWringer2Goal = rowBottomPressWringer.Cells[6];
            cellBottomPressWringer2Goal.Format.Font.Size = 8;
            cellBottomPressWringer2Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cellBottomPressWringer2Goal.AddParagraph($"{positionGoal}");

            var cellBottomPressWringer2Info = rowBottomPressWringer2.Cells[4];
            cellBottomPressWringer2Info.Format.Font.Size = 8;
            cellBottomPressWringer2Info.MergeRight = 2;
            cellBottomPressWringer2Info.VerticalAlignment = VerticalAlignment.Center;
            cellBottomPressWringer2Info.Shading.Color = Colors.LightBlue;
            cellBottomPressWringer2Info.AddParagraph($"{GetClothingInfo(clothingBottomPressWringer2)}");
            #endregion

            #region Bottom Press Wringer PM 3
            // ----- Bottom Press Wringer Position PM 3 ----- //
            var cellBottomPressWringer3Past = rowBottomPressWringer.Cells[7];
            cellBottomPressWringer3Past.Format.Font.Size = 8;
            cellBottomPressWringer3Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(3, 33);
            if (pastAverage > 0)
                cellBottomPressWringer3Past.AddParagraph($"{pastAverage}");

            var cellBottomPressWringer3Current = rowBottomPressWringer.Cells[8];
            cellBottomPressWringer3Current.Format.Font.Size = 8;
            cellBottomPressWringer3Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothingBottomPressWringer3);
            positionGoal = GetPositionGoal(3, 33);
            if (currentAge > 0)
                cellBottomPressWringer3Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cellBottomPressWringer3Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cellBottomPressWringer3Current.Shading.Color = Colors.Yellow;
            }


            var cellBottomPressWringer3Goal = rowBottomPressWringer.Cells[9];
            cellBottomPressWringer3Goal.Format.Font.Size = 8;
            cellBottomPressWringer3Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cellBottomPressWringer3Goal.AddParagraph($"{positionGoal}");

            var cellBottomPressWringer3Info = rowBottomPressWringer2.Cells[7];
            cellBottomPressWringer3Info.Format.Font.Size = 8;
            cellBottomPressWringer3Info.MergeRight = 2;
            cellBottomPressWringer3Info.VerticalAlignment = VerticalAlignment.Center;
            cellBottomPressWringer3Info.Shading.Color = Colors.LightBlue;
            cellBottomPressWringer3Info.AddParagraph($"{GetClothingInfo(clothingBottomPressWringer3)}");
            #endregion

            #region Bottom Press Wringer PM 4
            // ----- Bottom Press Wringer Position PM 4 ----- //
            var cellBottomPressWringer4Past = rowBottomPressWringer.Cells[10];
            cellBottomPressWringer4Past.Format.Font.Size = 8;
            cellBottomPressWringer4Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(4, 33);
            if (pastAverage > 0)
                cellBottomPressWringer4Past.AddParagraph($"{pastAverage}");

            var cellBottomPressWringer4Current = rowBottomPressWringer.Cells[11];
            cellBottomPressWringer4Current.Format.Font.Size = 8;
            cellBottomPressWringer4Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothingBottomPressWringer4);
            positionGoal = GetPositionGoal(4, 33);
            if (currentAge > 0)
                cellBottomPressWringer4Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cellBottomPressWringer4Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cellBottomPressWringer4Current.Shading.Color = Colors.Yellow;
            }


            var cellBottomPressWringer4Goal = rowBottomPressWringer.Cells[12];
            cellBottomPressWringer4Goal.Format.Font.Size = 8;
            cellBottomPressWringer4Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cellBottomPressWringer4Goal.AddParagraph($"{positionGoal}");

            var cellBottomPressWringer4Info = rowBottomPressWringer2.Cells[10];
            cellBottomPressWringer4Info.Format.Font.Size = 8;
            cellBottomPressWringer4Info.MergeRight = 2;
            cellBottomPressWringer4Info.VerticalAlignment = VerticalAlignment.Center;
            cellBottomPressWringer4Info.Shading.Color = Colors.LightBlue;
            cellBottomPressWringer4Info.AddParagraph($"{GetClothingInfo(clothingBottomPressWringer4)}");
            #endregion

            #endregion // Bottom Press Wringer Position

            #region Top Press Wringer Position
            // add spacer row
            spacer = weeklyPMTable.AddRow();
            spacer.Height = 12;
            cellSpacer = spacer.Cells[0];
            cellSpacer.MergeRight = 12;
            cellSpacer.AddParagraph("");

            var clothingTopPressWringer1 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "top press wringer" && c.PM_Number == 1);
            var clothingTopPressWringer2 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "top press wringer" && c.PM_Number == 2);
            var clothingTopPressWringer3 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "top press wringer" && c.PM_Number == 3);
            var clothingTopPressWringer4 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "top press wringer" && c.PM_Number == 4);
            var rowTopPressWringer = weeklyPMTable.AddRow();
            rowTopPressWringer.Height = 10;
            var rowTopPressWringer2 = weeklyPMTable.AddRow();
            rowTopPressWringer2.Height = 10;

            var cellTopPressWringer = rowTopPressWringer.Cells[0];
            cellTopPressWringer.Format.Font.Size = 8;
            cellTopPressWringer.VerticalAlignment = VerticalAlignment.Center;
            cellTopPressWringer.MergeDown = 1;
            cellTopPressWringer.AddParagraph("TOP PRESS WRINGER");

            #region Top Press Wringer PM 1
            // ----- Top Press Wringer Position PM 1 ----- //
            var cellTopPressWringer1Past = rowTopPressWringer.Cells[1];
            cellTopPressWringer1Past.Format.Font.Size = 8;
            cellTopPressWringer1Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(1, 34);
            if (pastAverage > 0)
                cellTopPressWringer1Past.AddParagraph($"{pastAverage}");

            var cellTopPressWringer1Current = rowTopPressWringer.Cells[2];
            cellTopPressWringer1Current.Format.Font.Size = 8;
            cellTopPressWringer1Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothingTopPressWringer1);
            positionGoal = GetPositionGoal(1, 34);
            if (currentAge > 0)
                cellTopPressWringer1Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cellTopPressWringer1Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cellTopPressWringer1Current.Shading.Color = Colors.Yellow;
            }


            var cellTopPressWringer1Goal = rowTopPressWringer.Cells[3];
            cellTopPressWringer1Goal.Format.Font.Size = 8;
            cellTopPressWringer1Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cellTopPressWringer1Goal.AddParagraph($"{positionGoal}");

            var cellTopPressWringer1Info = rowTopPressWringer2.Cells[1];
            cellTopPressWringer1Info.Format.Font.Size = 8;
            cellTopPressWringer1Info.MergeRight = 2;
            cellTopPressWringer1Info.VerticalAlignment = VerticalAlignment.Center;
            cellTopPressWringer1Info.Shading.Color = Colors.LightBlue;
            cellTopPressWringer1Info.AddParagraph($"{GetClothingInfo(clothingTopPressWringer1)}");
            #endregion

            #region Top Press Wringer PM 2
            // ----- Top Press Wringer Position PM 2 ----- //
            var cellTopPressWringer2Past = rowTopPressWringer.Cells[4];
            cellTopPressWringer2Past.Format.Font.Size = 8;
            cellTopPressWringer2Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(2, 34);
            if (pastAverage > 0)
                cellTopPressWringer2Past.AddParagraph($"{pastAverage}");

            var cellTopPressWringer2Current = rowTopPressWringer.Cells[5];
            cellTopPressWringer2Current.Format.Font.Size = 8;
            cellTopPressWringer2Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothingTopPressWringer2);
            positionGoal = GetPositionGoal(2, 34);
            if (currentAge > 0)
                cellTopPressWringer2Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cellTopPressWringer2Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cellTopPressWringer2Current.Shading.Color = Colors.Yellow;
            }


            var cellTopPressWringer2Goal = rowTopPressWringer.Cells[6];
            cellTopPressWringer2Goal.Format.Font.Size = 8;
            cellTopPressWringer2Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cellTopPressWringer2Goal.AddParagraph($"{positionGoal}");

            var cellTopPressWringer2Info = rowTopPressWringer2.Cells[4];
            cellTopPressWringer2Info.Format.Font.Size = 8;
            cellTopPressWringer2Info.MergeRight = 2;
            cellTopPressWringer2Info.VerticalAlignment = VerticalAlignment.Center;
            cellTopPressWringer2Info.Shading.Color = Colors.LightBlue;
            cellTopPressWringer2Info.AddParagraph($"{GetClothingInfo(clothingTopPressWringer2)}");
            #endregion

            #region Top Press Wringer PM 3
            // ----- Top Press Wringer Position PM 3 ----- //
            var cellTopPressWringer3Past = rowTopPressWringer.Cells[7];
            cellTopPressWringer3Past.Format.Font.Size = 8;
            cellTopPressWringer3Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(3, 34);
            if (pastAverage > 0)
                cellTopPressWringer3Past.AddParagraph($"{pastAverage}");

            var cellTopPressWringer3Current = rowTopPressWringer.Cells[8];
            cellTopPressWringer3Current.Format.Font.Size = 8;
            cellTopPressWringer3Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothingTopPressWringer3);
            positionGoal = GetPositionGoal(3, 34);
            if (currentAge > 0)
                cellTopPressWringer3Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cellTopPressWringer3Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cellTopPressWringer3Current.Shading.Color = Colors.Yellow;
            }


            var cellTopPressWringer3Goal = rowTopPressWringer.Cells[9];
            cellTopPressWringer3Goal.Format.Font.Size = 8;
            cellTopPressWringer3Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cellTopPressWringer3Goal.AddParagraph($"{positionGoal}");

            var cellTopPressWringer3Info = rowTopPressWringer2.Cells[7];
            cellTopPressWringer3Info.Format.Font.Size = 8;
            cellTopPressWringer3Info.MergeRight = 2;
            cellTopPressWringer3Info.VerticalAlignment = VerticalAlignment.Center;
            cellTopPressWringer3Info.Shading.Color = Colors.LightBlue;
            cellTopPressWringer3Info.AddParagraph($"{GetClothingInfo(clothingTopPressWringer3)}");
            #endregion

            #region Top Press Wringer PM 4
            // ----- Top Press Wringer Position PM 4 ----- //
            var cellTopPressWringer4Past = rowTopPressWringer.Cells[10];
            cellTopPressWringer4Past.Format.Font.Size = 8;
            cellTopPressWringer4Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(4, 34);
            if (pastAverage > 0)
                cellTopPressWringer4Past.AddParagraph($"{pastAverage}");

            var cellTopPressWringer4Current = rowTopPressWringer.Cells[11];
            cellTopPressWringer4Current.Format.Font.Size = 8;
            cellTopPressWringer4Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothingTopPressWringer4);
            positionGoal = GetPositionGoal(4, 34);
            if (currentAge > 0)
                cellTopPressWringer4Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cellTopPressWringer4Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cellTopPressWringer4Current.Shading.Color = Colors.Yellow;
            }


            var cellTopPressWringer4Goal = rowTopPressWringer.Cells[12];
            cellTopPressWringer4Goal.Format.Font.Size = 8;
            cellTopPressWringer4Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cellTopPressWringer4Goal.AddParagraph($"{positionGoal}");

            var cellTopPressWringer4Info = rowTopPressWringer2.Cells[10];
            cellTopPressWringer4Info.Format.Font.Size = 8;
            cellTopPressWringer4Info.MergeRight = 2;
            cellTopPressWringer4Info.VerticalAlignment = VerticalAlignment.Center;
            cellTopPressWringer4Info.Shading.Color = Colors.LightBlue;
            cellTopPressWringer4Info.AddParagraph($"{GetClothingInfo(clothingTopPressWringer4)}");
            #endregion

            #endregion // Top Press Wringer Position

            #region Couch Paper Roll Position
            // add spacer row
            spacer = weeklyPMTable.AddRow();
            spacer.Height = 12;
            cellSpacer = spacer.Cells[0];
            cellSpacer.MergeRight = 12;
            cellSpacer.AddParagraph("");

            var clothingCouchPaper1 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "couch paper roll" && c.PM_Number == 1);
            var clothingCouchPaper2 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "couch paper roll" && c.PM_Number == 2);
            var clothingCouchPaper3 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "couch paper roll" && c.PM_Number == 3);
            var clothingCouchPaper4 = clothings.SingleOrDefault(c => c.Position.Position1.ToLower() == "couch paper roll" && c.PM_Number == 4);
            var rowCouchPaper = weeklyPMTable.AddRow();
            rowCouchPaper.Height = 10;
            var rowCouchPaper2 = weeklyPMTable.AddRow();
            rowCouchPaper2.Height = 10;

            var cellCouchPaper = rowCouchPaper.Cells[0];
            cellCouchPaper.Format.Font.Size = 8;
            cellCouchPaper.VerticalAlignment = VerticalAlignment.Center;
            cellCouchPaper.MergeDown = 1;
            cellCouchPaper.AddParagraph("COUCH PAPER ROLL");

            #region Couch Paper PM 1
            // ----- Couch Paper Roll Position PM 1 ----- //
            var cellCouchPaper1Past = rowCouchPaper.Cells[1];
            cellCouchPaper1Past.Format.Font.Size = 8;
            cellCouchPaper1Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(1, 36);
            if (pastAverage > 0)
                cellCouchPaper1Past.AddParagraph($"{pastAverage}");

            var cellCouchPaper1Current = rowCouchPaper.Cells[2];
            cellCouchPaper1Current.Format.Font.Size = 8;
            cellCouchPaper1Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothingCouchPaper1);
            positionGoal = GetPositionGoal(1, 36);
            if (currentAge > 0)
                cellCouchPaper1Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cellCouchPaper1Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cellCouchPaper1Current.Shading.Color = Colors.Yellow;
            }


            var cellCouchPaper1Goal = rowCouchPaper.Cells[3];
            cellCouchPaper1Goal.Format.Font.Size = 8;
            cellCouchPaper1Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cellCouchPaper1Goal.AddParagraph($"{positionGoal}");

            var cellCouchPaper1Info = rowCouchPaper2.Cells[1];
            cellCouchPaper1Info.Format.Font.Size = 8;
            cellCouchPaper1Info.MergeRight = 2;
            cellCouchPaper1Info.VerticalAlignment = VerticalAlignment.Center;
            cellCouchPaper1Info.Shading.Color = Colors.LightBlue;
            cellCouchPaper1Info.AddParagraph($"{GetClothingInfo(clothingCouchPaper1)}");
            #endregion

            #region Couch Paper PM 2
            // ----- Couch Paper Roll Position PM 2 ----- //
            var cellCouchPaper2Past = rowCouchPaper.Cells[4];
            cellCouchPaper2Past.Format.Font.Size = 8;
            cellCouchPaper2Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(2, 36);
            if (pastAverage > 0)
                cellCouchPaper2Past.AddParagraph($"{pastAverage}");

            var cellCouchPaper2Current = rowCouchPaper.Cells[5];
            cellCouchPaper2Current.Format.Font.Size = 8;
            cellCouchPaper2Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothingCouchPaper2);
            positionGoal = GetPositionGoal(2, 36);
            if (currentAge > 0)
                cellCouchPaper2Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cellCouchPaper2Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cellCouchPaper2Current.Shading.Color = Colors.Yellow;
            }


            var cellCouchPaper2Goal = rowCouchPaper.Cells[6];
            cellCouchPaper2Goal.Format.Font.Size = 8;
            cellCouchPaper2Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cellCouchPaper2Goal.AddParagraph($"{positionGoal}");

            var cellCouchPaper2Info = rowCouchPaper2.Cells[4];
            cellCouchPaper2Info.Format.Font.Size = 8;
            cellCouchPaper2Info.MergeRight = 2;
            cellCouchPaper2Info.VerticalAlignment = VerticalAlignment.Center;
            cellCouchPaper2Info.Shading.Color = Colors.LightBlue;
            cellCouchPaper2Info.AddParagraph($"{GetClothingInfo(clothingCouchPaper2)}");
            #endregion

            #region Couch Paper PM 3
            // ----- Couch Paper Roll Position PM 3 ----- //
            var cellCouchPaper3Past = rowCouchPaper.Cells[7];
            cellCouchPaper3Past.Format.Font.Size = 8;
            cellCouchPaper3Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(3, 36);
            if (pastAverage > 0)
                cellCouchPaper3Past.AddParagraph($"{pastAverage}");

            var cellCouchPaper3Current = rowCouchPaper.Cells[8];
            cellCouchPaper3Current.Format.Font.Size = 8;
            cellCouchPaper3Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothingCouchPaper3);
            positionGoal = GetPositionGoal(3, 36);
            if (currentAge > 0)
                cellCouchPaper3Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cellCouchPaper3Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cellCouchPaper3Current.Shading.Color = Colors.Yellow;
            }


            var cellCouchPaper3Goal = rowCouchPaper.Cells[9];
            cellCouchPaper3Goal.Format.Font.Size = 8;
            cellCouchPaper3Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cellCouchPaper3Goal.AddParagraph($"{positionGoal}");

            var cellCouchPaper3Info = rowCouchPaper2.Cells[7];
            cellCouchPaper3Info.Format.Font.Size = 8;
            cellCouchPaper3Info.MergeRight = 2;
            cellCouchPaper3Info.VerticalAlignment = VerticalAlignment.Center;
            cellCouchPaper3Info.Shading.Color = Colors.LightBlue;
            cellCouchPaper3Info.AddParagraph($"{GetClothingInfo(clothingCouchPaper3)}");
            #endregion

            #region Couch Paper PM 4
            // ----- Couch Paper Roll Position PM 4 ----- //
            var cellCouchPaper4Past = rowCouchPaper.Cells[10];
            cellCouchPaper4Past.Format.Font.Size = 8;
            cellCouchPaper4Past.VerticalAlignment = VerticalAlignment.Center;
            pastAverage = GetPastAverage(4, 36);
            if (pastAverage > 0)
                cellCouchPaper4Past.AddParagraph($"{pastAverage}");

            var cellCouchPaper4Current = rowCouchPaper.Cells[11];
            cellCouchPaper4Current.Format.Font.Size = 8;
            cellCouchPaper4Current.VerticalAlignment = VerticalAlignment.Center;
            currentAge = GetCurrentAge(clothingCouchPaper4);
            positionGoal = GetPositionGoal(4, 36);
            if (currentAge > 0)
                cellCouchPaper4Current.AddParagraph($"{currentAge}");
            if (currentAge > 0 || positionGoal > 0)
            {
                if (currentAge >= positionGoal)
                    cellCouchPaper4Current.Shading.Color = Colors.Red;
                else if (positionGoal - currentAge <= 35)
                    cellCouchPaper4Current.Shading.Color = Colors.Yellow;
            }
            

            var cellCouchPaper4Goal = rowCouchPaper.Cells[12];
            cellCouchPaper4Goal.Format.Font.Size = 8;
            cellCouchPaper4Goal.VerticalAlignment = VerticalAlignment.Center;
            if (positionGoal > 0)
                cellCouchPaper4Goal.AddParagraph($"{positionGoal}");

            var cellCouchPaper4Info = rowCouchPaper2.Cells[10];
            cellCouchPaper4Info.Format.Font.Size = 8;
            cellCouchPaper4Info.MergeRight = 2;
            cellCouchPaper4Info.VerticalAlignment = VerticalAlignment.Center;
            cellCouchPaper4Info.Shading.Color = Colors.LightBlue;
            cellCouchPaper4Info.AddParagraph($"{GetClothingInfo(clothingCouchPaper4)}");
            #endregion

            #endregion // Couch Paper Roll Position

            return weeklyPMTable;
        }

        private static double GetPastAverage(int pmID, int posID)
        {
            var averageAge = 0.0;
            var average = db.Clothings.Where(c => c.PM_Number == pmID && c.PositionID == posID && c.StatusID == 3).OrderByDescending(d => d.Date_Removed_From_Mac).FirstOrDefault();
            if (average != null)
            {
                if (average.Age != null)
                    averageAge = Math.Round((double)average.Age);
            }
            return averageAge;
        }

        private static int GetCurrentAge(Clothing clothing)
        {
            if (clothing != null)
            {
                if (clothing.Age > 0 && clothing.StatusID != 2)
                {
                    return Convert.ToInt32(clothing.Age);
                }
                else if (clothing.Date_Placed_On_Mac != null)
                {
                    var age = DateTime.Now - (DateTime)clothing.Date_Placed_On_Mac;
                    return age.Days;
                }
            }
            return 0;
        }

        private static int GetPositionGoal(int pmID, int posID)
        {
            if (pmID > 0 && posID > 0)
            {
                var goal = db.Goals.SingleOrDefault(g => g.PM_ID == pmID && g.PositionID == posID);
                return Convert.ToInt32(goal.Goal1);
            }
            return 0;
        }

        private static string GetClothingInfo(Clothing clothing)
        {
            if (clothing != null)
            {
                if (clothing.Manufacturer != null)
                    return $"{clothing.Manufacturer.Name} {clothing.Serial_Number}";
                else return $"NA {clothing.Serial_Number}";
            }
            return "NA";

        }
    }
}