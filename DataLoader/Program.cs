using System;
using System.IO;

namespace DataLoader
{
    class Program
    {
        private static FinchContext db = new FinchContext();
        static void Main(string[] args)
        {
            var filepath = @"C:\Users\tsmith.OTI\Documents\clients\finch\clothing2.csv";
            var lines = File.ReadAllLines(filepath);

            foreach (var line in lines)
            {
                if (line != lines[0])
                {
                    Clothing clothing = new Clothing();
                    var fields = line.Split(',');

                    clothing.PM_Number = Convert.ToInt32(fields[0]);
                    clothing.RollTypeID = GetRollTypeID(fields[1].Trim());
                    clothing.PositionID = GetPositionID(fields[2].Trim());
                    clothing.Dimensions = fields[3].Trim();
                    clothing.Serial_Number = fields[4].Trim();
                    clothing.Manufacturer.Name = fields[5].Trim();
                    clothing.TypeID = GetTypeID(fields[6].Trim());
                    clothing.RollWeight = !string.IsNullOrEmpty(fields[7]) ? Convert.ToInt32(fields[7].Trim()) : 0;
                    clothing.CurrentDia = !string.IsNullOrEmpty(fields[8]) ? Convert.ToDecimal(fields[8].Trim()) : 0;
                    clothing.MinDia = !string.IsNullOrEmpty(fields[9]) ? Convert.ToDecimal(fields[9].Trim()) : 0;
                    clothing.Crown = !string.IsNullOrEmpty(fields[10]) ? Convert.ToDecimal(fields[10].Trim()) : 0;
                    clothing.CoverMaterial = fields[11].Trim();
                    clothing.HoleGroovePattern = fields[12].Trim();
                    clothing.SpecifiedHardness = !string.IsNullOrEmpty(fields[13]) ? Convert.ToInt32(fields[13].Trim()) : 0;
                    clothing.MeasuredHardness = !string.IsNullOrEmpty(fields[14]) ? Convert.ToInt32(fields[14].Trim()) : 0;
                    clothing.SpecifiedRa = !string.IsNullOrEmpty(fields[15]) ? Convert.ToInt32(fields[15].Trim()) : 0;
                    clothing.MeasuredRa = !string.IsNullOrEmpty(fields[16]) ? Convert.ToInt32(fields[16].Trim()) : 0;
                    clothing.CoverDate = !string.IsNullOrEmpty(fields[17]) ? DateTime.Parse(fields[17].Trim()) : DateTime.Parse("2019/01/01");
                    clothing.Date_Received = !string.IsNullOrEmpty(fields[18]) ? DateTime.Parse(fields[18].Trim()) : DateTime.Parse("2019/01/01");
                    clothing.Date_Received = !string.IsNullOrEmpty(fields[19]) ? DateTime.Parse(fields[19].Trim()) : DateTime.Parse("2019/01/01");
                    clothing.Date_Removed_From_Mac = !string.IsNullOrEmpty(fields[20]) ? DateTime.Parse(fields[20].Trim()) : DateTime.Parse("2019/01/01");
                    clothing.Age = !string.IsNullOrEmpty(fields[21]) ? Convert.ToInt32(fields[21]) : 0;
                    clothing.StatusID = GetStatusID(fields[22].Trim());
                    clothing.LocationID = GetLocationID(fields[23].Trim());
                    clothing.Comments = fields[24].Trim();


                    Console.WriteLine($"Added {clothing.Serial_Number} Type: {fields[6].Trim()}");
                    db.Clothings.Add(clothing);
                    db.SaveChanges();
                }



            }

        }

        private static int GetRollTypeID(string rollType)
        {
            var rollTypeID = 0;
            switch (rollType.ToLower())
            {
                case "roll":
                    rollTypeID = 1;
                    break;
                case "clothing":
                    rollTypeID = 2;
                    break;
                default:
                    break;
            }

            return rollTypeID;
        }

        private static int GetLocationID(string location)
        {
            var locationID = 0;
            switch (location.ToLower())
            {
                case "bsmt de rack":
                    locationID = 1;
                    break;
                case "bsmt we rack":
                    locationID = 2;
                    break;
                case "c-gerneration seam":
                case "c-generation seam":
                    locationID = 3;
                    break;
                case "on floor basement":
                    locationID = 4;
                    break;
                case "rack 2":
                    locationID = 5;
                    break;
                case "rack 3":
                    locationID = 6;
                    break;
                case "#4 pulp prep rack":
                    locationID = 7;
                    break;
                case "pulp prep":
                    locationID = 9;
                    break;
                default:
                    locationID = 8;
                    break;
            }

            return locationID;
        }

        private static int GetStatusID(string status)
        {
            var statusID = 0;
            switch (status.ToLower())
            {
                case "inventory":
                    statusID = 1;
                    break;
                case "on machine":
                    statusID = 2;
                    break;
                case "history":
                    statusID = 3;
                    break;
                default:
                    statusID = 1;
                    break;
            }

            return statusID;

        }

        private static int GetTypeID(string type)
        {
            var typeID = 0;
            switch (type.ToLower())
            {
                case "aerorun":
                    typeID = 1;
                    break;
                case "dryline q590":
                    typeID = 2;
                    break;
                case "multiseam":
                    typeID = 3;
                    break;
                case "prineline lc":
                case "printline lc":
                    typeID = 4;
                    break;
                case "seam apertech":
                    typeID = 5;
                    break;
                case "thermonetics":
                    typeID = 6;
                    break;
                case "thermonetics 500":
                    typeID = 7;
                    break;
                case "ultra 5001":
                    typeID = 8;
                    break;
                case "printflex v3l":
                    typeID = 9;
                    break;
                case "printex xd":
                    typeID = 10;
                    break;
                case "voith":
                    typeID = 11;
                    break;
                case "seamtech plus ii":
                    typeID = 12;
                    break;
                case "aeropulse xl":
                    typeID = 13;
                    break;
                case "printex r501":
                    typeID = 14;
                    break;
                case "seam apertech 300":
                    typeID = 1011;
                    break;
                case "printex q209, edge plus":
                    typeID = 1012;
                    break;
                case "printflex as3 g2":
                    typeID = 1013;
                    break;
                case "seamtech plus 300":
                    typeID = 1014;
                    break;
                case "seamtech ii reg":
                    typeID = 1016;
                    break;
                case "aerorun 150, t40":
                    typeID = 1017;
                    break;
                case "aerorun 300, t40":
                    typeID = 1018;
                    break;
                case "aerorun 400, t40":
                    typeID = 1019;
                    break;
                case "thermonetics xxl 400, t40":
                    typeID = 1020;
                    break;
                case "thermonetics sym 300, t40":
                    typeID = 1021;
                    break;
                case "printex r209":
                    typeID = 1022;
                    break;
                case "printflex as3":
                    typeID = 1023;
                    break;
                case "seamtech plus":
                    typeID = 1024;
                    break;
                default:
                    typeID = 15;
                    break;
            }

            return typeID;
        }

        private static int GetPositionID(string position)
        {
            var positionID = 0;
            switch (position.ToLower())
            {
                case "wire":
                    positionID = 2;
                    break;
                case "1st press felt":
                    positionID = 3;
                    break;
                case "2nd press felt":
                    positionID = 4;
                    break;
                case "3rd press felt":
                    positionID = 5;
                    break;
                case "1st top dryer felt":
                    positionID = 6;
                    break;
                case "2nd top dryer felt":
                    positionID = 7;
                    break;
                case "3rd top dryer felt":
                    positionID = 8;
                    break;
                case "4th top dryer felt":
                    positionID = 9;
                    break;
                case "1st bottom dryer felt":
                    positionID = 10;
                    break;
                case "2nd bottom dryer felt":
                    positionID = 11;
                    break;
                case "3rd bottom dryer felt":
                    positionID = 12;
                    break;
                case "4th bottom dryer felt":
                    positionID = 13;
                    break;
                case "breast roll":
                    positionID = 14;
                    break;
                case "dandy roll":
                    positionID = 15;
                    break;
                case "lumpbreaker roll":
                    positionID = 16;
                    break;
                case "suction pickup roll":
                    positionID = 17;
                    break;
                case "1st press top roll":
                    positionID = 18;
                    break;
                case "1st press bottom roll":
                    positionID = 19;
                    break;
                case "2nd press top roll":
                    positionID = 20;
                    break;
                case "2nd press bottom roll":
                    positionID = 21;
                    break;
                case "3rd press top roll":
                    positionID = 22;
                    break;
                case "3rd press bottom roll":
                    positionID = 23;
                    break;
                case "smoother top roll":
                    positionID = 24;
                    break;
                case "smoother bottom roll":
                    positionID = 25;
                    break;
                case "hard size roll":
                    positionID = 26;
                    break;
                case "soft size roll":
                    positionID = 27;
                    break;
                case "aquatherm roll":
                    positionID = 28;
                    break;
                case "nibco roll":
                    positionID = 29;
                    break;
                case "couch roll":
                    positionID = 30;
                    break;
                case "l_in hope roll":
                    positionID = 31;
                    break;
                case "l_out hope roll":
                    positionID = 32;
                    break;
                case "bottom press wringer":
                    positionID = 33;
                    break;
                case "top press wringer":
                    positionID = 34;
                    break;
                case "couch paper roll":
                    positionID = 36;
                    break;
                default:
                    positionID = 35;
                    break;
            }

            return positionID;
        }
    }
}
