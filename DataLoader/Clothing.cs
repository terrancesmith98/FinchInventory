namespace DataLoader
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("Clothing")]
    public partial class Clothing
    {
        public int ID { get; set; }

        public int PM_Number { get; set; }

        public int PositionID { get; set; }

        public int TypeID { get; set; }

        [Required]
        [StringLength(50)]
        public string Serial_Number { get; set; }

        [Column(TypeName = "date")]
        public DateTime? Date_Received { get; set; }

        [Column(TypeName = "date")]
        public DateTime? Date_Placed_On_Mac { get; set; }

        [Column(TypeName = "date")]
        public DateTime? Date_Removed_From_Mac { get; set; }

        public int StatusID { get; set; }

        public int? LocationID { get; set; }

        public string Comments { get; set; }

        public int? Age { get; set; }

        [StringLength(50)]
        public string Dimensions { get; set; }

        public int RollTypeID { get; set; }

        public int? RollWeight { get; set; }

        public decimal? CurrentDia { get; set; }

        public decimal? MinDia { get; set; }

        public decimal? Crown { get; set; }

        [StringLength(50)]
        public string CoverMaterial { get; set; }

        [StringLength(50)]
        public string HoleGroovePattern { get; set; }

        public int? SpecifiedHardness { get; set; }

        public int? MeasuredHardness { get; set; }

        public int? SpecifiedRa { get; set; }

        public int? MeasuredRa { get; set; }

        [Column(TypeName = "date")]
        public DateTime? CoverDate { get; set; }

        public int? ManufacturerID { get; set; }

        public virtual Location Location { get; set; }

        public virtual Machine Machine { get; set; }

        public virtual Manufacturer Manufacturer { get; set; }

        public virtual Position Position { get; set; }

        public virtual RollType RollType { get; set; }

        public virtual Status Status { get; set; }

        public virtual Type Type { get; set; }
    }
}
