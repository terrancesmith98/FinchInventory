namespace FinchInventory.Models
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    public partial class Goal
    {
        public int ID { get; set; }

        public int PositionID { get; set; }

        public int PM_ID { get; set; }

        [Column("Goal")]
        public int? Goal1 { get; set; }

        public virtual Machine Machine { get; set; }

        public virtual Position Position { get; set; }
    }
}
