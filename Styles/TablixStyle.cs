using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace INC.RDLCBuilder.Styles
{
    public class TablixStyle
    {
        public TablixStyle()
        {
            this.HeaderHight = 0.8m;
            this.RowHight = 0.8m;
            this.Columns = new List<ColumnStyle>();
        }

        public string Name { get; set; }

        public string DataSetName { get; set; }

        public decimal HeaderHight { get; set; }

        public decimal RowHight { get; set; }

        public List<ColumnStyle> Columns { get; set; }

        public PositionStyle Position { get; set; }
    }
}