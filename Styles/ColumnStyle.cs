using INC.RDLCBuilder.Enums;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Web;

namespace INC.RDLCBuilder.Styles
{
    public class ColumnStyle
    {
        private string _name = string.Empty;

        public ColumnStyle()
        {
            this.TextAlign = TextAlign.Center;
        }


        public string Name
        {
            get
            {
                return this._name;
            }
            set
            {
                this._name = value;
            }
        }

        public decimal Width { get; set; }

        public TextAlign TextAlign { get; set; }

        public string DataType { get; set; }

        public string Format { get; set; }
    } 
}