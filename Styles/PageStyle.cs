using INC.RDLCBuilder.Enums;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace INC.RDLCBuilder.Styles
{
    public class PageStyle
    {
        public static List<KeyValuePair<Page, PageStyle>> PAGE_STYLE = new List<KeyValuePair<Page, PageStyle>>()
        {
            new KeyValuePair<Page, PageStyle>(Page.A3, new PageStyle(42m,29.7m)),
            new KeyValuePair<Page, PageStyle>(Page.A4, new PageStyle(29.7m,21m)),
        };

        public PageStyle(decimal height, decimal width)
        {
            this.Height = height;
            this.Width = width;
        }

        public decimal Height { get; set; }

        public decimal Width { get; set; }

        public static PageStyle GetPageStyle(Page page)
        {
            return PAGE_STYLE.Find(x => x.Key == page).Value;
        }
    }
}