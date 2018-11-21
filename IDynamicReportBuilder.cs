using INC.RDLCBuilder.Enums;
using INC.RDLCBuilder.Styles;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace INC.RDLCBuilder
{
    public interface IDynamicReportBuilder
    {
        IDynamicReportBuilder SetHeader(HeaderStyle style);

        IDynamicReportBuilder SetBody(BodyStyle style);

        IDynamicReportBuilder SetFooter(FooterStyle style);

        IDynamicReportBuilder SetPage(PageStyle page, PageOrtientation pageOrtientation);

        IDynamicReportBuilder AddTablix(TablixStyle style, Position position);

        System.IO.Stream Build();
    }
}
