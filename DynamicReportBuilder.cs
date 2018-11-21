using INC.RDLCBuilder.Enums;
using INC.RDLCBuilder.Styles;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Web;
using System.Xml;
using System.Xml.Serialization;

namespace INC.RDLCBuilder
{
    public class DynamicReportBuilder : IDynamicReportBuilder
    {
        #region Constant

        protected string _unit = "cm";

        protected const string _docTemplate=
            "<?xml version=\"1.0\" encoding=\"utf-8\"?><Report xmlns:rd=\"http://schemas.microsoft.com/SQLServer/reporting/reportdesigner\" xmlns=\"http://schemas.microsoft.com/sqlserver/reporting/2008/01/reportdefinition\">" +
            "<DataSources>" +
            "   <DataSource Name=\"ReportDataSource\">" +
            "       <ConnectionProperties>" +
            "           <DataProvider>SQL</DataProvider>" +
            "           <ConnectString />" +
            "       </ConnectionProperties>" +
            "       <rd:DataSourceID>3eecdab9-6b4b-4836-ad62-95e4aee65ea8</rd:DataSourceID>" +
            "    </DataSource>" +
            "</DataSources>" +
            "<DataSets>@DataSets</DataSets>" +
            "<Body>@PageBody</Body>" +
            "<Width>@BodyWidth</Width>" +
            "<Page>" +
            "   @PageHeader"+
            "   @PageFooter" +
            "   @PageStyle "+
            "</Page>" +
            "<rd:ReportID>809f16cf-ea78-4469-bf43-965c4afe69d0</rd:ReportID>" +
            "<rd:ReportUnitType>Cm</rd:ReportUnitType>" +
            "</Report>";

        protected const string _dataSetPattern = 
            "    <DataSet Name=\"@DataSetName\">" +
            "       <Fields>@Fields</Fields>" +
            "       <Query>" +
            "           <DataSourceName>ReportDataSource</DataSourceName>" +
            "           <CommandText />" +
            "       </Query>" +
            "    </DataSet>";

        protected const string _tablixPattern =
            " <Tablix Name=\"Tablix@DataSetNameGuid\">" +
            "   <TablixBody>" +
            "       <TablixColumns>@TablixColumns</TablixColumns>" +
            "       <TablixRows>" +
            "           <TablixRow>" +
            "               <Height>@RowHeight</Height>" +
            "               <TablixCells>@TablixHeader</TablixCells>" +
            "           </TablixRow>" +
            "           <TablixRow>" +
            "               <Height>@RowHeight</Height>" +
            "               <TablixCells>@TablixCells</TablixCells>" +
            "           </TablixRow>" +
            "       </TablixRows>" +
            "   </TablixBody>" +
            "   <TablixColumnHierarchy>" +
            "       <TablixMembers>@TablixMember</TablixMembers>" +
            "   </TablixColumnHierarchy>" +
            "   <TablixRowHierarchy>" +
            "       <TablixMembers>" +
            "           <TablixMember>" +
            "               <FixedData>true</FixedData>" +
            "               <KeepWithGroup>After</KeepWithGroup>" +
            "               <RepeatOnNewPage>true</RepeatOnNewPage>" +
            "           </TablixMember>" +
            "           <TablixMember>" +
            "               <Group Name=\"Detail@DataSetName\" />" +
            "           </TablixMember>" +
            "       </TablixMembers>" +
            "   </TablixRowHierarchy>" +
            "   <RepeatRowHeaders>true</RepeatRowHeaders>"+
            "   <RepeatColumnHeaders>true</RepeatColumnHeaders>" +
            "   <DataSetName>@DataSetName</DataSetName>" +
            "   <Top>@Top</Top>" +
            "   <Left>@Left</Left>" +
            "   <Height>@Height</Height>" +
            "   <Width>@Width</Width>" +
            "   <Style>" +
            "       <Border>" +
            "           <Style>None</Style>" +
            "       </Border>" +
            "   </Style>" +
            "</Tablix>";


        protected const string _pageHeaderFooterContentPattern =
                   "<Height>@Height</Height>" +
                   "<PrintOnFirstPage>@PrintOnFirstPage</PrintOnFirstPage>" +
                   "<PrintOnLastPage>@PrintOnLastPage</PrintOnLastPage>" +
                   "<ReportItems>@ReportItems</ReportItems>" +
                   "<Style>" +
                   "   <Border>" +
                   "       <Style>None</Style>" +
                   "   </Border>" +
                   "</Style>";

        protected const string _pageBodyContentPattern =
                "<ReportItems>@ReportItems</ReportItems>" +
                "<Style />" +
                "<Height>@Height</Height>";

        protected const string _pageStylePattern =
            "   <PageHeight>@PageHeight</PageHeight>" +
            "   <PageWidth>@PageWidth</PageWidth>" +
            "   <LeftMargin>2cm</LeftMargin>" +
            "   <RightMargin>2cm</RightMargin>" +
            "   <TopMargin>2cm</TopMargin>" +
            "   <BottomMargin>2cm</BottomMargin>" +
            "   <ColumnSpacing>0.13cm</ColumnSpacing>" +
            "   <Style />";

        #endregion



        #region Private Parmaters

        private HeaderStyle _headerStyle;
        private BodyStyle _bodyStyle;
        private FooterStyle _footerStyle;
        private PageStyle _pageStyle = PageStyle.GetPageStyle(Page.A4);
        private PageOrtientation _pageOrtientation = PageOrtientation.Portrait;
        private Dictionary<Position, List<TablixStyle>> _tablixDic;

        private int _fieldSymbio = 0;
        private Dictionary<string, List<KeyValuePair<string, string>>> _dataSetFieldMap = new Dictionary<string, List<KeyValuePair<string, string>>>();
        #endregion

        #region Construct

        public DynamicReportBuilder() 
        {
            _tablixDic = new Dictionary<Position, List<TablixStyle>>();
            _tablixDic.Add(Position.Header, new List<TablixStyle>());
            _tablixDic.Add(Position.Footer, new List<TablixStyle>());
            _tablixDic.Add(Position.Body, new List<TablixStyle>());
        }

        #endregion

        #region Implementation

        public IDynamicReportBuilder SetHeader(Styles.HeaderStyle style)
        {
            this._headerStyle = style;
            return this;
        }

        public IDynamicReportBuilder SetBody(Styles.BodyStyle style)
        {
            this._bodyStyle = style;
            return this;
        }

        public IDynamicReportBuilder SetFooter(Styles.FooterStyle style)
        {
            this._footerStyle = style;
            return this;
        }

        public IDynamicReportBuilder AddTablix(Styles.TablixStyle style, Enums.Position position)
        {
            var orignal = _tablixDic[position];
            orignal.Add(style);
            _tablixDic[position] = orignal;
            return this;
        }

        public IDynamicReportBuilder SetPage(PageStyle pageStyle, PageOrtientation pageOrtientation)
        {
            this._pageStyle = pageStyle;
            this._pageOrtientation = pageOrtientation;
            return this;
        }

        public Stream Build()
        {
            var reportContent = _docTemplate
                .Replace("@PageHeader", BuildHeader(this._headerStyle))
                .Replace("@PageFooter", BuildFooter(this._footerStyle))
                .Replace("@PageStyle", BuildPageStyle(this._pageStyle, this._pageOrtientation))
                .Replace("@DataSets", BuildDataSet())
                .Replace("@PageBody", BuildBody(this._bodyStyle))
                .Replace("@BodyWidth", BuildBodyWidth(this._pageStyle, this._pageOrtientation));
                
            var doc = new XmlDocument();
            doc.LoadXml(reportContent);

            return GetRdlcStream(doc);
        }

        #endregion

        #region Private Method

        private string BuildHeader(HeaderStyle style) 
        {
            if (style != null)
            {
                return "<PageFooter>" + 
                    _pageHeaderFooterContentPattern
                       .Replace("@Height", style.Height.ToString() + this._unit)
                       .Replace("@PrintOnFirstPage", style.PrintOnFirstPage ? "true" : "false")
                       .Replace("@PrintOnLastPage", style.PrintOnFirstPage ? "true" : "false")
                       .Replace("@ReportItems", BuildReportItems(Position.Header)) +
                   "</PageFooter>";
            }
            else
            {
                return string.Empty;
            }
        }

        private string BuildFooter(FooterStyle style)
        {
            if (style != null)
            {
                return "<PageHeader>"+ 
                    _pageHeaderFooterContentPattern
                       .Replace("@Height", style.Height.ToString() + this._unit)
                       .Replace("@PrintOnFirstPage", style.PrintOnFirstPage ? "true" : "false")
                       .Replace("@PrintOnLastPage", style.PrintOnFirstPage ? "true" : "false")
                       .Replace("@ReportItems", BuildReportItems(Position.Footer)) + 
                   "</PageHeader>";
            }
            else
            {
                return string.Empty;
            }
        }

        public string BuildBody(BodyStyle style) 
        {    
            if (style != null)
            {
                return _pageBodyContentPattern
                    .Replace("@Height", style.Height.ToString() + this._unit)
                    .Replace("@ReportItems", BuildReportItems(Position.Body));
            }
            else
            {
                return string.Empty;
            }
        }

        public string BuildBodyWidth(PageStyle style, PageOrtientation pageOrtientation) 
        {
            if (_pageStyle == null)
            {
                throw new NullReferenceException("Page style must setting.");
            }
            else
            {
                var pageWidth = 0m;

                if (pageOrtientation == PageOrtientation.Portrait)
                {
                    pageWidth = style.Width;
                }
                else
                {
                    pageWidth = style.Height;
                }

                return (pageWidth - 4.0m).ToString() + _unit;
            }
        }

        public string BuildPageStyle(PageStyle style ,PageOrtientation pageOrtientation)
        {
            if (_pageStyle == null)
            {
                throw new NullReferenceException("Page style must setting.");
            }
            else { 
                var pageHeight=0m;
                var pageWidth=0m;

                if(pageOrtientation==PageOrtientation.Portrait){
                    pageHeight=style.Height;
                    pageWidth=style.Width;
                }else{
                    pageWidth=style.Height;
                    pageHeight=style.Width;
                }

                return _pageStylePattern
                    .Replace("@PageHeight", pageHeight.ToString() + _unit)
                    .Replace("@PageWidth", pageWidth.ToString() + _unit);
            }
        }

        #region BuildDataSet

        private string BuildDataSet() 
        {
            var dataset = new StringBuilder();
            var dataSetConstruct = new Dictionary<string, Dictionary<string,string>>();
            this._tablixDic.SelectMany(x => x.Value).ToList().ForEach(x =>
            {
                if (dataSetConstruct.ContainsKey(x.DataSetName))
                {
                    var existCols = dataSetConstruct[x.DataSetName];
                    foreach (var col in x.Columns)
                    {
                        if (!existCols.ContainsKey(col.Name))
                        {
                            existCols.Add(col.Name, col.DataType);
                        }
                    }

                    dataSetConstruct[x.DataSetName] = existCols;
                }
                else
                {
                    dataSetConstruct.Add(x.DataSetName, x.Columns.ToDictionary(y => y.Name, y => y.DataType));
                }
            });

            foreach (var item in dataSetConstruct)
            {
                var fieldMap=new List<KeyValuePair<string,string>>();
                var fields = new StringBuilder();
                foreach (var col in item.Value)
                {
                    _fieldSymbio += 1;
                    var colName = "Field" + _fieldSymbio.ToString();
                    fieldMap.Add(new KeyValuePair<string, string>(col.Key, "Field" + _fieldSymbio.ToString()));
                    fields.AppendFormat(
                       "<Field Name=\"{0}\"><DataField>{1}</DataField><rd:TypeName>{2}</rd:TypeName></Field>",
                       colName, col.Key, string.IsNullOrWhiteSpace(col.Value) ? "System.String" : col.Value);
                }

                dataset.Append(_dataSetPattern.Replace("@DataSetName", item.Key).Replace("@Fields", fields.ToString()));

                if (_dataSetFieldMap.ContainsKey(item.Key))
                {
                    _dataSetFieldMap[item.Key].AddRange(fieldMap);
                }
                else
                {
                    _dataSetFieldMap.Add(item.Key, fieldMap);
                }
            }

            return dataset.ToString();
        }

        #endregion
        
        #region BuildReportItems

        private string BuildReportItems(Position position)
        {
            var reportItems = new StringBuilder();
            reportItems.Append(BuildTablixs(position));

            return reportItems.ToString();
        }

        private string BuildTablixs(Position position) 
        {
            var tablixsSB = new StringBuilder();
            var tablixs = this._tablixDic[position];
            foreach (var item in tablixs)
            {
                tablixsSB.Append(BuildTablix(item));
            }

            return tablixsSB.ToString();
        }

        private string BuildTablix(TablixStyle style)
        {
            var coloums = new StringBuilder();
            var tablixHearders = new StringBuilder();
            var tablixCells = new StringBuilder();
            var tablixMembers = new StringBuilder();
            var availableWidth = style.Position.Width;
            var colCount = 0;

            foreach (var col in style.Columns)
            {
                if (col.Width > 0m)
                {
                    availableWidth -= col.Width;
                }
                else {
                    colCount += 1;
                }
            }

            foreach (var col in style.Columns)
            {
                var format = col.Format;
                var colName = col.Name;
                if (_dataSetFieldMap.ContainsKey(style.DataSetName))
                {
                    var fields = _dataSetFieldMap[style.DataSetName].Where(x => x.Key == col.Name).ToList();
                    if (fields.Count > 0)
                    {
                        colName = fields[0].Value;
                    }
                }

                if (col.Width == 0m)
                    col.Width = availableWidth / colCount;
                coloums.AppendFormat("<TablixColumn><Width>{0}cm</Width></TablixColumn>", col.Width);
                tablixHearders.AppendFormat("<TablixCell><CellContents>" +
                                            "<Textbox Name=\"Textbox{1}\"><CanGrow>true</CanGrow><KeepTogether>true</KeepTogether><Paragraphs><Paragraph>" +
                                            "<TextRuns><TextRun><Value>{0}</Value><Style><FontSize>10pt</FontSize><FontWeight>Bold</FontWeight><Color>#ffffff</Color></Style></TextRun></TextRuns><Style><TextAlign>Center</TextAlign></Style></Paragraph></Paragraphs>" +
                                            "<rd:DefaultName>Textbox{1}</rd:DefaultName><Style><Border><Color>LightGrey</Color><Style>Solid</Style></Border>" +
                                            "<BackgroundColor>#337ab7</BackgroundColor><VerticalAlign>Middle</VerticalAlign><PaddingLeft>2pt</PaddingLeft><PaddingRight>2pt</PaddingRight><PaddingTop>2pt</PaddingTop><PaddingBottom>2pt</PaddingBottom></Style></Textbox></CellContents></TablixCell>",
                                            col.Name, Guid.NewGuid().ToString().Replace("-", "_"));
                tablixCells.AppendFormat("<TablixCell><CellContents><Textbox Name=\"Textbox{2}1\"><CanGrow>true</CanGrow><KeepTogether>true</KeepTogether>" +
                                          "<Paragraphs><Paragraph><TextRuns><TextRun><Value>=Fields!{0}.Value</Value>{1}</TextRun></TextRuns><Style><TextAlign>Center</TextAlign></Style></Paragraph></Paragraphs>" +
                                          "<rd:DefaultName>Textbox{2}1</rd:DefaultName><Style><Border><Color>LightGrey</Color><Style>Solid</Style></Border>" +
                                          "<PaddingLeft>2pt</PaddingLeft><PaddingRight>2pt</PaddingRight><PaddingTop>2pt</PaddingTop><PaddingBottom>2pt</PaddingBottom></Style></Textbox></CellContents></TablixCell>",
                                          colName, string.IsNullOrWhiteSpace(format) ? "" : string.Format("<Style><Format>{0}</Format></Style>", format), Guid.NewGuid().ToString().Replace("-", "_"));

                tablixMembers.AppendFormat("<TablixMember />");
            }

            return _tablixPattern
                .Replace("@DataSetNameGuid", Guid.NewGuid().ToString().Replace("-", "_"))
                .Replace("@DataSetName",style.DataSetName)
                .Replace("@TablixColumns", coloums.ToString())
                .Replace("@TablixHeader", tablixHearders.ToString())
                .Replace("@TablixCells", tablixCells.ToString())
                .Replace("@TablixMember", tablixMembers.ToString())
                .Replace("@RowHeight", style.RowHight + this._unit)
                .Replace("@Top", style.Position.Top.ToString() + this._unit)
                .Replace("@Left", style.Position.Left.ToString() + this._unit)
                .Replace("@Height", style.Position.Height.ToString() + this._unit)
                .Replace("@Width", style.Position.Width.ToString() + this._unit);
        }

        #endregion


        private Stream GetRdlcStream(XmlDocument xmlDoc)
        {
            Stream ms = new MemoryStream();
            XmlSerializer serializer = new XmlSerializer(typeof(XmlDocument));
            serializer.Serialize(ms, xmlDoc);

            ms.Position = 0;
            return ms;
        }

        #endregion
    }
}