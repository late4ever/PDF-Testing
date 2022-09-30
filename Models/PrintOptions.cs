using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using Telerik.Windows.Documents.Model;
using Telerik.Windows.Documents.Spreadsheet.FormatProviders;
using Telerik.Windows.Documents.Spreadsheet.Model.Printing;

namespace Kendo.Models
{
    /// <summary>
    /// Model for Spreadsheet print options
    /// </summary>
    public class PrintOptions
    {
        public ExportWhat Source { get; set; }
        public PaperTypes PaperSize { get; set; }
        public PageOrientation Orientation { get; set; }
        public PrintMargins Margins { get; set; }
        public bool CenterHorizontally { get; set; }
        public bool CenterVertically { get; set; }
        public bool PrintGridlines { get; set; }
    }
}