using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Telerik.Windows.Documents.Spreadsheet.Model.Printing;
using Telerik.Windows.Documents.Spreadsheet.Utilities;

namespace Kendo.Models
{
    /// <summary>
    /// Model for print margins (cm)
    /// </summary>
    public class PrintMargins
    {
        public double Top { get; set; }
        public double Right { get; set; }
        public double Bottom { get; set; }
        public double Left { get; set; }

        /// <summary>
        /// Converts the margins to PageMargins object
        /// </summary>
        /// <param name="margins">Print margins set in cm</param>
        /// <returns>Page margins in device independent pixels</returns>
        public static implicit operator PageMargins(PrintMargins margins)
        {
            return new PageMargins(
                UnitHelper.CmToDip(margins.Left),
                UnitHelper.CmToDip(margins.Top),
                UnitHelper.CmToDip(margins.Right),
                UnitHelper.CmToDip(margins.Bottom)
            );
        }
    }
}
