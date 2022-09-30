using Kendo.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web.Mvc;
using System.Windows;
using System.Windows.Media;
using Telerik.Web.Spreadsheet;
using Telerik.Windows.Documents.Fixed.Model;
using Telerik.Windows.Documents.Fixed.Model.ColorSpaces;
using Telerik.Windows.Documents.Fixed.Model.Editing;
using Telerik.Windows.Documents.Model;
using Telerik.Windows.Documents.Spreadsheet.FormatProviders.Pdf.Export;
using Telerik.Windows.Documents.Spreadsheet.Model;

namespace PDF_Testing.Controllers
{
    public partial class SpreadsheetController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult Server_Side_Pdf_Export()
        {
            return View();
        }

        [HttpPost]
        public ActionResult SaveAsPDF(string data, string activeSheet, PrintOptions options, string selection)
        {
            var workbook = Telerik.Web.Spreadsheet.Workbook.FromJson(data);
            var document = workbook.ToDocument();

            var sheetProvider = new Telerik.Windows.Documents.Spreadsheet.FormatProviders.Pdf.PdfFormatProvider
            {
                ExportSettings = new PdfExportSettings(options.Source, false)
            };

            document.ActiveSheet = document.Worksheets.First(sheet => sheet.Name == activeSheet);

            foreach (var sheet in document.Worksheets)
            {
                var pageSetup = sheet.WorksheetPageSetup;

                pageSetup.PaperType = options.PaperSize;

                if (sheet.Name.Equals("Food Order #1"))
                {
                    pageSetup.PageOrientation = PageOrientation.Portrait;
                }
                else
                {
                    pageSetup.PageOrientation = PageOrientation.Landscape;
                }

                pageSetup.FitToPages = true;
                pageSetup.Margins = options.Margins;
                pageSetup.CenterHorizontally = true;
                pageSetup.CenterVertically = options.CenterVertically;

                pageSetup.PrintOptions.PrintGridlines = options.PrintGridlines;

            }

            using (var stream = new MemoryStream())
            {
                if (selection != "" && options.Source.ToString() == "Selection")
                {
                    int[] rangeParams = selection.Split(',').Select(int.Parse).ToArray();
                    var topLeftRow = rangeParams[0];
                    var topLeftCow = rangeParams[1];
                    var bottomRightRow = rangeParams[2];
                    var bottomRightCol = rangeParams[3];
                    var range = new CellRange(topLeftRow, topLeftCow, bottomRightRow, bottomRightCol);
                    sheetProvider.ExportSettings = new PdfExportSettings(new List<CellRange>() { range });
                }
                sheetProvider.Export(document, stream);

                var pdfProvider = new Telerik.Windows.Documents.Fixed.FormatProviders.Pdf.PdfFormatProvider();
                RadFixedDocument doc = pdfProvider.Import(stream);

                
                GenerateCoverPage(doc);
                DrawHeaderFooterWatermark(doc);

                pdfProvider.Export(doc, stream);
                var mimeType = MimeTypes.PDF;
                return File(stream.ToArray(), mimeType, "Print.pdf");
            }
        }

        private void GenerateCoverPage(RadFixedDocument document)
        {
            RadFixedPage coverPage = new RadFixedPage();

            using (Stream imageStream = System.IO.File.OpenRead(AppDomain.CurrentDomain.BaseDirectory + "Cover.jpg"))
            {
                Telerik.Windows.Documents.Fixed.Model.Resources.ImageSource image = new Telerik.Windows.Documents.Fixed.Model.Resources.ImageSource(imageStream);
                coverPage.Size = new Size(image.Width, image.Height);
                FixedContentEditor imagePageEditor = new FixedContentEditor(coverPage);
                imagePageEditor.DrawImage(image);
            }

            document.Pages.AddPage();
            for (var i = document.Pages.Count - 1; i > 0; i--)
            {
                document.Pages[i] = document.Pages[i - 1];
            }
            document.Pages[0] = coverPage;
        }

        private void DrawHeaderFooterWatermark(RadFixedDocument document)
        {
            int numberOfPages = document.Pages.Count;
            for (int pageIndex = 0; pageIndex < numberOfPages; pageIndex++)
            {
                int pageNumber = pageIndex + 1;
                RadFixedPage currentPage = document.Pages[pageIndex];
                DrawHeaderAndFooterToPage(currentPage, pageNumber, numberOfPages);
                AddWatermarkText(currentPage, "Watermark text!", 100);
            }
        }

        private void DrawHeaderAndFooterToPage(RadFixedPage page, int pageNumber, int numberOfPages)
        {
            FixedContentEditor pageEditor = new FixedContentEditor(page);

            Block header = new Block();
            Telerik.Windows.Documents.Fixed.Model.Resources.ImageSource imageSource = ImportImage(AppDomain.CurrentDomain.BaseDirectory + "logo-light.png");
            Size imageSize = new Size(168, 50);
            header.InsertImage(imageSource, imageSize);
            header.Measure();

            double headerOffsetX = (page.Size.Width / 2) - (header.DesiredSize.Width / 2);
            double headerOffsetY = 50;
            pageEditor.Position.Translate(headerOffsetX, headerOffsetY);
            pageEditor.DrawBlock(header);

            Block footer = new Block();
            footer.InsertText(String.Format("Page {0} of {1}", pageNumber, numberOfPages));
            footer.Measure();

            double footerOffsetX = (page.Size.Width / 2) - (footer.DesiredSize.Width / 2);
            double footerOffsetY = page.Size.Height - 50 - footer.DesiredSize.Height;
            pageEditor.Position.Translate(footerOffsetX, footerOffsetY);
            pageEditor.DrawBlock(footer);
        }

        private Telerik.Windows.Documents.Fixed.Model.Resources.ImageSource ImportImage(string inputLogoFile)
        {
            Telerik.Windows.Documents.Fixed.Model.Resources.ImageSource imageSource;
            using (FileStream source = System.IO.File.Open(inputLogoFile, FileMode.Open))
            {
                imageSource = new Telerik.Windows.Documents.Fixed.Model.Resources.ImageSource(source);
            }

            return imageSource;
        }

        private void AddWatermarkText(RadFixedPage page, string text, byte transparency)
        {
            FixedContentEditor editor = new FixedContentEditor(page);

            Block block = new Block();
            block.TextProperties.FontSize = 80;
            block.TextProperties.TrySetFont(new FontFamily("Arial"), FontStyles.Normal, FontWeights.Bold);
            block.HorizontalAlignment = Telerik.Windows.Documents.Fixed.Model.Editing.Flow.HorizontalAlignment.Center;
            block.GraphicProperties.FillColor = new RgbColor(transparency, 255, 0, 0);
            block.InsertText(text);

            double angle = -45;
            editor.Position.Rotate(angle);
            editor.Position.Translate(0, page.Size.Width/1.5);
            editor.DrawBlock(block, new Size(page.Size.Width / Math.Abs(Math.Sin(angle)), page.Size.Height/2));
        }
    }
}