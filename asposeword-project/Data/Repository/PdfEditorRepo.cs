using Aspose.Pdf;
using System;
using System.IO;
using asposeword_project.Data.Interfaces;
using Aspose.Pdf.Text;

namespace asposeword_project.Data.Repository
{
    public class PdfEditorRepo : IPdfEditorRepo
    {

        public void MakeComplexDocument()
        {
            string _dataDir = "Files\\pdfEditor";
            // Initialize document object
            Document document = new Document();

            // Add page
            Page page = document.Pages.Add();


            // Add Header
            var header = new TextFragment("Allion Technologies - Audit result");
            header.TextState.Font = FontRepository.FindFont("Arial");
            header.TextState.FontSize = 20;
            header.HorizontalAlignment = HorizontalAlignment.Center;
            // header.Position = new Position(0, 820 );

            HeaderFooter h = new HeaderFooter();
            h.Paragraphs.Add(header);
            page.Header = h;
            page.Footer = h;


            // Add description
            var descriptionText = "Visitors must buy tickets online and tickets are limited to 5,000 per day. Ferry service is operating at half capacity and on a reduced schedule. Expect lineups.";
            var description = new TextFragment(descriptionText);
            description.TextState.Font = FontRepository.FindFont("Times New Roman");
            description.TextState.FontSize = 14;
            description.HorizontalAlignment = HorizontalAlignment.Left;
            page.Paragraphs.Add(description);


            // Add table
            var table = new Table
            {
                ColumnWidths = "200",
                Border = new BorderInfo(BorderSide.Box, 1f, Color.DarkSlateGray),
                DefaultCellBorder = new BorderInfo(BorderSide.Box, 0.5f, Color.Black),
                DefaultCellPadding = new MarginInfo(4.5, 4.5, 4.5, 4.5),
                Margin =
                {
                    Bottom = 10
                },
                DefaultCellTextState =
                {
                    Font =  FontRepository.FindFont("Helvetica")
                }
            };

            var headerRow = table.Rows.Add();
            headerRow.Cells.Add("Departs City");
            headerRow.Cells.Add("Departs Island");
            foreach (Cell headerRowCell in headerRow.Cells)
            {
                headerRowCell.BackgroundColor = Color.Gray;
                headerRowCell.DefaultCellTextState.ForegroundColor = Color.WhiteSmoke;
            }

            var time = new TimeSpan(6, 0, 0);
            var incTime = new TimeSpan(0, 30, 0);
            for (int i = 0; i < 10; i++)
            {
                var dataRow = table.Rows.Add();
                dataRow.Cells.Add(time.ToString(@"hh\:mm"));
                time = time.Add(incTime);
                dataRow.Cells.Add(time.ToString(@"hh\:mm"));
            }

            page.Paragraphs.Add(table);

            document.Save(System.IO.Path.Combine(_dataDir, "Complex.pdf"));

        }
    }
}
