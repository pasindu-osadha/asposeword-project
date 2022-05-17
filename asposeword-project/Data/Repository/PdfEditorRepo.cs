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
            header.Position = new Position(100, 800);
            HeaderFooter h = new HeaderFooter();
            h.Paragraphs.Add(header);
            page.Header = h;

            //Header table 
            var headerTable = new Table
            {
                DefaultCellPadding = new MarginInfo(4.5, 4.5, 4.5, 4.5),
                Margin =
                {
                    Bottom = 20
                }
            };
            headerTable.ColumnWidths = "150 150 150";

            var codeName_Row = headerTable.Rows.Add();
            codeName_Row.Cells.Add("Code");
            var code_Name_Row_cell = codeName_Row.Cells.Add("Name");
            code_Name_Row_cell.ColSpan = 2;

            var codeName_Data_Row = headerTable.Rows.Add();
            codeName_Data_Row.Cells.Add("IPAS 25A"); // API 
            var codeName_Data_Row_cell = codeName_Data_Row.Cells.Add("Room Cleaning on Discharge"); // API
            codeName_Data_Row_cell.ColSpan = 2;


            var description_Row = headerTable.Rows.Add();
            var description_Row_cell = description_Row.Cells.Add("Description");
            description_Row_cell.ColSpan = 3;

            var description_Row_Data = headerTable.Rows.Add();
            var description_Row_Data_cell = description_Row_Data.Cells.Add("This checklist is to be commenced on the day of discharge and completed prior to the room/ unit being occupied by another resident. "); // API
            description_Row_Data_cell.ColSpan = 3;


            var completedBY_Row = headerTable.Rows.Add();
            completedBY_Row.Cells.Add("Date Taken");
            completedBY_Row.Cells.Add("Completed by");
            completedBY_Row.Cells.Add("Designation");

            var completedBY_Data_Row = headerTable.Rows.Add();
            completedBY_Data_Row.Cells.Add("10/05/2022"); //API
            completedBY_Data_Row.Cells.Add("Salindra Gulawita"); //API
            completedBY_Data_Row.Cells.Add("T"); //API

            page.Paragraphs.Add(headerTable);


            //Data table 

            var dataTable = new Table
            {
                DefaultCellPadding = new MarginInfo(4.5, 4.5, 4.5, 4.5),
                Margin =
                {
                    Bottom = 10
                }
            };

            dataTable.ColumnWidths = "220 40 40 40 40 40";

            var dataTable_HeaderRow = dataTable.Rows.Add();
            dataTable_HeaderRow.Cells.Add("criteria");
            dataTable_HeaderRow.Cells.Add("complted");
            dataTable_HeaderRow.Cells.Add("findir");
            dataTable_HeaderRow.Cells.Add("response");
            dataTable_HeaderRow.Cells.Add("corrective");
            dataTable_HeaderRow.Cells.Add("actions");

            for (int i = 0; i < 125; i++)
            {
                var dataTable_DataRow = dataTable.Rows.Add();
                dataTable_DataRow.Cells.Add("Hand-basin/vanity top and sides wiped down with cleaning agent and mirror cleaned ");//API
                dataTable_DataRow.Cells.Add("-");
                dataTable_DataRow.Cells.Add("-");
                dataTable_DataRow.Cells.Add("-");
                dataTable_DataRow.Cells.Add("-");
                dataTable_DataRow.Cells.Add("-");

            }


            page.Paragraphs.Add(dataTable);



            document.Save(System.IO.Path.Combine(_dataDir, "Complex.pdf"));

        }
    }
}
