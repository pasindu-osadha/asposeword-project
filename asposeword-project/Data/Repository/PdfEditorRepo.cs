using Aspose.Pdf;
using System;
using System.IO;
using asposeword_project.Data.Interfaces;
using Aspose.Pdf.Text;
using asposeword_project.Dtos.DocumentDtos;

namespace asposeword_project.Data.Repository
{
    public class PdfEditorRepo : IPdfEditorRepo
    {

        public void MakeComplexDocument(RequestPDFDto Dto)
        {
            string _dataDir = "Files\\pdfEditor";
            // Initialize document object
            Document document = new Document();

            // Add page
            Page page = document.Pages.Add();

            //Add header 
            var headerTable = new Table
            {
                DefaultCellPadding = new MarginInfo(4.5, 4.5, 4.5, 4.5),
                Margin =
                {
                    Top= 20
                }
            };
            headerTable.ColumnAdjustment = ColumnAdjustment.AutoFitToWindow;
            var headerTable_Row = headerTable.Rows.Add();
            headerTable_Row.Cells.Add(DateTime.Now.ToString());
            headerTable_Row.Cells.Add("HCSL Quality Assurance system");
            HeaderFooter headerFooter = new HeaderFooter();
            headerFooter.Paragraphs.Add(headerTable);
            page.Header = headerFooter;



            //Title 
            var title = new TextFragment("Allion Technologies - Audit result");
            title.TextState.FontSize = 20;
            title.HorizontalAlignment = HorizontalAlignment.Center;
            page.Paragraphs.Add(title);


            //Header table 
            var headerInformationTable = new Table
            {
                DefaultCellPadding = new MarginInfo(4.5, 4.5, 4.5, 4.5),
                Margin =
                {   
                    Bottom = 20
                }
            };
            headerInformationTable.ColumnWidths = "150 150 150";

            var codeName_Row = headerInformationTable.Rows.Add();
            codeName_Row.Cells.Add("Code");
            var code_Name_Row_cell = codeName_Row.Cells.Add("Name");
            code_Name_Row_cell.ColSpan = 2;

            var codeName_Data_Row = headerInformationTable.Rows.Add();
            codeName_Data_Row.Cells.Add(Dto.Code); // API 
            var codeName_Data_Row_cell = codeName_Data_Row.Cells.Add(Dto.Name); // API
            codeName_Data_Row_cell.ColSpan = 2;


            var description_Row = headerInformationTable.Rows.Add();
            var description_Row_cell = description_Row.Cells.Add("Description");
            description_Row_cell.ColSpan = 3;

            var description_Row_Data = headerInformationTable.Rows.Add();
            var description_Row_Data_cell = description_Row_Data.Cells.Add(Dto.Description); // API
            description_Row_Data_cell.ColSpan = 3;


            var completedBY_Row = headerInformationTable.Rows.Add();
            completedBY_Row.Cells.Add("Date Taken");
            completedBY_Row.Cells.Add("Completed by");
            completedBY_Row.Cells.Add("Designation");

            var completedBY_Data_Row = headerInformationTable.Rows.Add();
            completedBY_Data_Row.Cells.Add(Dto.DateTaken.ToString()); //API
            completedBY_Data_Row.Cells.Add(Dto.CompletedBy); //API
            completedBY_Data_Row.Cells.Add(Dto.Designation); //API

            page.Paragraphs.Add(headerInformationTable);


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

            //Image image = new Image();
            //image.File = "Files\\imageFile\\yes.png";


            for (int i = 0; i < Dto.tableData.Count; i++)
            {


                var dataTable_DataRow = dataTable.Rows.Add();
                dataTable_DataRow.Cells.Add(Dto.tableData[i].Criteria);//API

                Image img = new Image();

                switch (Dto.tableData[i].Completed)
                {
                    case 0:
                        img.File = "Files//imageFile//no.png";
                        break;
                    case 1:
                        img.File = "Files//imageFile//yes.png";
                        break;
                    case 2:
                        img.File = "Files//imageFile//undefine.png";
                        break;
                    default:
                        img.File = "Files//imageFile//undefine.png";
                        break;
                }

                img.FixHeight = 10;
                img.FixWidth = 10;
                var c = dataTable_DataRow.Cells.Add();
                c.Paragraphs.Add(img);

                dataTable_DataRow.Cells.Add(Dto.tableData[i].Findir);
                dataTable_DataRow.Cells.Add(Dto.tableData[i].Response);
                dataTable_DataRow.Cells.Add(Dto.tableData[i].Corrective);
                dataTable_DataRow.Cells.Add(Dto.tableData[i].Action);

            }


            page.Paragraphs.Add(dataTable);



            document.Save(System.IO.Path.Combine(_dataDir, "Complex.pdf"));

        }
    }
}
