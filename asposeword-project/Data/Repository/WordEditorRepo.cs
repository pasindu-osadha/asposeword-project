using asposeword_project.Data.Interfaces;
using Aspose.Words;
using asposeword_project.Dtos.DocumentDtos;
using Aspose.Words.Fields;
using NUnit.Framework;

namespace asposeword_project.Data.Repository
{
    public class WordEditorRepo : IWordEditorRepo
    {
        public bool saveWordtoPdf()
        {
            try
            {
                Document document = new Document("Files\\sample.docx");
                document.Save("Files\\sample.pdf");
                return true;
            }
            catch (Exception)
            {
                return false;
                throw;  
            }
           

        }

        public bool createDocument(DocumentCreateRequestDto documentCreateRequestDto)
        {
            try
            {
                Document document = new Document();
                DocumentBuilder documentBuilder = new DocumentBuilder(document);

                documentBuilder.Writeln(documentCreateRequestDto.content);

                document.Save("Files\\" + documentCreateRequestDto.filetype);


                return true;
            }
            catch (Exception)
            {
                return false;
                throw;
            }

        
        }

        public bool createForms()
        {   
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Write("First name : ");
            builder.InsertTextInput("firstname", TextFormFieldType.Regular, "", "Enter your first name ", 0);
            builder.InsertBreak(BreakType.LineBreak);

            builder.Write("Last name : ");
            builder.InsertTextInput("lastname", TextFormFieldType.Regular, "", "Enter your last name ", 0);
            builder.InsertBreak(BreakType.LineBreak);

            builder.Write("Address : ");
            builder.InsertTextInput("address", TextFormFieldType.Regular, "", "Enter your address ", 0);
            builder.InsertBreak(BreakType.LineBreak);


            doc.Save("Files\\FormTemplete.docx");


            return true;
        } 

    }
}
