using asposeword_project.Data.Interfaces;
using Aspose.Words;
using asposeword_project.Dtos.DocumentDtos;
using Aspose.Words.Fields;
using NUnit.Framework;
using Aspose.Words.Reporting;
using Newtonsoft.Json;
using System.Data;

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
            try
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
            catch (Exception)
            {
                return false;
                throw;
            }
           
        } 


        public void UpdateFormData(FormDataDtos formData)
        {
            try
            {
                Document doc = new Document("Files\\FormTemplete.docx");
                FormFieldCollection documentFormFields = doc.Range.FormFields;

                FormField firstnamefield = documentFormFields["firstname"];
                FormField lastnamefield = documentFormFields["lastname"];
                FormField address = documentFormFields["address"];


                if (firstnamefield.Type.Equals(FieldType.FieldFormTextInput))
                    firstnamefield.Result = formData.firstName;

                if (lastnamefield.Type.Equals(FieldType.FieldFormTextInput))
                    lastnamefield.Result = formData.lastName;

                if (address.Type.Equals(FieldType.FieldFormTextInput))
                    address.Result = formData.address;

                doc.Save("Files\\FormTempleteEdited.docx");
            }
            catch (Exception)
            {

                throw;
            }
           
        }


        public void createDocUsingJson()
        {
            
            Document doc = new Document("Files\\json\\template.docx");
            JsonDataSource dataSource = new JsonDataSource("Files\\json\\datasource.json");
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, dataSource, "managers");
            doc.Save("Files\\json\\jsonresult.docx");
        }

        public void createDocUsingJsonFile(dynamic data)
        {
            string jsonString = JsonConvert.SerializeObject(data);
            File.WriteAllText(@"Files\\jsoncustom\\customdatasource.json", jsonString);

            Document doc = new Document("Files\\jsoncustom\\template.docx");
            JsonDataSource dataSource = new JsonDataSource("Files\\jsoncustom\\customdatasource.json");
           
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, dataSource, "managers");
            doc.Save("Files\\jsoncustom\\customjsonresult.docx");

        }

    }
}
