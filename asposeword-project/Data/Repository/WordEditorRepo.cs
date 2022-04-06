using asposeword_project.Data.Interfaces;
using Aspose.Words;
using asposeword_project.Dtos.DocumentDtos;

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

    }
}
