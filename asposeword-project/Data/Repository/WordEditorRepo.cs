using asposeword_project.Data.Interfaces;
using Aspose.Words;

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

                throw;
                
            }
            return false;

        }
    }
}
