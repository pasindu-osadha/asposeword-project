using asposeword_project.Dtos.DocumentDtos;

namespace asposeword_project.Data.Interfaces
{
    public interface IWordEditorRepo
    {
        public bool saveWordtoPdf();
        public bool createDocument(DocumentCreateRequestDto documentCreateRequestDto);
        public bool createForms();
        public void UpdateFormData(FormDataDtos formData);
        public void createDocUsingJson();
        public void createDocUsingJsonFile(dynamic data);
    }
}
