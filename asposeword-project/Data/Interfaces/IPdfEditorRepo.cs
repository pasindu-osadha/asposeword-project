using asposeword_project.Dtos.DocumentDtos;

namespace asposeword_project.Data.Interfaces
{
    public interface IPdfEditorRepo
    {
       void MakeComplexDocument(RequestPDFDto Dto);
    }
}
