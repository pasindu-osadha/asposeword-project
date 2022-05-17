using asposeword_project.Data.Interfaces;
using asposeword_project.Dtos.DocumentDtos;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;

namespace asposeword_project.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class PdfEditorController : ControllerBase
    {
        private readonly IPdfEditorRepo _repo;

        public PdfEditorController(IPdfEditorRepo repo )
        {
            _repo = repo;
        }
        [HttpPost]
        [Route("createPdf")]
        public ActionResult craetePdf(RequestPDFDto requestPDFDto)
        {
            _repo.MakeComplexDocument(requestPDFDto);
            return Ok(); 
        }
    }
}
