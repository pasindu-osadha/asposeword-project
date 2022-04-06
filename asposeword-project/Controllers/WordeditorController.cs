using asposeword_project.Data.Interfaces;
using asposeword_project.Dtos.DocumentDtos;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;

namespace asposeword_project.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class WordeditorController : ControllerBase
    {
        private readonly IWordEditorRepo _repo;

        public WordeditorController(IWordEditorRepo repo)
        {
            _repo = repo;
        }

        [HttpGet]
        [Route("createPdf")]
        public ActionResult createPdf()
        {
            if (_repo.saveWordtoPdf())
            {
                return Ok();
            }
            else
            {
                return BadRequest();
            }
        }

        [HttpPost]
        [Route("createDocument")]
        public ActionResult createDocument(DocumentCreateRequestDto documentCreateRequestDto)
        {
            if (_repo.createDocument(documentCreateRequestDto))
                return Ok();
            else
                return BadRequest();
        }

        [HttpPost]
        [Route("createform")]
        public ActionResult createform()
        {
            _repo.createForms();
            return Ok();
        }
    }
}
