using asposeword_project.Data.Interfaces;
using asposeword_project.Dtos.DocumentDtos;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Text.Json;

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
            if (_repo.createForms())
                return Ok();
            else return BadRequest();
        }

        [HttpPost]
        [Route("editformdata")]
        public ActionResult EditFormData( FormDataDtos formData)
        {
            _repo.UpdateFormData(formData);
            return Ok();
        }

        [HttpPost]
        [Route("json")]
        public ActionResult jsonmethod([FromBody] System.Text.Json.JsonElement entity)
        {
            _repo.createDocUsingJson();
            return Ok();
        }



        [HttpPost]
        [Route ("testapi")]
        public ActionResult testapi()
        {
            //dynamic stuff = JsonConvert.DeserializeObject("{ 'Name': 'Jon Smith', 'Address': { 'City': 'New York', 'State': 'NY' }, 'Age': 42 }");

            dynamic dstuff = JsonConvert.DeserializeObject("[{'Name':'John Smith','Contract':[{'Client':{'Name':'A Company'},'Price':1200000},{'Client':{'Name':'B Ltd.'},'Price':750000},{'Client':{'Name':'C & D'},'Price':350000}]}]");

           // dynamic stuff = JObject.Parse("[{'Name':'John Smith','Contract':[{'Client':{'Name':'A Company'},'Price':1200000},{'Client':{'Name':'B Ltd.'},'Price':750000},{'Client':{'Name':'C & D'},'Price':350000}]}]");

            string name = dstuff[0].Name;
          

          //  string jsonString =JsonConvert.SerializeObject(dstuff);
            _repo.createDocUsingJsonFile(dstuff);
            return Ok();
        }
    }
}
