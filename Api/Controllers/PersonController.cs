using DataAccessLibrary;
using Microsoft.AspNetCore.Mvc;
using System;
using System.Threading.Tasks;

namespace Api.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class PersonController : ControllerBase
    {
        private readonly IDataAccess _dataAccess;

        public PersonController(IDataAccess dataAccess)
        {
            _dataAccess = dataAccess;
        }

        [HttpGet]
        public async Task<ActionResult> GetPeople()
        {
            try
            {
                var result = await _dataAccess.LoadFile();

                return Ok(result);
            }
            catch(Exception ex)
            {
                return StatusCode(500, ex.Message);
            }
        }

        [HttpPost]
        public async Task<ActionResult> PostPeople()
        {
            try
            {
                await _dataAccess.CreateFile();
                return Ok("File created successfully");
            }
            catch(Exception ex)
            {
                return StatusCode(500, ex.Message);
            }
        }
    }
}
