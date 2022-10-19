using Microsoft.AspNetCore.Mvc;

namespace Pricer.API.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class TestController : ControllerBase
    {
        private readonly ILogger<TestController> _logger;

        public TestController(ILogger<TestController> logger)
        {
            _logger = logger;
        }

        [HttpGet]
        public IActionResult Test()
        {
            return Ok(new { test = true });
        }
    }
}