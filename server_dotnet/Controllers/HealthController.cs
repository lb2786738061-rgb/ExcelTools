using Microsoft.AspNetCore.Mvc;

namespace ExcelTools.Api.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class HealthController : ControllerBase
    {
        [HttpGet]
        public IActionResult GetHealth()
        {
            return Ok(new
            {
                status = "ok",
                service = "Office 智能美化工具箱 C# (.NET Core / Web API) 后端服务",
                timestamp = DateTime.UtcNow.ToString("o"),
                platform = ".NET 8.0 / C#"
            });
        }
    }
}

