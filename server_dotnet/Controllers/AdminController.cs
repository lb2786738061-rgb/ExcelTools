using Microsoft.AspNetCore.Mvc;
using ExcelTools.Api.Models;
using ExcelTools.Api.Services;

namespace ExcelTools.Api.Controllers
{
    [ApiController]
    [Route("api/admin")]
    public class AdminController : ControllerBase
    {
        private readonly DatabaseService _db;
        private readonly TokenService _tokens;

        public AdminController(DatabaseService db, TokenService tokens)
        {
            _db = db;
            _tokens = tokens;
        }

        private IActionResult? RequireAdmin(out User? admin)
        {
            admin = null;
            var auth = _tokens.VerifyFromRequest(Request);
            if (auth == null)
            {
                return Unauthorized(new { error = "未登录或登录凭证无效", code = "NEED_LOGIN" });
            }

            var user = _db.FindUserById(auth.Id);
            if (user == null || user.Role != "admin")
            {
                return StatusCode(403, new { error = "越权操作，需要管理员权限！" });
            }

            admin = user;
            return null;
        }

        [HttpGet("dashboard")]
        public IActionResult GetDashboard()
        {
            var denied = RequireAdmin(out _);
            if (denied != null) return denied;
            return Ok(_db.GetStats());
        }

        [HttpPost("cards/generate")]
        public IActionResult GenerateCards([FromBody] GenerateCardsDto dto)
        {
            var denied = RequireAdmin(out _);
            if (denied != null) return denied;

            var cards = _db.GenerateVipCards(dto.Count, dto.Type, dto.Value);
            return Ok(new
            {
                message = $"成功生成 {cards.Count} 张卡密（次数包或包月）",
                cards
            });
        }
    }
}
