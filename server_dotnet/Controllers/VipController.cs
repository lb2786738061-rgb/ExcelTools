using Microsoft.AspNetCore.Mvc;
using ExcelTools.Api.Models;
using ExcelTools.Api.Services;

namespace ExcelTools.Api.Controllers
{
    [ApiController]
    [Route("api/vip")]
    public class VipController : ControllerBase
    {
        private readonly DatabaseService _db;
        private readonly TokenService _tokens;

        public VipController(DatabaseService db, TokenService tokens)
        {
            _db = db;
            _tokens = tokens;
        }

        [HttpPost("redeem")]
        public IActionResult Redeem([FromBody] RedeemDto dto)
        {
            var auth = _tokens.VerifyFromRequest(Request);
            if (auth == null)
            {
                return Unauthorized(new { error = "未登录或登录凭证无效", code = "NEED_LOGIN" });
            }

            var user = _db.FindUserById(auth.Id);
            if (user == null)
            {
                return Unauthorized(new { error = "用户不存在或凭证已失效", code = "NEED_LOGIN" });
            }

            if (string.IsNullOrWhiteSpace(dto.Code))
            {
                return BadRequest(new { error = "请输入卡密兑换码！" });
            }

            var result = _db.RedeemCard(user.Id, dto.Code.Trim());
            if (!result.Success || result.User == null)
            {
                return BadRequest(new { error = result.Message });
            }

            return Ok(new
            {
                message = result.Message,
                user = _db.PublicUser(result.User)
            });
        }
    }
}
