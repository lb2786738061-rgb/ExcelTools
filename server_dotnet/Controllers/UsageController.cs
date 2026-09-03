using Microsoft.AspNetCore.Mvc;
using ExcelTools.Api.Models;
using ExcelTools.Api.Services;

namespace ExcelTools.Api.Controllers
{
    [ApiController]
    [Route("api/usage")]
    public class UsageController : ControllerBase
    {
        private readonly DatabaseService _db;
        private readonly TokenService _tokens;

        public UsageController(DatabaseService db, TokenService tokens)
        {
            _db = db;
            _tokens = tokens;
        }

        [HttpPost("consume")]
        public IActionResult Consume([FromBody] ConsumeUsageDto dto)
        {
            var auth = _tokens.VerifyFromRequest(Request);
            if (auth == null)
                return Unauthorized(new { error = "未登录或登录凭证无效", code = "NEED_LOGIN" });

            var user = _db.FindUserById(auth.Id);
            if (user == null)
                return Unauthorized(new { error = "用户不存在或凭证已失效", code = "NEED_LOGIN" });

            var feature = (dto.Feature ?? "beautify").Trim().ToLowerInvariant();
            if (feature is not ("beautify" or "agent"))
                feature = "beautify";

            var result = _db.ConsumeAiUsage(user.Id, feature, dto.FileName);
            if (!result.Success)
                return StatusCode(403, new { error = result.Message, code = result.Code ?? "NEED_RECHARGE" });

            return Ok(new
            {
                message = result.Message,
                quota = new
                {
                    success = true,
                    isVip = result.IsMonthly,
                    isMonthly = result.IsMonthly,
                    balance = result.Balance,
                    expireAt = result.ExpireAt
                }
            });
        }
    }
}
