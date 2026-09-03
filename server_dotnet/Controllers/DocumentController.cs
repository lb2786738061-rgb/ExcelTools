using Microsoft.AspNetCore.Mvc;
using ExcelTools.Api.Models;
using ExcelTools.Api.Services;

namespace ExcelTools.Api.Controllers
{
    [ApiController]
    [Route("api/document")]
    public class DocumentController : ControllerBase
    {
        private readonly DatabaseService _db;
        private readonly TokenService _tokens;

        public DocumentController(DatabaseService db, TokenService tokens)
        {
            _db = db;
            _tokens = tokens;
        }

        [HttpPost("process")]
        public IActionResult ProcessDocument([FromBody] ProcessDocDto dto)
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

            return Ok(new
            {
                message = "文件处理不扣次数。次数和包月只用于一键美化、智能助手。",
                quota = new
                {
                    success = true,
                    isVip = _db.IsMonthlyActive(user),
                    isMonthly = _db.IsMonthlyActive(user),
                    balance = user.Balance,
                    expireAt = user.VipExpireAt
                }
            });
        }
    }
}
