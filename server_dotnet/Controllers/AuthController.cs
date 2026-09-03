using Microsoft.AspNetCore.Mvc;
using ExcelTools.Api.Models;
using ExcelTools.Api.Services;

namespace ExcelTools.Api.Controllers
{
    [ApiController]
    public class AuthController : ControllerBase
    {
        private readonly DatabaseService _db;
        private readonly TokenService _tokens;

        public AuthController(DatabaseService db, TokenService tokens)
        {
            _db = db;
            _tokens = tokens;
        }

        [HttpPost("api/auth/register")]
        public IActionResult Register([FromBody] RegisterDto dto)
        {
            if (string.IsNullOrWhiteSpace(dto.Username) || string.IsNullOrWhiteSpace(dto.Password))
            {
                return BadRequest(new { error = "用户名和密码不能为空" });
            }

            var existing = _db.FindUserByUsername(dto.Username);
            if (existing != null)
            {
                return BadRequest(new { error = "该用户名已被注册，请更换！" });
            }

            var newUser = _db.CreateUser(dto.Username, dto.Password);
            return Ok(new
            {
                message = "注册成功！已默认赠送 3 次免费试用体验",
                token = _tokens.CreateToken(newUser),
                user = _db.PublicUser(newUser)
            });
        }

        [HttpPost("api/auth/login")]
        public IActionResult Login([FromBody] LoginDto dto)
        {
            if (string.IsNullOrWhiteSpace(dto.Username) || string.IsNullOrWhiteSpace(dto.Password))
            {
                return BadRequest(new { error = "用户名和密码不能为空" });
            }

            var user = _db.FindUserByUsername(dto.Username);
            if (user == null || user.Role == "guest" || !_db.VerifyPassword(dto.Password, user.PasswordHash))
            {
                return Unauthorized(new { error = "用户名或密码不正确" });
            }

            if (!user.PasswordHash.StartsWith("pbkdf2$", StringComparison.Ordinal))
            {
                _db.UpgradePasswordHash(user.Id, dto.Password);
            }

            return Ok(new
            {
                message = "登录成功！",
                token = _tokens.CreateToken(user),
                user = _db.PublicUser(user)
            });
        }

        [HttpPost("api/auth/guest")]
        public IActionResult CreateGuest()
        {
            var guest = _db.CreateGuestUser();
            return Ok(new
            {
                message = "已创建访客试用会话",
                token = _tokens.CreateToken(guest),
                user = _db.PublicUser(guest)
            });
        }

        [HttpGet("api/user/profile")]
        public IActionResult GetProfile()
        {
            var auth = _tokens.VerifyFromRequest(Request);
            if (auth == null)
            {
                return Unauthorized(new { error = "未登录或登录凭证无效", code = "NEED_LOGIN" });
            }

            var user = _db.FindUserById(auth.Id);
            if (user == null) return Unauthorized(new { error = "用户不存在或凭证已失效", code = "NEED_LOGIN" });

            return Ok(new { user = _db.PublicUser(user) });
        }
    }
}
