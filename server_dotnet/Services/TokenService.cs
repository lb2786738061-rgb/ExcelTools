using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using ExcelTools.Api.Models;

namespace ExcelTools.Api.Services
{
    public class TokenPayload
    {
        public string Id { get; set; } = "";
        public string Username { get; set; } = "";
        public string Role { get; set; } = "user";
    }

    /// <summary>
    /// HS256 JWT，与 Node server/token.js 使用同一算法。
    /// </summary>
    public class TokenService
    {
        private const string DefaultSecret = "exceltools-dev-hmac-secret-change-in-production";
        private const int DefaultExpiresSeconds = 7 * 24 * 60 * 60;

        private readonly byte[] _secretBytes;
        private readonly int _expiresSeconds;

        public TokenService(IConfiguration config)
        {
            var secret = FirstNonEmpty(
                config["JWT_SECRET"],
                Environment.GetEnvironmentVariable("JWT_SECRET"),
                DefaultSecret)!;
            _secretBytes = Encoding.UTF8.GetBytes(secret);

            var expRaw = FirstNonEmpty(
                config["JWT_EXPIRES_SECONDS"],
                Environment.GetEnvironmentVariable("JWT_EXPIRES_SECONDS"));
            _expiresSeconds = int.TryParse(expRaw, out var seconds) && seconds > 0
                ? seconds
                : DefaultExpiresSeconds;
        }

        public string CreateToken(User user)
        {
            var header = Base64UrlEncode(Encoding.UTF8.GetBytes("{\"alg\":\"HS256\",\"typ\":\"JWT\"}"));
            var now = DateTimeOffset.UtcNow.ToUnixTimeSeconds();
            var payloadJson = JsonSerializer.Serialize(new Dictionary<string, object>
            {
                ["id"] = user.Id,
                ["username"] = user.Username,
                ["role"] = user.Role,
                ["iat"] = now,
                ["exp"] = now + _expiresSeconds
            });
            var payload = Base64UrlEncode(Encoding.UTF8.GetBytes(payloadJson));
            var data = $"{header}.{payload}";
            var sig = Base64UrlEncode(HmacSha256(data));
            return $"{data}.{sig}";
        }

        public TokenPayload? VerifyFromRequest(HttpRequest request)
        {
            if (!request.Headers.TryGetValue("Authorization", out var authHeader) || string.IsNullOrEmpty(authHeader))
                return null;

            var parts = authHeader.ToString().Split(' ');
            if (parts.Length < 2 || !string.Equals(parts[0], "Bearer", StringComparison.OrdinalIgnoreCase))
                return null;

            return Verify(parts[1]);
        }

        public TokenPayload? Verify(string token)
        {
            if (string.IsNullOrWhiteSpace(token)) return null;
            var parts = token.Split('.');
            if (parts.Length != 3) return null;

            var data = $"{parts[0]}.{parts[1]}";
            var expected = Base64UrlEncode(HmacSha256(data));
            if (!FixedTimeEquals(parts[2], expected)) return null;

            try
            {
                var json = Encoding.UTF8.GetString(Base64UrlDecode(parts[1]));
                using var doc = JsonDocument.Parse(json);
                var root = doc.RootElement;
                if (!root.TryGetProperty("exp", out var expEl) || expEl.GetInt64() < DateTimeOffset.UtcNow.ToUnixTimeSeconds())
                    return null;
                if (!root.TryGetProperty("id", out var idEl)) return null;
                var id = idEl.GetString();
                if (string.IsNullOrEmpty(id)) return null;

                return new TokenPayload
                {
                    Id = id,
                    Username = root.TryGetProperty("username", out var u) ? u.GetString() ?? "" : "",
                    Role = root.TryGetProperty("role", out var r) ? r.GetString() ?? "user" : "user"
                };
            }
            catch
            {
                return null;
            }
        }

        private byte[] HmacSha256(string data)
        {
            using var hmac = new HMACSHA256(_secretBytes);
            return hmac.ComputeHash(Encoding.UTF8.GetBytes(data));
        }

        private static bool FixedTimeEquals(string a, string b)
        {
            var ba = Encoding.UTF8.GetBytes(a);
            var bb = Encoding.UTF8.GetBytes(b);
            if (ba.Length != bb.Length) return false;
            return CryptographicOperations.FixedTimeEquals(ba, bb);
        }

        private static string Base64UrlEncode(byte[] bytes)
        {
            return Convert.ToBase64String(bytes).TrimEnd('=').Replace('+', '-').Replace('/', '_');
        }

        private static byte[] Base64UrlDecode(string input)
        {
            var s = input.Replace('-', '+').Replace('_', '/');
            switch (s.Length % 4)
            {
                case 2: s += "=="; break;
                case 3: s += "="; break;
            }
            return Convert.FromBase64String(s);
        }

        private static string? FirstNonEmpty(params string?[] values)
        {
            foreach (var v in values)
            {
                if (!string.IsNullOrWhiteSpace(v)) return v;
            }
            return null;
        }
    }
}
