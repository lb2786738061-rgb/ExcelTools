using System.Text.Json.Serialization;

namespace ExcelTools.Api.Models
{
    public class User
    {
        [JsonPropertyName("id")]
        public string Id { get; set; } = string.Empty;

        [JsonPropertyName("username")]
        public string Username { get; set; } = string.Empty;

        [JsonPropertyName("passwordHash")]
        public string PasswordHash { get; set; } = string.Empty;

        [JsonPropertyName("role")]
        public string Role { get; set; } = "user";

        [JsonPropertyName("isVip")]
        public bool IsVip { get; set; } = false;

        [JsonPropertyName("vipExpireAt")]
        public long VipExpireAt { get; set; } = 0;

        [JsonPropertyName("balance")]
        public int Balance { get; set; } = 3;

        [JsonPropertyName("createdAt")]
        public long CreatedAt { get; set; } = DateTimeOffset.UtcNow.ToUnixTimeMilliseconds();
    }

    public class VipCard
    {
        [JsonPropertyName("code")]
        public string Code { get; set; } = string.Empty;

        [JsonPropertyName("type")]
        public string Type { get; set; } = "month";

        [JsonPropertyName("value")]
        public int Value { get; set; } = 30;

        [JsonPropertyName("status")]
        public string Status { get; set; } = "unused";

        [JsonPropertyName("usedBy")]
        public string? UsedBy { get; set; }

        [JsonPropertyName("usedAt")]
        public long? UsedAt { get; set; }

        [JsonPropertyName("createdAt")]
        public long CreatedAt { get; set; } = DateTimeOffset.UtcNow.ToUnixTimeMilliseconds();
    }

    public class UsageLog
    {
        [JsonPropertyName("id")]
        public string Id { get; set; } = string.Empty;

        [JsonPropertyName("userId")]
        public string UserId { get; set; } = string.Empty;

        [JsonPropertyName("username")]
        public string Username { get; set; } = string.Empty;

        [JsonPropertyName("fileName")]
        public string FileName { get; set; } = string.Empty;

        [JsonPropertyName("type")]
        public string Type { get; set; } = string.Empty;

        [JsonPropertyName("cost")]
        public int Cost { get; set; } = 0;

        [JsonPropertyName("createdAt")]
        public long CreatedAt { get; set; } = DateTimeOffset.UtcNow.ToUnixTimeMilliseconds();
    }

    public class DbStore
    {
        [JsonPropertyName("users")]
        public List<User> Users { get; set; } = new();

        [JsonPropertyName("vipCards")]
        public List<VipCard> VipCards { get; set; } = new();

        [JsonPropertyName("usageLogs")]
        public List<UsageLog> UsageLogs { get; set; } = new();
    }

    // DTO 模型
    public class RegisterDto
    {
        public string? Username { get; set; }
        public string? Password { get; set; }
    }

    public class LoginDto
    {
        public string? Username { get; set; }
        public string? Password { get; set; }
    }

    public class RedeemDto
    {
        public string? Code { get; set; }
    }

    public class ProcessDocDto
    {
        public string? FileName { get; set; }
    }

    public class ConsumeUsageDto
    {
        public string? Feature { get; set; }
        public string? FileName { get; set; }
    }

    public class GenerateCardsDto
    {
        public int Count { get; set; } = 5;
        public string Type { get; set; } = "month";
        public int Value { get; set; } = 30;
    }
}

