using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using ExcelTools.Api.Models;

namespace ExcelTools.Api.Services
{
    public class DatabaseService
    {
        private const int Pbkdf2Iterations = 100000;
        private const int Pbkdf2KeyLen = 32;

        private readonly string _dataDir;
        private readonly string _storeFile;
        private readonly object _lockObj = new();
        private readonly IConfiguration _config;

        public DatabaseService(IWebHostEnvironment env, IConfiguration config)
        {
            _config = config;
            _dataDir = Path.Combine(env.ContentRootPath, "..", "server", "data");
            _storeFile = Path.Combine(_dataDir, "db.json");
            InitDb();
        }

        private string EnvOr(string key, string fallback)
        {
            var value = _config[key];
            if (!string.IsNullOrWhiteSpace(value)) return value.Trim();
            var env = Environment.GetEnvironmentVariable(key);
            return string.IsNullOrWhiteSpace(env) ? fallback : env.Trim();
        }

        public string HashPassword(string pwd)
        {
            var salt = RandomNumberGenerator.GetBytes(16);
            var hash = Rfc2898DeriveBytes.Pbkdf2(
                Encoding.UTF8.GetBytes(pwd),
                salt,
                Pbkdf2Iterations,
                HashAlgorithmName.SHA256,
                Pbkdf2KeyLen);
            return $"pbkdf2${Pbkdf2Iterations}${Convert.ToHexString(salt).ToLowerInvariant()}${Convert.ToHexString(hash).ToLowerInvariant()}";
        }

        public bool VerifyPassword(string pwd, string stored)
        {
            if (string.IsNullOrEmpty(pwd) || string.IsNullOrEmpty(stored)) return false;

            if (stored.StartsWith("pbkdf2$", StringComparison.Ordinal))
            {
                var parts = stored.Split('$');
                if (parts.Length != 4) return false;
                if (!int.TryParse(parts[1], out var iterations)) return false;
                var salt = Convert.FromHexString(parts[2]);
                var expected = Convert.FromHexString(parts[3]);
                var actual = Rfc2898DeriveBytes.Pbkdf2(
                    Encoding.UTF8.GetBytes(pwd),
                    salt,
                    iterations,
                    HashAlgorithmName.SHA256,
                    expected.Length);
                return CryptographicOperations.FixedTimeEquals(actual, expected);
            }

            using var sha256 = SHA256.Create();
            var legacy = sha256.ComputeHash(Encoding.UTF8.GetBytes(pwd));
            var legacyHex = Convert.ToHexString(legacy).ToLowerInvariant();
            var a = Encoding.UTF8.GetBytes(legacyHex);
            var b = Encoding.UTF8.GetBytes(stored.ToLowerInvariant());
            if (a.Length != b.Length) return false;
            return CryptographicOperations.FixedTimeEquals(a, b);
        }

        public bool IsMonthlyActive(User user)
        {
            var now = DateTimeOffset.UtcNow.ToUnixTimeMilliseconds();
            return user.VipExpireAt > now;
        }

        public object PublicUser(User user)
        {
            var monthly = IsMonthlyActive(user);
            return new
            {
                id = user.Id,
                username = user.Username,
                role = user.Role,
                isVip = monthly,
                isMonthly = monthly,
                isGuest = user.Role == "guest",
                vipExpireAt = user.VipExpireAt,
                monthlyExpireAt = user.VipExpireAt,
                balance = user.Balance,
                createdAt = user.CreatedAt
            };
        }

        private void InitDb()
        {
            lock (_lockObj)
            {
                if (!Directory.Exists(_dataDir))
                {
                    Directory.CreateDirectory(_dataDir);
                }

                if (!File.Exists(_storeFile))
                {
                    var adminUser = EnvOr("ADMIN_USERNAME", "admin");
                    var adminPass = EnvOr("ADMIN_PASSWORD", "admin123");
                    var demoUser = EnvOr("DEMO_USERNAME", "testuser");
                    var demoPass = EnvOr("DEMO_PASSWORD", "123456");

                    var defaultStore = new DbStore
                    {
                        Users = new List<User>
                        {
                            new User
                            {
                                Id = "admin-001",
                                Username = adminUser,
                                PasswordHash = HashPassword(adminPass),
                                Role = "admin",
                                IsVip = true,
                                VipExpireAt = 2524608000000,
                                Balance = 99999
                            },
                            new User
                            {
                                Id = "user-demo",
                                Username = demoUser,
                                PasswordHash = HashPassword(demoPass),
                                Role = "user",
                                IsVip = true,
                                VipExpireAt = 2524608000000,
                                Balance = 99999
                            }
                        },
                        VipCards = new List<VipCard>
                        {
                            new VipCard
                            {
                                Code = "VIP-MONTH-8888-9999",
                                Type = "month",
                                Value = 30,
                                Status = "unused"
                            },
                            new VipCard
                            {
                                Code = "VIP-YEAR-6666-8888",
                                Type = "year",
                                Value = 365,
                                Status = "unused"
                            }
                        }
                    };

                    SaveStore(defaultStore);
                }
            }
        }

        public DbStore GetStore()
        {
            lock (_lockObj)
            {
                try
                {
                    if (!File.Exists(_storeFile)) return new DbStore();
                    var json = File.ReadAllText(_storeFile, Encoding.UTF8);
                    return JsonSerializer.Deserialize<DbStore>(json) ?? new DbStore();
                }
                catch
                {
                    return new DbStore();
                }
            }
        }

        public void SaveStore(DbStore store)
        {
            lock (_lockObj)
            {
                var options = new JsonSerializerOptions { WriteIndented = true };
                var json = JsonSerializer.Serialize(store, options);
                File.WriteAllText(_storeFile, json, Encoding.UTF8);
            }
        }

        public User? FindUserByUsername(string username)
        {
            var store = GetStore();
            return store.Users.FirstOrDefault(u => u.Username.Equals(username, StringComparison.OrdinalIgnoreCase));
        }

        public User? FindUserById(string id)
        {
            var store = GetStore();
            return store.Users.FirstOrDefault(u => u.Id == id);
        }

        public User CreateUser(string username, string password)
        {
            var store = GetStore();
            var newUser = new User
            {
                Id = "usr-" + DateTimeOffset.UtcNow.ToUnixTimeMilliseconds() + "-" + Random.Shared.Next(100, 999),
                Username = username,
                PasswordHash = HashPassword(password),
                Role = "user",
                IsVip = false,
                VipExpireAt = 0,
                Balance = 3,
                CreatedAt = DateTimeOffset.UtcNow.ToUnixTimeMilliseconds()
            };

            store.Users.Add(newUser);
            SaveStore(store);
            return newUser;
        }

        public User CreateGuestUser()
        {
            var store = GetStore();
            var trial = 3;
            if (int.TryParse(EnvOr("GUEST_TRIAL_BALANCE", "3"), out var parsed) && parsed >= 0)
            {
                trial = parsed;
            }

            var id = "gst-" + Guid.NewGuid().ToString("N");
            var guest = new User
            {
                Id = id,
                Username = "guest_" + id[^8..],
                PasswordHash = HashPassword(Convert.ToHexString(RandomNumberGenerator.GetBytes(24))),
                Role = "guest",
                IsVip = false,
                VipExpireAt = 0,
                Balance = trial,
                CreatedAt = DateTimeOffset.UtcNow.ToUnixTimeMilliseconds()
            };
            store.Users.Add(guest);
            SaveStore(store);
            return guest;
        }

        public void UpgradePasswordHash(string userId, string password)
        {
            var store = GetStore();
            var user = store.Users.FirstOrDefault(u => u.Id == userId);
            if (user == null) return;
            user.PasswordHash = HashPassword(password);
            SaveStore(store);
        }

        public (bool Success, string Message, string? Code, bool IsMonthly, int Balance, long ExpireAt) ConsumeQuota(string userId, string fileName)
        {
            return ConsumeAiUsage(userId, "process", fileName);
        }

        public (bool Success, string Message, string? Code, bool IsMonthly, int Balance, long ExpireAt) ConsumeAiUsage(string userId, string feature, string? fileName)
        {
            var store = GetStore();
            var user = store.Users.FirstOrDefault(u => u.Id == userId);
            if (user == null) return (false, "用户不存在", "NEED_LOGIN", false, 0, 0);

            var now = DateTimeOffset.UtcNow.ToUnixTimeMilliseconds();
            var monthly = IsMonthlyActive(user);
            var label = feature == "agent" ? "智能助手" : "一键美化";

            if (monthly)
            {
                store.UsageLogs.Add(new UsageLog
                {
                    Id = "log-" + now,
                    UserId = userId,
                    Username = user.Username,
                    FileName = string.IsNullOrWhiteSpace(fileName) ? label : fileName,
                    Type = "包月不扣次-" + label,
                    Cost = 0,
                    CreatedAt = now
                });
                SaveStore(store);
                return (true, "包月有效，不扣次数", null, true, user.Balance, user.VipExpireAt);
            }

            if (user.Balance <= 0)
            {
                return (false, $"{label}次数已用完。请购买次数包，或开通包月后继续使用。", "NEED_RECHARGE", false, 0, 0);
            }

            user.Balance -= 1;
            store.UsageLogs.Add(new UsageLog
            {
                Id = "log-" + now,
                UserId = userId,
                Username = user.Username,
                FileName = string.IsNullOrWhiteSpace(fileName) ? label : fileName,
                Type = "扣除1次-" + label,
                Cost = 1,
                CreatedAt = now
            });
            SaveStore(store);

            return (true, $"已扣除 1 次{label}（剩余 {user.Balance} 次）", null, false, user.Balance, 0);
        }

        public (bool Success, string Message, User? User, VipCard? Card) RedeemCard(string userId, string code)
        {
            var store = GetStore();
            var user = store.Users.FirstOrDefault(u => u.Id == userId);
            if (user == null) return (false, "用户不存在", null, null);
            if (user.Role == "guest")
            {
                return (false, "访客试用账号无法兑换卡密，请先注册或登录正式账号", null, null);
            }

            var card = store.VipCards.FirstOrDefault(c => c.Code.Equals(code.Trim(), StringComparison.OrdinalIgnoreCase) && c.Status == "unused");
            if (card == null) return (false, "无效或已被使用的兑换卡密！", null, null);

            var now = DateTimeOffset.UtcNow.ToUnixTimeMilliseconds();
            card.Status = "redeemed";
            card.UsedBy = user.Username;
            card.UsedAt = now;

            string message;
            if (card.Type == "count")
            {
                user.Balance += card.Value;
                message = $"已充入 {card.Value} 次，可用于一键美化和智能助手。当前剩余 {user.Balance} 次。";
            }
            else
            {
                user.IsVip = true;
                var days = card.Value > 0 ? card.Value : 30;
                var currentExpire = user.VipExpireAt > now ? user.VipExpireAt : now;
                user.VipExpireAt = currentExpire + (long)days * 24 * 60 * 60 * 1000;
                message = $"包月已开通，期内一键美化和智能助手不限次数。有效期至 {DateTimeOffset.FromUnixTimeMilliseconds(user.VipExpireAt).ToLocalTime():yyyy-MM-dd}。";
            }

            SaveStore(store);
            return (true, message, user, card);
        }

        public List<VipCard> GenerateVipCards(int count = 5, string type = "month", int value = 30)
        {
            var store = GetStore();
            var generated = new List<VipCard>();
            var safeCount = Math.Clamp(count, 1, 100);

            for (int i = 0; i < safeCount; i++)
            {
                var randomKey = Convert.ToHexString(RandomNumberGenerator.GetBytes(4));
                var prefix = type.Equals("count", StringComparison.OrdinalIgnoreCase) ? "COUNT" : "MONTH";
                var code = $"{prefix}-{DateTimeOffset.UtcNow.ToUnixTimeMilliseconds().ToString()[^4..]}-{randomKey}";
                var card = new VipCard
                {
                    Code = code,
                    Type = type,
                    Value = value,
                    Status = "unused",
                    CreatedAt = DateTimeOffset.UtcNow.ToUnixTimeMilliseconds()
                };
                store.VipCards.Add(card);
                generated.Add(card);
            }

            SaveStore(store);
            return generated;
        }

        public object GetStats()
        {
            var store = GetStore();
            var now = DateTimeOffset.UtcNow.ToUnixTimeMilliseconds();
            var totalUsers = store.Users.Count(u => u.Role != "guest");
            var activeVips = store.Users.Count(u => u.VipExpireAt > now);
            var totalLogs = store.UsageLogs.Count;
            var unusedCards = store.VipCards.Count(c => c.Status == "unused");

            return new
            {
                totalUsers,
                activeVips,
                activeMonthly = activeVips,
                totalLogs,
                unusedCards,
                cards = store.VipCards,
                users = store.Users.Select(PublicUser).ToList(),
                logs = store.UsageLogs.TakeLast(50)
            };
        }
    }
}
