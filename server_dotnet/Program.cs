using ExcelTools.Api.Services;

var builder = WebApplication.CreateBuilder(args);

// 配置监听端口：优先环境变量 PORT，无则默认 5001
var portEnv = Environment.GetEnvironmentVariable("PORT");
if (!string.IsNullOrEmpty(portEnv) && int.TryParse(portEnv, out int customPort))
{
    builder.WebHost.UseUrls($"http://0.0.0.0:{customPort}");
}
else
{
    builder.WebHost.UseUrls("http://0.0.0.0:5001");
}

// 注册 CORS 跨域策略
builder.Services.AddCors(options =>
{
    options.AddPolicy("AllowAll", policy =>
    {
        policy.AllowAnyOrigin()
              .AllowAnyHeader()
              .AllowAnyMethod();
    });
});

builder.Services.AddControllers();
builder.Services.AddSingleton<TokenService>();
builder.Services.AddSingleton<DatabaseService>();
builder.Services.AddHttpClient("zhipu", client =>
{
    client.BaseAddress = new Uri("https://open.bigmodel.cn/");
    client.Timeout = TimeSpan.FromSeconds(90);
});

var app = builder.Build();

// 应用 CORS 与 路由
app.UseCors("AllowAll");

app.UseAuthorization();
app.MapControllers();

Console.WriteLine("=========================================================================");
Console.WriteLine("🚀 商业化 C# (.NET 8.0 / Web API) 后端服务器已启动：http://localhost:5001");
Console.WriteLine("=========================================================================");

app.Run();
