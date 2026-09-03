using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using Microsoft.AspNetCore.Mvc;

namespace ExcelTools.Api.Controllers
{
    [ApiController]
    [Route("api/agent")]
    public class AgentController : ControllerBase
    {
        private readonly IHttpClientFactory _http;
        private readonly IConfiguration _config;

        public AgentController(IHttpClientFactory http, IConfiguration config)
        {
            _http = http;
            _config = config;
        }

        public class AgentAskDto
        {
            public string? Message { get; set; }
            public string? Mode { get; set; }
            public JsonElement Snapshot { get; set; }
        }

        [HttpPost("run")]
        public async Task<IActionResult> Run([FromBody] AgentAskDto dto, CancellationToken ct)
        {
            var message = (dto.Message ?? "").Trim();
            if (string.IsNullOrEmpty(message))
                return BadRequest(new { error = "请输入要做的事" });

            var apiKey = FirstNonEmpty(
                Environment.GetEnvironmentVariable("ZHIPU_API_KEY"),
                Environment.GetEnvironmentVariable("ZAI_API_KEY"),
                _config["Zhipu:ApiKey"]);

            if (string.IsNullOrWhiteSpace(apiKey))
                return Ok(new { source = "none", reply = "", actions = Array.Empty<object>() });

            var mode = string.Equals(dto.Mode, "word", StringComparison.OrdinalIgnoreCase) ? "word" : "excel";
            var snapshotJson = dto.Snapshot.ValueKind is JsonValueKind.Undefined or JsonValueKind.Null
                ? "{}"
                : dto.Snapshot.GetRawText();

            var system = mode == "word"
                ? """
                  你是办公助手。用户不会用 Word，请直接改文档，不要让用户去点功能区。
                  只输出 JSON：{"reply":"用大白话说明做了什么","actions":[{"cmd":"命令","payload":任意}]}
                  允许的 cmd：insertText,beautify,redhead,toc,table,checkbox,yesno,radio,dropdown,multiselect,style,bold,bullet,number,align,link
                  insertText 的 payload：{"text":"要写入的正文"}。table 可用 {"rows":3,"cols":3}。
                  radio/dropdown/multiselect 的 payload 带 options 字符串数组，禁止弹窗。
                  style 用 H1/H2/P。align 用 left/center/right/justify。
                  用户说做什么就调用对应命令真正做完。不要编造命令。
                  """
                : """
                  你是办公助手。用户不会用 Excel，看到助手就会让你把活干完。必须调用命令真正改表，禁止只改一两个字或只涂色来应付。
                  只输出 JSON：{"reply":"用大白话说明","actions":[{"cmd":"命令","payload":对象}]}
                  坐标：表头是第0行，数据从第1行起。列 A=0。可以一次输出多条 actions。
                  行=横向 row，列=纵向 col。「第一行/表头」不是第一列。

                  写内容：
                  setCell {"addr":"B2","v":"文本或数字"} 或 {"r":1,"c":1,"v":100} 或 {"addr":"C10","f":"=SUM(C2:C9)"}
                  setCells {"cells":[{...},{...}]}
                  renameHeader {"header":"旧名","to":"新名"} 或 {"col":0,"to":"序号"}

                  结构：
                  insertRow {"at":行索引}  insertCol {"at":列索引,"header":"列名"}  addColumn {"header":"新列"}
                  deleteRow {"at":n}  deleteCol {"at":n}  deleteBlankRows
                  merge {"range":"A1:C1"}  mergeColumns {"c1":3,"c2":4}
                  slashHeader {"r":0,"c":0,"top":"部分","bottom":"序号"} 斜线表头必须用这条，禁止用文字假装。
                  水印必须带上单位或姓名：watermark {"company":"华润集团","person":"张三","tone":"for"}。tone 可用 for（仅供此人使用）、confidential（内部机密）、draft（草案）、custom（payload.text 原句）。禁止只盖「内部机密 严禁外传」。
                  splitColumn {"header":"列名","sep":","}  trimAll
                  fillSeries {"col":0,"r1":1,"r2":maxR,"start":1,"header":"序号"}

                  计算与整理（整行一起动）：
                  sort {"col":列索引或 "header":"部门","dir":"asc|desc"}
                  unique 先不必 select；filterBy {"header":"部门","value":"销售"} 或 {"col":0,"values":["销售"]}
                  filter 只开下拉筛选。unfreeze / freezeTop / freezeLeft
                  autosum/average/max/min/count 可带 {"r1":1,"c1":列,"r2":maxR,"c2":列}，结果写在下方。
                  groupSum {"by":"部门","value":"工资"} 按列分组求和，写在表下方。
                  replace {"find":"旧","replace":"新"}  直接全部替换，不要只打开窗口。

                  格式：
                  fillRow {"row":0,"color":"#b45f06"}  fillColumn {"col":n,"color":"#d9ead3"}
                  fill/bold/align/border/numFmt/fontColor 可带 r1,c1,r2,c2。褐色 #b45f06，浅绿 #d9ead3。
                  highlightIf {"header":"工资","op":"gt|lt|eq|contains","value":100,"color":"#dc2626"}
                  highlightNegatives  numFmt {"r1":1,"c1":n,"r2":maxR,"c2":n,"numFmt":"currency|thousand|percent|int"}
                  dropdown {"r1":1,"c1":n,"r2":maxR,"c2":n,"options":["是","否"]} 不要弹窗。
                  checkbox/yesno 同样带范围。chart {"type":"bar|line"}

                  列名必须对照 headers。用户要筛选、汇总、排序、写公式、插列、改格子时，禁止用 beautify 或涂色代替。
                  """;

            return await AskZhipu(apiKey, system, $"用户说：{message}\n当前快照：{snapshotJson}", ct);
        }

        [HttpPost("beautify")]
        public async Task<IActionResult> Beautify([FromBody] AgentAskDto dto, CancellationToken ct)
        {
            var apiKey = FirstNonEmpty(
                Environment.GetEnvironmentVariable("ZHIPU_API_KEY"),
                Environment.GetEnvironmentVariable("ZAI_API_KEY"),
                _config["Zhipu:ApiKey"]);

            if (string.IsNullOrWhiteSpace(apiKey))
                return Ok(new { source = "none", reply = "", actions = Array.Empty<object>() });

            var mode = string.Equals(dto.Mode, "word", StringComparison.OrdinalIgnoreCase) ? "word" : "excel";
            var snapshotJson = dto.Snapshot.ValueKind is JsonValueKind.Undefined or JsonValueKind.Null
                ? "{}"
                : dto.Snapshot.GetRawText();

            var system = mode == "word"
                ? """
                  你是商务文档美工。根据文档快照只给主题色，不要插入目录、不要加红头（交接/备注/清单不是公文）。
                  只输出 JSON：{"reply":"说明版式","actions":[{"cmd":"beautify","payload":{"headerFill":"#1f4e79"}}]}
                  交接、部署、备注类用藏青 #1f4e79；财务用墨绿 #217346。不要改正文。
                  """
                : """
                  你是Excel报表美工。根据 headers 和 preview 判断这是工资表、台账还是普通清单，给出专业配色。
                  只输出 JSON：{"reply":"说明版式","actions":[{"cmd":"命令","payload":对象}]}
                  第一条必须是 beautify，payload：{"headerFill":"#1f4e79","headerColor":"#ffffff","stripe":"#eef3f8","totalFill":"#fff2cc"}
                  工资/财务用墨绿 #217346；政务/名册用藏青 #1f4e79；深灰 #334155 也可。不要彩虹色。
                  可追加、不要重复 beautify：
                  highlightIf 标出异常（如奖金为负）；freezeTop；styleRange 微调某列对齐或 numFmt。
                  禁止改单元格里的业务文字和数字。禁止对手机号/身份证/序号求和。禁止用 fillColumn 整列纯色盖住斑马纹。
                  """;

            return await AskZhipu(apiKey, system, $"请为下面这份文件做一键专业排版。\n{snapshotJson}", ct);
        }

        private async Task<IActionResult> AskZhipu(string apiKey, string system, string user, CancellationToken ct)
        {
            var model = _config["Zhipu:Model"] ?? "glm-4.6v";
            var body = new Dictionary<string, object?>
            {
                ["model"] = model,
                ["temperature"] = 0.3,
                ["thinking"] = new { type = "disabled" },
                ["messages"] = new object[]
                {
                    new { role = "system", content = system },
                    new { role = "user", content = user }
                }
            };

            try
            {
                var client = _http.CreateClient("zhipu");
                using var req = new HttpRequestMessage(HttpMethod.Post, "api/paas/v4/chat/completions");
                req.Headers.Authorization = new AuthenticationHeaderValue("Bearer", apiKey);
                req.Content = new StringContent(JsonSerializer.Serialize(body), Encoding.UTF8, "application/json");
                using var res = await client.SendAsync(req, ct);
                var raw = await res.Content.ReadAsStringAsync(ct);
                if (!res.IsSuccessStatusCode)
                    return Ok(new { source = "none", reply = "", actions = Array.Empty<object>(), error = raw });

                using var doc = JsonDocument.Parse(raw);
                var msg = doc.RootElement.GetProperty("choices")[0].GetProperty("message");
                var content = "";
                if (msg.TryGetProperty("content", out var cEl) && cEl.ValueKind == JsonValueKind.String)
                    content = cEl.GetString() ?? "";
                if (string.IsNullOrWhiteSpace(content) && msg.TryGetProperty("reasoning_content", out var rEl))
                    content = rEl.GetString() ?? "";

                using var parsed = JsonDocument.Parse(ExtractJson(content));
                var root = parsed.RootElement;
                var reply = root.TryGetProperty("reply", out var r) ? r.GetString() : "";
                object actions = root.TryGetProperty("actions", out var a)
                    ? JsonSerializer.Deserialize<object>(a.GetRawText()) ?? Array.Empty<object>()
                    : Array.Empty<object>();

                return Ok(new { source = "zhipu", model, reply, actions });
            }
            catch
            {
                return Ok(new { source = "none", reply = "", actions = Array.Empty<object>() });
            }
        }

        private static string ExtractJson(string content)
        {
            var t = (content ?? "").Trim();
            var i = t.IndexOf('{');
            var j = t.LastIndexOf('}');
            if (i >= 0 && j > i) return t[i..(j + 1)];
            return "{\"reply\":\"\",\"actions\":[]}";
        }

        private static string? FirstNonEmpty(params string?[] values)
        {
            foreach (var v in values)
            {
                if (!string.IsNullOrWhiteSpace(v)) return v.Trim();
            }
            return null;
        }
    }
}
