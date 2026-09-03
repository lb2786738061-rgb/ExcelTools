export function createKnowledgePanel(onRequestVip) {
  const container = document.createElement('div');
  container.className = 'knowledge-panel';
  
  container.innerHTML = `
    <div class="knowledge-header">
      <div class="knowledge-title-wrapper">
        <h1 class="knowledge-main-title">Excel 极速公式字典与提效宝典</h1>
        <p class="knowledge-subtitle">全场景求和公式/函数大全（覆盖 99% 工作需求）</p>
      </div>
      <div class="knowledge-stats">
        <span class="stat-badge">🔥 热门收藏 12.5w+</span>
        <span class="stat-badge">⏱️ 预估阅读 15 分钟</span>
      </div>
    </div>
    
    <div class="knowledge-content">
      <!-- 基础与进阶免费展示区 -->
      <section class="knowledge-section">
        <h2>一、基础核心求和（最常用，入门必学）</h2>
        <div class="func-card">
          <div class="func-header">
            <h3>1. SUM 函数（最基础的无条件求和）</h3>
            <span class="func-tag basic">基础</span>
          </div>
          <p><strong>核心功能</strong>：对指定的单元格、单元格区域、常量进行求和，是所有求和场景的基础。</p>
          <div class="code-block">
            <div class="code-header">语法</div>
            <code>SUM(数值1, [数值2], [数值3], ...)</code>
          </div>
          <div class="func-example">
            <h4>💡 实用示例：</h4>
            <pre><code>=SUM(A1:A10)  // 对A1到A10的连续单元格求和
=SUM(A1:A10,C1:C10)  // 同时对A列1-10行、C列1-10行两个区域求和
=SUM(10,20,30)  // 对固定数字求和，结果为60</code></pre>
          </div>
        </div>
      </section>

      <section class="knowledge-section">
        <h2>二、条件求和（按规则筛选后求和，工作中80%的求和需求都用这类）</h2>
        
        <div class="func-card">
          <div class="func-header">
            <h3>1. SUMIF 函数（单条件求和）</h3>
            <span class="func-tag standard">进阶</span>
          </div>
          <p><strong>核心功能</strong>：对满足<strong>1个条件</strong>的单元格区域求和。</p>
          <div class="code-block">
            <div class="code-header">语法</div>
            <code>SUMIF(条件区域, 条件, [求和区域])</code>
          </div>
          <div class="func-example">
            <h4>💡 实用示例：</h4>
            <pre><code>=SUMIF(B:B,"销售部",C:C)  // 求B列中为“销售部”的所有行，C列的业绩总和
=SUMIF(A:A,">100",A:A)  // 求A列中大于100的数值总和
=SUMIF(A:A,"*苹果*",B:B)  // 求A列中包含“苹果”关键词的行，B列的销量总和</code></pre>
          </div>
        </div>

        <div class="func-card">
          <div class="func-header">
            <h3>2. SUMIFS 函数（多条件求和）</h3>
            <span class="func-tag standard">进阶</span>
          </div>
          <p><strong>核心功能</strong>：对<strong>同时满足多个条件</strong>的单元格区域求和，是 SUMIF 的多条件增强版。</p>
          <div class="code-block">
            <div class="code-header">语法</div>
            <code>SUMIFS(求和区域, 条件区域1, 条件1, 条件区域2, 条件2, ...)</code>
          </div>
          <p class="warning-text">⚠️ 注意：求和区域必须放在最前面，和 SUMIF 顺序相反，这是核心易错点！</p>
          <div class="func-example">
            <h4>💡 实用示例：</h4>
            <pre><code>=SUMIFS(D:D,B:B,"销售部",C:C,"1月")  // 求销售部1月的业绩总和（2个条件）
=SUMIFS(A:A,A:A,">100",A:A,"&lt;500")  // 求A列中大于100且小于500的数值总和</code></pre>
          </div>
        </div>
      </section>

      <section class="knowledge-section">
        <h2>三、数组/乘积求和（复杂计算、权重、多条件兼容场景）</h2>
        <div class="func-card">
          <div class="func-header">
            <h3>1. SUMPRODUCT 函数（乘积求和，全能型求和函数）</h3>
            <span class="func-tag standard">进阶</span>
          </div>
          <p><strong>核心功能</strong>：将多个数组的对应元素相乘，再对乘积求和；可替代SUMIFS实现多条件求和，兼容性更好。</p>
          <div class="code-block">
            <div class="code-header">语法</div>
            <code>SUMPRODUCT(数组1, [数组2], [数组3], ...)</code>
          </div>
          <div class="func-example">
            <h4>💡 实用示例：</h4>
            <pre><code>// 基础用法：加权求和，求总销售额（数量列×单价列）
=SUMPRODUCT(B1:B10,C1:C10) 

// 多条件求和用法（替代SUMIFS）：求销售部1月的业绩总和
=SUMPRODUCT((B:B="销售部")*(C:C="1月")*D:D)</code></pre>
          </div>
        </div>
      </section>

      <!-- 高级内容：完全开放展示区 -->
      <section class="knowledge-section">
        <h2>四、筛选/可见单元格求和（隐藏行/筛选后专用，避坑必学）</h2>
        <div class="func-card">
          <div class="func-header">
            <h3>1. SUBTOTAL 函数（忽略隐藏行/筛选后求和）</h3>
            <span class="func-tag standard">进阶</span>
          </div>
          <p><strong>核心功能</strong>：对可见单元格求和，自动忽略手动隐藏的行、筛选掉的行；SUM函数会包含隐藏行的数值，这是核心区别。</p>
          <div class="code-block">
            <div class="code-header">语法</div>
            <code>SUBTOTAL(109, C:C)</code>
          </div>
        </div>
        
        <h2>五、Excel 365 动态数组求和（新版专属，更灵活）</h2>
        <div class="func-card">
          <div class="func-header">
            <h3>1. FILTER+SUM 动态条件求和</h3>
            <span class="func-tag standard">高阶</span>
          </div>
          <p><strong>核心功能</strong>：先按条件筛选出符合要求的区域，再对筛选结果求和，支持动态更新和超复杂逻辑嵌套。</p>
          <div class="code-block">
            <div class="code-header">语法</div>
            <code>=SUM(FILTER(D:D,(B:B="销售部")*(C:C="1月")))</code>
          </div>
        </div>
      </section>

      <section class="knowledge-section">
        <h2>💡 核心避坑要点</h2>
        <ul class="pitfall-list">
          <li><strong>SUMIF 和 SUMIFS 参数顺序相反</strong>：SUMIF是「条件区域→条件→求和区域」，SUMIFS是「求和区域→条件区域1→条件1」，写反会直接出错。</li>
          <li><strong>筛选后求和不要用 SUM</strong>：SUM会包含隐藏行/筛选掉的行，必须用 <code>SUBTOTAL(109,...)</code>。</li>
          <li><strong>条件里的文本要加英文双引号</strong>：比如 <code>"销售部"</code>，不能用中文引号，否则会报错。</li>
          <li><strong>多条件求和的条件区域维度必须一致</strong>：SUMIFS/SUMPRODUCT的所有条件区域、求和区域，必须是相同的行数/列数。</li>
        </ul>
      </section>

    </div>
  `;

  const btnUnlockVip = container.querySelector('#btnUnlockVip');
  if (btnUnlockVip && onRequestVip) {
    btnUnlockVip.addEventListener('click', () => {
      onRequestVip();
    });
  }

  return container;
}
