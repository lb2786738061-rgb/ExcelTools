/**
 * ShareModal 微信与飞书快捷分享模态框组件
 * 模拟选择分享至微信联系人/微信群、飞书联系人/飞书云文档
 */

export function createShareModal() {
  const backdrop = document.createElement('div');
  backdrop.style.cssText = `
    position: fixed;
    top: 0; left: 0; right: 0; bottom: 0;
    background: rgba(15, 23, 42, 0.75);
    backdrop-filter: blur(8px);
    display: none;
    align-items: center;
    justify-content: center;
    z-index: 100;
  `;

  backdrop.innerHTML = `
    <div style="
      background: var(--bg-panel);
      border: 1px solid var(--border-color);
      border-radius: var(--radius-lg);
      width: 440px;
      max-width: 90vw;
      padding: 1.5rem;
      box-shadow: var(--shadow-lg);
    ">
      <div style="display: flex; align-items: center; justify-content: space-between; margin-bottom: 1.25rem;">
        <h3 style="font-size: 1.1rem; color: var(--text-main);">一键分享至社群平台</h3>
        <button id="closeModalBtn" style="background: none; border: none; color: var(--text-muted); cursor: pointer; font-size: 1.2rem;">✕</button>
      </div>

      <p style="font-size: 0.85rem; color: var(--text-muted); margin-bottom: 1rem;">选择分享目标平台与发送形式：</p>

      <div style="display: flex; flex-direction: column; gap: 0.75rem; margin-bottom: 1.5rem;">
        <div class="share-option" data-target="wechat" style="
          display: flex; align-items: center; justify-content: space-between;
          padding: 0.8rem 1rem; background: rgba(15, 23, 42, 0.6);
          border: 1px solid var(--border-color); border-radius: var(--radius-md);
          cursor: pointer; transition: var(--transition-fast);
        ">
          <div style="display: flex; align-items: center; gap: 0.75rem;">
            <div style="width: 36px; height: 36px; background: #07c160; border-radius: 8px; display: flex; align-items: center; justify-content: center; color: white; font-weight: bold;">微</div>
            <div>
              <div style="font-weight: 600; font-size: 0.95rem; color: var(--text-main);">分享给微信好友 / 微信群</div>
              <div style="font-size: 0.75rem; color: var(--text-muted);">发送 Office 原文件或高清长图</div>
            </div>
          </div>
          <span style="color: var(--primary-500); font-size: 0.85rem;">发送 ›</span>
        </div>

        <div class="share-option" data-target="feishu" style="
          display: flex; align-items: center; justify-content: space-between;
          padding: 0.8rem 1rem; background: rgba(15, 23, 42, 0.6);
          border: 1px solid var(--border-color); border-radius: var(--radius-md);
          cursor: pointer; transition: var(--transition-fast);
        ">
          <div style="display: flex; align-items: center; gap: 0.75rem;">
            <div style="width: 36px; height: 36px; background: #3370ff; border-radius: 8px; display: flex; align-items: center; justify-content: center; color: white; font-weight: bold;">飞</div>
            <div>
              <div style="font-weight: 600; font-size: 0.95rem; color: var(--text-main);">发送至飞书联系人 / 群组</div>
              <div style="font-size: 0.75rem; color: var(--text-muted);">发送至对话框或转存至飞书云文档</div>
            </div>
          </div>
          <span style="color: var(--primary-500); font-size: 0.85rem;">发送 ›</span>
        </div>
      </div>

      <div style="font-size: 0.75rem; color: var(--text-dim); text-align: center;">
        在手机客户端 APP 中支持一键调起调唤系统分享通道
      </div>
    </div>
  `;

  const closeBtn = backdrop.querySelector('#closeModalBtn');
  closeBtn.addEventListener('click', () => hide());

  backdrop.querySelectorAll('.share-option').forEach(item => {
    item.addEventListener('click', (e) => {
      const target = e.currentTarget.dataset.target;
      const name = target === 'wechat' ? '微信' : '飞书';
      alert(`已成功将美化文档发送至 ${name}！在移动端 APP 中将直接唤起 ${name} 聊天选择窗口。`);
      hide();
    });
  });

  function show() {
    backdrop.style.display = 'flex';
  }

  function hide() {
    backdrop.style.display = 'none';
  }

  return {
    element: backdrop,
    show,
    hide
  };
}
