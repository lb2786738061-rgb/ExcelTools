

/**
 * Toast 消息通知组件
 * 劫持原生 window.alert，提供更加现代和精美的 UI 体验。
 */

export function showToast(message, type = 'info') {
  // 智能识别类型（如果是从原生 alert 劫持过来的，可能没有传 type）
  if (type === 'info') {
    if (/成功|完成|🎉/.test(message)) type = 'success';
    else if (/失败|错误|未登录|请先|仅支持/.test(message)) type = 'error';
    else if (/VIP|💎/.test(message)) type = 'warning';
  }

  const containerId = 'toast-container';
  let container = document.getElementById(containerId);
  if (!container) {
    container = document.createElement('div');
    container.id = containerId;
    container.style.cssText = `
      position: fixed;
      top: 20px;
      left: 50%;
      transform: translateX(-50%);
      z-index: 9999;
      display: flex;
      flex-direction: column;
      gap: 10px;
      pointer-events: none;
    `;
    document.body.appendChild(container);
  }

  const toast = document.createElement('div');
  
  // 基础样式
  let bg = '#ffffff';
  let border = '#e2e8f0';
  let icon = '💡';
  let color = '#1e293b';

  if (type === 'success') {
    bg = '#ecfdf5';
    border = '#10b981';
    icon = '✅';
    color = '#065f46';
  } else if (type === 'error') {
    bg = '#fef2f2';
    border = '#ef4444';
    icon = '❌';
    color = '#991b1b';
  } else if (type === 'warning') {
    bg = '#fffbeb';
    border = '#f59e0b';
    icon = '💎';
    color = '#92400e';
  }

  // 这里的样式模仿 macOS / iOS 通知风格，带有磨砂玻璃效果和阴影
  toast.style.cssText = `
    background: ${bg};
    border-left: 4px solid ${border};
    color: ${color};
    padding: 12px 20px;
    border-radius: 8px;
    box-shadow: 0 10px 25px -5px rgba(0, 0, 0, 0.1), 0 8px 10px -6px rgba(0, 0, 0, 0.1);
    font-size: 0.95rem;
    font-weight: 500;
    display: flex;
    align-items: center;
    gap: 10px;
    transform: translateY(-20px);
    opacity: 0;
    transition: all 0.3s cubic-bezier(0.16, 1, 0.3, 1);
    max-width: 400px;
    line-height: 1.5;
  `;

  // 处理多行文本
  const textHtml = message.replace(/\n/g, '<br/>');
  
  toast.innerHTML = `<span style="font-size: 1.2rem;">${icon}</span> <span>${textHtml}</span>`;
  container.appendChild(toast);

  // 触发滑入动画
  requestAnimationFrame(() => {
    toast.style.transform = 'translateY(0)';
    toast.style.opacity = '1';
  });

  // 3秒后移除
  setTimeout(() => {
    toast.style.transform = 'translateY(-20px)';
    toast.style.opacity = '0';
    setTimeout(() => {
      if (toast.parentNode) toast.parentNode.removeChild(toast);
    }, 300);
  }, 3500);
}

// 劫持原生 window.alert
const originalAlert = window.alert;
window.alert = function(message) {
  showToast(String(message));
};
