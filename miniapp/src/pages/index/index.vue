<template>
  <view class="container">
    <view class="header">
      <text class="title">一键智能规整与美化</text>
      <text class="subtitle">告别加班，全自动生成精美报表</text>
    </view>
    
    <view class="upload-box" @click="chooseFile">
      <text class="icon">📂</text>
      <text class="hint">点击选择微信聊天记录中的文件</text>
      <text class="sub-hint">支持 .xlsx / .csv</text>
    </view>
    
    <!-- 后续完善核心控制面板组件 -->
  </view>
</template>

<script setup>
import { ref } from 'vue';

const chooseFile = () => {
  // 调用微信原生的聊天文件拾取器 (H5 环境下可降级处理)
  uni.chooseMessageFile({
    count: 1,
    type: 'file',
    extension: ['.xlsx', '.xls', '.csv'],
    success: (res) => {
      console.log('用户选中的文件:', res.tempFiles[0]);
      uni.showToast({
        title: '已获取文件，准备智能解析',
        icon: 'none'
      });
      // TODO: 接入底层的 excelEngine 进行内存解析
    }
  });
};
</script>

<style scoped>
.container {
  padding: 30rpx;
}
.header {
  margin-bottom: 40rpx;
}
.title {
  font-size: 40rpx;
  font-weight: bold;
  color: #0f172a;
  display: block;
}
.subtitle {
  font-size: 26rpx;
  color: #64748b;
  margin-top: 10rpx;
}
.upload-box {
  background: rgba(59, 130, 246, 0.05);
  border: 2rpx dashed #3b82f6;
  border-radius: 16rpx;
  padding: 60rpx 0;
  display: flex;
  flex-direction: column;
  align-items: center;
  justify-content: center;
}
.icon {
  font-size: 80rpx;
  margin-bottom: 20rpx;
}
.hint {
  font-size: 32rpx;
  color: #334155;
  font-weight: 500;
}
.sub-hint {
  font-size: 24rpx;
  color: #94a3b8;
  margin-top: 10rpx;
}
</style>
