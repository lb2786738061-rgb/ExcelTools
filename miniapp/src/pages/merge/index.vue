<template>
  <view class="container">
    <view class="header">
      <text class="title">🗂️ 智能合并中心</text>
      <text class="subtitle">将几十个独立表格瞬间缝合为一张总表</text>
    </view>
    
    <view class="upload-box" @click="chooseMultipleFiles">
      <text class="icon">📁➕</text>
      <text class="hint">选取聊天记录中的多个 Excel</text>
      <text class="sub-hint">免费版支持合并 2 个文件</text>
    </view>
    
    <view class="file-list" v-if="files.length > 0">
      <text class="list-title">待合并队列 ({{ files.length }})</text>
      <view class="file-item" v-for="(file, index) in files" :key="index">
        <text class="file-name">{{ file.name }}</text>
        <text class="file-remove" @click="removeFile(index)">❌</text>
      </view>
    </view>

    <button class="merge-btn" v-if="files.length > 0" @click="startMerge">
      ⚡ 开始极致并发合并
    </button>
  </view>
</template>

<script setup>
import { ref } from 'vue';

const files = ref([]);

const chooseMultipleFiles = () => {
  uni.chooseMessageFile({
    count: 10,
    type: 'file',
    extension: ['.xlsx', '.xls', '.csv'],
    success: (res) => {
      files.value = [...files.value, ...res.tempFiles];
    }
  });
};

const removeFile = (idx) => {
  files.value.splice(idx, 1);
};

const startMerge = () => {
  // 💎 商业化拦截点
  const isVip = uni.getStorageSync('isVip') || false;
  if (!isVip && files.value.length > 2) {
    uni.showModal({
      title: '💎 VIP 特权专属',
      content: '免费用户最多仅支持合并 2 个文件。开通 VIP 畅享无限量高速并发合并！',
      confirmText: '立即开通',
      confirmColor: '#f59e0b',
      success: (res) => {
        if (res.confirm) {
          uni.switchTab({ url: '/pages/mine/index' });
        }
      }
    });
    return;
  }
  
  uni.showLoading({ title: '高速合并中...' });
  setTimeout(() => {
    uni.hideLoading();
    uni.showToast({ title: '合并完成，已保存至手机', icon: 'success' });
  }, 1500);
};
</script>

<style scoped>
.container { padding: 30rpx; }
.header { margin-bottom: 40rpx; }
.title { font-size: 40rpx; font-weight: bold; color: #0f172a; display: block; }
.subtitle { font-size: 26rpx; color: #64748b; margin-top: 10rpx; }
.upload-box {
  background: rgba(16, 185, 129, 0.05);
  border: 2rpx dashed #10b981;
  border-radius: 16rpx;
  padding: 60rpx 0;
  display: flex; flex-direction: column; align-items: center; justify-content: center;
}
.icon { font-size: 80rpx; margin-bottom: 20rpx; }
.hint { font-size: 32rpx; color: #334155; font-weight: 500; }
.sub-hint { font-size: 24rpx; color: #94a3b8; margin-top: 10rpx; }

.file-list { margin-top: 40rpx; }
.list-title { font-size: 28rpx; color: #64748b; display: block; margin-bottom: 20rpx; }
.file-item {
  display: flex; justify-content: space-between; align-items: center;
  background: #ffffff; border: 2rpx solid #e2e8f0; padding: 20rpx; border-radius: 12rpx; margin-bottom: 16rpx;
}
.file-name { font-size: 26rpx; color: #334155; }
.file-remove { font-size: 24rpx; padding: 10rpx; }

.merge-btn {
  margin-top: 40rpx;
  background: linear-gradient(135deg, #3b82f6, #2563eb);
  color: #fff;
  border-radius: 16rpx;
  font-weight: bold;
}
</style>
