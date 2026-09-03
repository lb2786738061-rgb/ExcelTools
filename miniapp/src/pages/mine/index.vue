<template>
  <view class="container">
    <view class="user-info-card">
      <view class="avatar">👤</view>
      <view class="info">
        <text class="name">微信用户</text>
        <text class="status" v-if="!isVip">普通体验用户</text>
        <text class="status vip-status" v-else>👑 尊贵的 VIP 会员</text>
      </view>
    </view>
    
    <view class="vip-banner" v-if="!isVip">
      <view class="vip-content">
        <text class="vip-title">💎 开通商业全栈 VIP</text>
        <text class="vip-desc">解锁无限量多表并发合并，解除高阶智能分析导出限制</text>
      </view>
      <button class="pay-btn" @click="handleRecharge">立即开通 ￥99/年</button>
    </view>
    
    <view class="menu-list">
      <view class="menu-item" @click="showComingSoon">
        <text>⚙️ 通用设置</text>
        <text class="arrow">></text>
      </view>
      <view class="menu-item" @click="showComingSoon">
        <text>🧾 开具发票</text>
        <text class="arrow">></text>
      </view>
      <view class="menu-item" @click="showComingSoon">
        <text>🎧 在线客服</text>
        <text class="arrow">></text>
      </view>
    </view>
  </view>
</template>

<script setup>
import { ref } from 'vue';
import { onShow } from '@dcloudio/uni-app';

const isVip = ref(false);

onShow(() => {
  isVip.value = uni.getStorageSync('isVip') || false;
});

const handleRecharge = () => {
  uni.showLoading({ title: '拉起微信支付...' });
  setTimeout(() => {
    uni.hideLoading();
    // 模拟充值成功
    uni.setStorageSync('isVip', true);
    isVip.value = true;
    uni.showToast({
      title: '👑 充值成功，欢迎回来老板！',
      icon: 'none',
      duration: 3000
    });
  }, 1500);
};

const showComingSoon = () => {
  uni.showToast({ title: '系统维护中，暂未开放', icon: 'none' });
};
</script>

<style scoped>
.container { background: #f8fafc; min-height: 100vh; padding: 20rpx; }
.user-info-card {
  background: #ffffff; border-radius: 16rpx; padding: 40rpx 30rpx;
  display: flex; align-items: center; gap: 30rpx; margin-bottom: 30rpx;
  box-shadow: 0 4rpx 12rpx rgba(0,0,0,0.02);
}
.avatar {
  width: 100rpx; height: 100rpx; background: #e2e8f0; border-radius: 50rpx;
  display: flex; align-items: center; justify-content: center; font-size: 50rpx;
}
.info { display: flex; flex-direction: column; }
.name { font-size: 34rpx; font-weight: bold; color: #1e293b; margin-bottom: 8rpx; }
.status { font-size: 24rpx; color: #64748b; }
.vip-status { color: #d97706; font-weight: bold; }

.vip-banner {
  background: linear-gradient(135deg, #1e293b, #0f172a);
  border-radius: 16rpx; padding: 40rpx 30rpx;
  display: flex; flex-direction: column; align-items: center; margin-bottom: 40rpx;
}
.vip-content { text-align: center; margin-bottom: 30rpx; }
.vip-title { color: #fbbf24; font-size: 36rpx; font-weight: bold; display: block; margin-bottom: 12rpx; }
.vip-desc { color: #94a3b8; font-size: 24rpx; }
.pay-btn {
  background: linear-gradient(135deg, #f59e0b, #d97706);
  color: #fff; border-radius: 50rpx; padding: 0 60rpx; font-size: 28rpx; font-weight: bold;
}

.menu-list { background: #ffffff; border-radius: 16rpx; padding: 0 30rpx; }
.menu-item {
  display: flex; justify-content: space-between; align-items: center;
  padding: 30rpx 0; border-bottom: 1rpx solid #f1f5f9; font-size: 28rpx; color: #334155;
}
.menu-item:last-child { border-bottom: none; }
.arrow { color: #cbd5e1; font-weight: bold; }
</style>
