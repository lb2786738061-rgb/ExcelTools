<template>
  <view class="container">
    <view class="header">
      <text class="title">📕 极速公式字典</text>
      <text class="subtitle">全场景求和函数速查，覆盖99%工作需求</text>
    </view>
    
    <view class="card">
      <view class="card-header">
        <text class="func-name">1. SUMIF (单条件求和)</text>
        <text class="tag">进阶</text>
      </view>
      <text class="desc">核心功能：对满足 1 个条件的单元格区域求和。</text>
      <view class="code-box">
        <text class="code">SUMIF(条件区域, 条件, [求和区域])</text>
      </view>
      <text class="example">💡 示例: =SUMIF(B:B,"销售部",C:C)</text>
    </view>
    
    <view class="card">
      <view class="card-header">
        <text class="func-name">2. SUMIFS (多条件求和)</text>
        <text class="tag">进阶</text>
      </view>
      <text class="desc">核心功能：对同时满足多个条件的单元格求和。</text>
      <view class="code-box">
        <text class="code">SUMIFS(求和区域, 条件1, 值1, ...)</text>
      </view>
      <text class="example">💡 示例: =SUMIFS(D:D,B:B,"销售部",C:C,"1月")</text>
    </view>
    
    <view class="card">
      <view class="card-header">
        <text class="func-name">3. SUBTOTAL (筛选可见求和)</text>
        <text class="tag high">高阶</text>
      </view>
      <text class="desc">核心功能：忽略隐藏行，仅对筛选后的可见数据求和。</text>
      <view class="code-box">
        <text class="code">SUBTOTAL(109, C:C)</text>
      </view>
      <text class="example">⚠️ SUM函数会包含隐藏行，这是最大区别！</text>
    </view>
  </view>
</template>

<style scoped>
.container { padding: 30rpx; background: #f8fafc; min-height: 100vh; }
.header { margin-bottom: 40rpx; }
.title { font-size: 40rpx; font-weight: bold; color: #0f172a; display: block; }
.subtitle { font-size: 26rpx; color: #64748b; margin-top: 10rpx; }
.card {
  background: #ffffff;
  border-radius: 16rpx;
  padding: 30rpx;
  margin-bottom: 24rpx;
  box-shadow: 0 4rpx 12rpx rgba(0,0,0,0.03);
}
.card-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 16rpx; }
.func-name { font-size: 32rpx; font-weight: bold; color: #1e293b; }
.tag { font-size: 22rpx; padding: 4rpx 12rpx; background: #e0f2fe; color: #0284c7; border-radius: 8rpx; }
.tag.high { background: #fef3c7; color: #d97706; }
.desc { font-size: 26rpx; color: #475569; line-height: 1.5; display: block; margin-bottom: 20rpx; }
.code-box { background: #1e293b; padding: 20rpx; border-radius: 8rpx; margin-bottom: 20rpx; }
.code { color: #38bdf8; font-family: monospace; font-size: 26rpx; }
.example { font-size: 24rpx; color: #10b981; font-weight: 500; }
</style>
