<template>
  <div>
    <div class="title">Information</div>
    <div class="items">
      <div class="item">
        <div class="name">Path:</div>
        <div class="value">{{ path }}</div>
      </div>
      <div class="item">
        <div class="name">Route Name:</div>
        <div class="value">{{ name }}</div>
      </div>
      <div class="item">
        <div class="name">Vue.js:</div>
        <div class="value">{{ vue }}</div>
      </div>
      <div class="item">
        <div class="name">Electron:</div>
        <div class="value">{{ electron }}</div>
      </div>
      <div class="item">
        <div class="name">Node:</div>
        <div class="value">{{ node }}</div>
      </div>
      <div class="item">
        <div class="name">Platform:</div>
        <div class="value">{{ platform }}</div>
      </div>
    </div>
  </div>
</template>

<script>
  export default {
    data () {
      return {
        electron: process.versions.electron,
        name: this.$route.name,
        node: process.versions.node,
        path: this.$route.path,
        platform: require('os').platform(),
        vue: require('vue/package.json').version
      }
    }
  }
</script>

<style scoped>
  .title {
    color: #667eea;
    font-size: 20px;
    font-weight: 600;
    letter-spacing: .25px;
    margin-top: 10px;
    margin-bottom: 16px;
    padding-bottom: 10px;
    border-bottom: 2px solid #667eea;
  }

  .items { margin-top: 8px; }

  .item {
    display: flex;
    margin-bottom: 12px;
    padding: 10px 12px;
    background: #f8f9fa;
    border-radius: 8px;
    transition: all 0.3s ease;
  }

  .item:hover {
    background: #f0f2f7;
    transform: translateX(4px);
  }

  .item .name {
    color: #6a6a6a;
    margin-right: 12px;
    font-weight: 500;
    min-width: 100px;
  }

  .item .value {
    color: #667eea;
    font-weight: 600;
  }
</style>
