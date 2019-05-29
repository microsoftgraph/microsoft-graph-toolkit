import Vue from 'vue';
import App from './App.vue';

// import { MgtBaseComponent } from '@microsoft/mgt/dist/es6/components/baseComponent';

// MgtBaseComponent.useShadowRoot = false;

Vue.config.productionTip = false;

new Vue({
  render: (h) => h(App),
}).$mount('#app');
