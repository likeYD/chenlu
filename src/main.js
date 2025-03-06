/*
 * @Author: lik-m lik-m@glodon.com
 * @Date: 2025-03-06 15:54:21
 * @LastEditors: lik-m lik-m@glodon.com
 * @LastEditTime: 2025-03-06 15:58:28
 * @FilePath: \excelpython\excel-viewer\src\main.js
 * @Description: 这是默认设置,请设置`customMade`, 打开koroFileHeader查看配置 进行设置: https://github.com/OBKoro1/koro1FileHeader/wiki/%E9%85%8D%E7%BD%AE
 */
import { createApp } from 'vue'
import ElementPlus from 'element-plus'
import 'element-plus/dist/index.css'
import App from './App.vue'

const app = createApp(App)
app.use(ElementPlus)
app.mount('#app')
