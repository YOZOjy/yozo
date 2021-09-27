### Demo介绍

这个Demo为永中文字常见的OA场景集成下的永中加载项——OA助手，项目使用了丰富的永中 API的功能，可以帮助大家能够快速理解并熟悉永中加载项机制以及和浏览器调用交互的流程。

### 工程结构

* icon 						图标文件。
* js						永中 加载项功能逻辑的js代码。
* dispatcherinfo.js         集中处理来自OA的传入的参数(打开文档携带参数)
* ribbon.js                 执行插件入口
* yozoCommon.js             yozojsAPI通用方法封装。
* otherslib					vue,jQueryd等第三方库。
* hisortyVersion.html				历史版本。
* importTemplate.html		导入模板页面。
* ribbon.html				加载项的默认加载页面。
* qrcode.html				插入二维码页面。
* redhead.html				插入红头页面。 
* selectSeal.html			插入签章页面。 
* ribbon.xml				自定义选项卡配置。 

### 注意事项

* 本工程只是演示demo
* 我们建议您结合具体的应用场景修改示例代码，这样更能够体现OA助手集成的应用场景
* 为了保护代码，建议对上线代码进行混淆

