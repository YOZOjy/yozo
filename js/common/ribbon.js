/**
 *自定义插件执行入口
 * 
 */

//ribbon.xml初始化调用
function OnYozoLoad() {
    setTimeout(() => {
        yozo.getRibbonUI().activateTab("YOZOWorkExtTab");
    }, 3000);
    console.log("Onload")//设置内置button状态
}

/**
* yozo内弹出web页面
* @param {*} html 文件名
* @param {*} title 窗口标题
* @param {*} hight 窗口高
* @param {*} width 窗口宽
*/
function onShowDialog(plugin, html, title, height, width, bModal) {
    var l_ActiveDoc = yozo.WApplication().getActiveDocument();
    if (!l_ActiveDoc) {
        alert("WPS当前没有可操作文档！")
        return;
    }
    if (typeof bModal == "undefined" || bModal == null) {
        bModal = true;
    }
    width *= window.devicePixelRatio;
    height *= window.devicePixelRatio;
    //var url = getHtmlURL(html);
    yozo.showDialog(plugin, html, title, height, width, bModal);
}

/**
* 导入模板到文档中 
*/
function OnImportTemplate(plugin) {
    onShowDialog(plugin, "importTemplate.html", "导入模板", 560, 400, false);
}
/**
 * 历史版本 
 */
function OnHistory(plugin) {
    onShowDialog(plugin, "historyVersion.html", "历史版本", 560, 400, false);
}

//自定义菜单按钮的点击执行事件
function OnAction(controlId, plugin) {
    // console.log("controlId=="+controlId);
    switch (controlId) {
        case "btnShowTaskPane": //打开右侧面板
            //第四个参数：1  右边 ，0  左边； 第五个参数： 面板宽度比设置范围 0 ~ 0.5
            yozo.showTaskPane(plugin, "taskpane.html", "智能推荐",1,0.45);
            break;
        case "btnOpenLocal": //打开本地文档并插入到文档中
            OpenLocalFile();
            break;
        case "btnSaveToServer": //保存到OA服务器
            //获取docid
            var doc = yozo.WApplication().getActiveDocument();
            //根据docid 来获取存好的json字符串
            var param = yozo.getPluginStorage().getItem(doc.getID());
            var uploadPath = JSON.parse(param).uploadPath;
            formUpload(uploadPath, ".doc");
            break;
        case "btnSaveAsFile": //另存为本地文件
            OnBtnSaveAsLocalFile();
            break;
        case "btnChangeToPDF": //转PDF文档并上传
            formUpload(".pdf");
            break;
        case "btnChangeToUOT": //转UOF文档并上传
            formUpload(".uof");
            break;
        case "btnChangeToOFD": //转OFD文档并上传
            formUpload(".ofd");
            break;
        //------------------------------------
        case "btnInsertRedHeader": //插入红头
            onShowDialog(plugin, "redhead.html", "套红", 700, 440, false); //套红头功能
            break;
        case "btnUploadOABackup": //文件备份
            OnUploadOABackupClicked();
            break;
        case "btnInsertSeal": //插入印章
            onShowDialog(plugin, "selectSeal.html", "印章", 700, 440, false);
            break;
        //------------------------------------
        //修订按钮组
        case "btnClearRevDoc": //执行 清稿 按钮
            OnBtnClearRevDoc();
            break;
        case "btnOpenRevision": //打开修订
            {
                //刷新按钮状态
                yozo.getRibbonUI().invalidateControl("btnOpenRevision")
                yozo.getRibbonUI().invalidateControl("btnCloseRevision")
                OnOpenRevisions(); //
                break;
            }
        case "btnCloseRevision": //关闭修订
            {

                yozo.getRibbonUI().invalidateControl("btnOpenRevision")
                yozo.getRibbonUI().invalidateControl("btnCloseRevision")
                OnCloseRevisions();
                break;
            }
        case "btnShowRevision": //显示修订
            break;
        case "btnAcceptAllRevisions": //接受所有修订功能
            OnAcceptAllRevisions();
            break;
        case "btnRejectAllRevisions": //拒绝修订
            OnRejectAllRevisions();
            break;
        //------------------------------------
        case "btnInsertPic": //插入图片
            DoInsertPicToDoc();
            break;
        case "btnInsertWaterMark": //插入水印
            OnInsertWaterMarkClicked();
            break;
        case "btnInsertDate": //插入日期
            OnInsertDateClicked();
            break;
        case "btnOpenScan": //打开扫描仪
            OnOpenScanBtnClicked();
            break;
        case "btnPageSetup": //打开页面设置
            OnPageSetupClicked();
            break;
        case "btnQRCode": //插入二维码
            OnInsertQRCode(plugin); //
            break;
        case "btnPrintDOC": // 打开打印设置
            OnPrintDocBtnClicked();
            break;
        case "btnInsertBookmark": //插入书签
            OnInsertBookmarkToDoc(plugin);
            break;
        case "btnImportTemplate": //导入模板
            OnImportTemplate(plugin);
            break;
        case "btnHistory": //历史版本
            OnHistory(plugin);
            break;
        case "btnSendMessage":
                {
                    /**
                     * 内部封装了主动响应前端发送的请求的方法
                     */
                    let currentTime = new Date()
                    let msgInfo =
                    {
                        id: 1,
                        name: 'YozoSoft'
                    }
                    let msgInfoStr = currentTime.toLocaleString() + ": " + JSON.stringify(msgInfo)
                    msgInfoStr = msgInfoStr.replace(/\"/g,"'")//先用此方法做个应急，202008月版本修复了这个问题
                    yozo.sendWebNotifies(plugin,"我是主动发送的消息， 内容可以自定义。   " + msgInfoStr); //如果想传一个对象，则使用JSON.stringify方法转成对象字符串。
                    break;
                }
        default:
            console.log(controlId + "未有相应的action事件" );
            ;
    }
}

/**
 * 打开插入书签页面
 */
function OnInsertBookmarkToDoc(plugin) {
    if (!yozo.WApplication().getActiveDocument()) {
        return;
    }
    var url = window.location.href;
    console.log(url);

    yozo.showDialog(plugin, "selectBookmark.html", "自定义书签", 700, 440, false);
}

/**
 *  作用：插入二维码图片
 */
function OnInsertQRCode(plugin) {
    //暂不支持插入base64图片
    yozo.showDialog(plugin, "QRCode.html", "插入二维码", 400, 400, false);
}
/**
 * 加载按钮图片
 * @param {*} controlId ribbon.xml 中按钮id
 * @param {*} plugin 
 * @returns 
 */
function GetImage(controlId, plugin) {
    //console.log("js--controlId==="+controlId);
    switch (controlId) {
        case "btnShowTaskPane":
            return "./icon/右侧面板.png"; //打开右侧面板
        case "btnShowDebugPane": //打开网页调试窗口
            return "./icon/debug.png"
        case "btnOpenLocalYOZO": //导入文件
            return "./icon/导入文件.png"
        case "btnSaveToServer": //保存到OA后台服务端
            return "./icon/保存到OA.png";
        case "btnSaveAsFile": //另存为本地文件
            return "./icon/保存本地.png";
        case "btnChangeToPDF": //输出为PDF格式
            return "./icon/转PDF上传.png";
        case "btnChangeToUOT": //
            return "./icon/转UOT上传.png";
        case "btnChangeToOFD": //转OFD上传
            return "./icon/转OFD上传.png"; //
        case "btnInsertRedHeader": //套红头
            return "./icon/套红头.png";
        case "btnInsertSeal": //印章
            return "./icon/印章.png";
        case "btnClearRevDoc": //清稿
            return "./icon/清稿.png"
        case "btnUploadOABackup": //备份正文
            return "./icon/备份正文.png";
        case "btnOpenRevision": //打开 修订
            return "./icon/打开修订.png";
        case "btnShowRevision": //显示修订
        case "btnCloseRevision": //关闭修订
            return "./icon/关闭修订.png";
        case "btnAcceptAllRevisions": // 接受修订
            return "./icon/接受修订.png";
        case "btnRejectAllRevisions": // 拒绝修订
            return "./icon/拒绝修订.png";
        case "btnInsertPic": //插入图片
            return "./icon/插入图片.png";
        case "btnHistory": //历史版本
            return "./icon/备份正文.png";
        case "btnInsertWaterMark": //插入水印
            return "./icon/插入日期.png";
        case "btnOpenScan": //打开扫描仪
            return "./icon/打开扫描仪.png"; //
        case "btnPageSetup": //打开页面设置
            return "./icon/页面设置.png";
        case "btnInsertDate": //插入日期
            return "./icon/插入日期.png";
        case "btnQRCode": //二维码
            return "./icon/二维码.png";
        case "btnPrintDOC": // 打印设置
            return "./icon/打印设置.png";
        case "btnInsertBookmark":
            return "./icon/导入书签.png";
        case "btnImportTemplate":
            return "./icon/导入模板.png";
        default:
            ;
    }
    return "./icon/导入模板.png";
}

/**
 * 作用：处理Ribbon按钮的是否可用
 * @param {*} controlId  ：Ribbon 的按钮控件
 */
function OnGetEnabled(controlId) {
    switch (controlId) {
        case "btnSaveToServer": //保存到OA服务器的相关按钮。判断，如果非OA文件，禁止点击
        case "btnChangeToPDF": //保存到PDF格式再上传
        case "btnChangeToUOT": //保存到UOT格式再上传
        case "btnChangeToOFD": //保存到OFD格式再上传
        case "SaveAll": //保存所有文档
            // return OnSetSaveToOAEnable();
            return true;
        case "btnCloseRevision":
            {
                // let bFlag = yozo.getPluginStorage().getItem(constStrEnum.RevisionEnableFlag)
                return false
            }
        case "btnOpenRevision":
            {
                // let bFlag = yozo.getPluginStorage().getItem(constStrEnum.RevisionEnableFlag)
                return true
            }
        default:
            ;
    }
    return true;
}