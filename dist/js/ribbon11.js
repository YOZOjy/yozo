document.write("<script language='javascript' src='js/yozoCommon.js'></script>");

//自定义菜单按钮的点击执行事件
function OnAction(controlId,plugin) {
  console.log("controlId=="+controlId);
  switch (controlId) {
      case "btnShowTaskPane": //打开WPS云文档入口
            yozo.showTaskPane(plugin,"http://123.60.5.149/plugin-ai-assistant","oa面板");
          //yozo.showDialog("http://localhost:9222","sss",500,600,true);
          break;
      case "btnOpenLocalWPSYUN": //打开本地文档并插入到文档中
            OpenLocalFile();
          break;
      case "WPSWorkExtTab":
            OnbtnTabClick();
          break;
      case "btnSaveToServer": //保存到OA服务器
          //wps.PluginStorage.setItem(constStrEnum.Save2OAShowConfirm, true)
          //OnBtnSaveToServer();
            formUpload("");
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
          OnInsertRedHeaderClick(); //套红头功能
          break;
      case "btnUploadOABackup": //文件备份
          OnUploadOABackupClicked();
          break;
      case "btnInsertSeal": //插入印章
          OnInsertSeal();
          break;
          //------------------------------------
          //修订按钮组
      case "btnClearRevDoc": //执行 清稿 按钮
          OnBtnClearRevDoc();
          break;
      case "btnOpenRevision": //打开修订
          {
              //let bFlag = wps.PluginStorage.getItem(constStrEnum.RevisionEnableFlag)
              //wps.PluginStorage.setItem(constStrEnum.RevisionEnableFlag, !bFlag)
              //通知wps刷新以下几个按钮的状态
            //  wps.ribbonUI.InvalidateControl("btnOpenRevision")
              //wps.ribbonUI.InvalidateControl("btnCloseRevision")
              OnOpenRevisions(); //
              break;
          }
      case "btnCloseRevision": //关闭修订
          {
             // let bFlag = wps.PluginStorage.getItem(constStrEnum.RevisionEnableFlag)
             // wps.PluginStorage.setItem(constStrEnum.RevisionEnableFlag, !bFlag)
              //通知wps刷新以下几个按钮的状态
             // wps.ribbonUI.InvalidateControl("btnOpenRevision")
             // wps.ribbonUI.InvalidateControl("btnCloseRevision")
              OnCloseRevisions();
              break;
          }
      case "btnShowRevision":
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
          OnInsertQRCode(); //
          break;
      case "btnPrintDOC": // 打开打印设置
          OnPrintDocBtnClicked();
          break;
      case "lblDocSourceValue": //OA公文提示
          OnOADocInfo();
          break;
      case "btnUserName": //点击用户
          OnUserNameSetClick();
          break;
      case "btnInsertBookmark": //插入书签
          OnInsertBookmarkToDoc(plugin);
          break;
      case "btnImportTemplate": //导入模板
          OnImportTemplate();
          break;
      // case "FileSaveAsMenu": //通过idMso进行「另存为」功能的自定义
      // case "FileSaveAs":
      case "FileSave": //通过idMso进行「保存」功能的自定义
          {
              if (pCheckIfOADoc()) { //文档来源是业务系统的，做自定义
                  alert("这是OA文档，将Ctrl+S动作做了重定义，可以调用OA的保存文件流到业务系统的接口。")
              } else { //本地的文档，期望不做自定义，通过转调idMso的方法实现
                  // wps.WpsApplication().CommandBars.ExecuteMso("FileSave");
                  wps.WpsApplication().CommandBars.ExecuteMso("SaveAll");
                  //此处一定不能去调用与重写idMso相同的ID，否则就是个无线递归了，即在这个场景下不可调用FileSaveAs和FileSaveAsMenu这两个方法
              }
              break;
          }
      case "ShowAlert_ContextMenuText":
          {
              let selectText = wps.WpsApplication().Selection.Text;
              alert("您选择的内容是：\n" + selectText);
              break;
          }
      case "btnSendMessage":
          {
              /**
               * 内部封装了主动响应前端发送的请求的方法
               */
              let currentTime = new Date()
              let msgInfo =
              {
                  id: 1,
                  name: 'kingsoft',
                  since: "1988"
              }
              let msgInfoStr = currentTime.toLocaleString() + ": " + JSON.stringify(msgInfo)
              msgInfoStr = msgInfoStr.replace(/\"/g,"'")//先用此方法做个应急，202008月版本修复了这个问题
              wps.OAAssist.WebNotify("我是主动发送的消息， 内容可以自定义。   " + msgInfoStr); //如果想传一个对象，则使用JSON.stringify方法转成对象字符串。
              // wps.OAAssist.WebNotify("我是主动发送的消息，内容可以自定义。如果想传一个对象，则使用JSON.stringify方法转成对象字符串。当前时间是："+ currentTime.toLocaleString()); //如果想传一个对象，则使用JSON.stringify方法转成对象字符串。
              break;
          }
      default:
          break;
  }
}

/**
 * 二维码
 */

/**
 * 
 * 套红
 * OnInsertRedHeaderClick()
{
    yozo.showDialog("D:/VisuslCode/noplugin/hello/dist/redhead.html", "套红", 700, 440, false);

}

/**
 * 
 * 签章
 */

function OnInsertSeal()
{
    yozo.showDialog(plugin,"selectSeal.html", "印章", 700, 440, false);
}
/**
 * 打开插入书签页面
 */
 function OnInsertBookmarkToDoc(plugin) {
    if (!yozo.WApplication().getActiveDocument()) {
        alert("没有文档");
        return;
    }
    var url = window.location.href;
    console.log(url);
    yozo.showDialog(plugin,"selectBookmark.html", "自定义书签", 700, 440, false);
}


function GetImage(controlId,plugin){
  console.log("js--controlId==="+controlId);
    switch (controlId) {
      case "btnShowTaskPane":
          return "./icon/右侧面板.png"; //打开WPS云文档
      case "btnOpenLocalWPSYUN": //导入文件
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
      case "btnShowRevision": //
          return "./icon/打开修订.png";
      case "btnCloseRevision": //关闭修订
          return "./icon/关闭修订.png";
      case "btnAcceptAllRevisions": // 接受修订
          return "./icon/接受修订.png";
      case "btnRejectAllRevisions": // 拒绝修订
          return "./icon/拒绝修订.png";
      case "btnSaveAsFile":
          return "";
      case "btnInsertPic": //插入图片
          return "./icon/插入图片.png";
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
      case "btnSendMessage":
          return "./icon/3.svg"
      case "btnDebug":
          return "./icon/导入模板.png"
      default:
          ;
  }
  return "./icon/c_default.png";
}


function OnGetVisible(controlId,plugin)
  {
    switch (controlId) {
        case "btnShowTaskPane":
            return true;//打开WPS云文档
        case "btnOpenLocalWPSYUN": //导入文件
        return true;
        case "btnSaveToServer": //保存到OA后台服务端
        return true;
        case "btnSaveAsFile": //另存为本地文件
        return true;
        case "btnChangeToPDF": //输出为PDF格式
        return true;
        case "btnChangeToUOT": //
        return true;
        case "btnChangeToOFD": //转OFD上传
        return true;
        case "btnInsertRedHeader": //套红头
        return true;
        case "btnInsertSeal": //印章
        return true;
        case "btnClearRevDoc": //清稿
           return true;
        case "btnUploadOABackup": //备份正文
            return true;
        case "btnOpenRevision": //打开 修订
           return true;
        case "btnShowRevision": //
            return true;
        case "btnCloseRevision": //关闭修订
            return true;
        case "btnAcceptAllRevisions": // 接受修订
            return true;
        case "btnRejectAllRevisions": // 拒绝修订
            return true;
        case "btnSaveAsFile":
            return true;
        case "btnInsertPic": //插入图片
            return true;
        case "btnOpenScan": //打开扫描仪
            return true;
        case "btnPageSetup": //打开页面设置
            return true;
        case "btnInsertDate": //插入日期
            return true;
        case "btnQRCode": //二维码
            return true;
        case "btnPrintDOC": // 打印设置
            return true;
        case "btnInsertBookmark":
            return true;
        case "btnImportTemplate":
            return true;
        case "btnSendMessage":
            return true;
        case "btnDebug":
            return true;
        default:
            ;
    }
    return true;
}


