document.write("<script language='javascript' src='js/yozoCommon.js'></script>");

// var socket;
// if (!window.WebSocket) {
//     window.WebSocket = window.MozWebSocket;
// }
// var result;
// if (window.WebSocket) {
//     socket = new WebSocket("ws://localhost:5937/ws");
//     socket.onmessage = function(event) {
//         result = event.data;
//         console.log("result=="+result);
//         dispatcher(JSON.parse(result));
//     };
//     socket.onopen = function(event) {
//        console.log("链接开启");
//     };
//     socket.onclose = function(event) {
//        console.log("连接被关闭");
//     };
// } else {
//     alert("你的浏览器不支持 WebSocket！");
// }

// function send(message) {
//     if (!window.WebSocket) {
//         return;
//     }
//     if (socket.readyState == WebSocket.OPEN) {
//         socket.send(message);
//     } else {
//         alert("连接没有开启.");
//     }
// }

/**
 * web页面调用Yozo的方法入口
 *  * info参数结构
 * info:[
 *      {
 *       '方法名':'方法参数',需要执行的方法
 *     },
 *     ...
 *   ]
 * @param {*} info
 */
 function Dispatcher(info) {
     var funcs = (JSON.parse(info)).funcs;

     console.log("funcs=="+funcs);
     
     //执行web页面传递的方法
     setTimeout(function(){
         for (var index = 0; index < funcs.length; index++) {
             
             var func = funcs[index];
             for (var key in func) {
                 if (key === "OpenDoc") { // OpenDoc 属于普通的打开文档的操作方式，文档落地操作
                     OpenDoc(func[key]); //进入打开文档处理函数
                //  } else if (key === "OnlineEditDoc") { //在线方式打开文档，属于文档不落地的方式打开
                //      OnlineEditDoc(func[key]);
                //  } else if (key === "NewDoc") {
                //      OpenDoc(func[key]);
                //  } else if (key === "UseTemplate") {
                //      OpenDoc(func[key]);
                //  } else if (key === "InsertRedHead") {
                //      InsertRedHead(func[key]);
                //  } else if (key === "taskPaneBookMark"){
                //      taskPaneBookMark(func[key])
                //  } else if (key === "ExitWPS") {
                //      ExitWPS(func[key])
                //  } else if (key === "GetDocStatus") {
                //      return GetDocStatus(func[key])
                //  } else if (key === "NewOfficialDocument"){
                //      return OpenDoc(func[key])
                 }
             }
         }
     },100)
     return JSON.stringify({message:"ok", app:yozo.WApplication().getName()})
 }

 /**
 * 打开服务器上的文件
 * @param {*} fileUrl 文件url路径
 */
function OpenDoc(params) {
    var l_strFileUrl = params.fileName; //来自OA网页端的OA文件下载路径
    yozo.WApplication().openDocumentRemote(l_strFileUrl,false);
}

// document.write("<script language='javascript' src='js/common/common.js'></script>");
// document.write("<script language='javascript' src='js/common/func_docProcess.js'></script>");
// function OnAction(controlId,plugin){
//   switch(controlId){
//     case 'btnShowDialog':
//       yozo.showDialog("www.baidu.com","aaa",400,500,true);
//       break
//       case 'btnShowTaskPane':
//         yozo.showTaskPane(plugin,"oa面板");
       
//         break

//   }
// }
//自定义菜单按钮的点击执行事件
// function OnAction(controlId,plugin) {
//   console.log("controlId=="+controlId);
//   switch (controlId) {
//       case "btnShowTaskPane":
//         console.log("aaaaaaaaaaaaaa");
//          yozo.showTaskPane(plugin,"oa面板");
//           break;
//   }
 
// }
//ribbon.xml初始化调用
 function OnYozoLoad(){
    console.log("Onload")//设置内置button状态
 }

//自定义菜单按钮的点击执行事件
function OnAction(controlId,plugin) {
 // console.log("controlId=="+controlId);
  switch (controlId) {
      case "btnShowTaskPane": //打开WPS云文档入口
      yozo.showTaskPane(plugin,"http://123.60.5.149/plugin-ai-assistant","智能推荐");
          
          //yozo.showDialog("http://localhost:9222","sss",500,600,true);
          break;
      case "btnShowDebugPane": //打开网页调试窗口
        yozo.showDebugWindow();
        
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
      yozo.showDialog(plugin,"redhead.html", "套红", 700, 440, false); //套红头功能
          break;
      case "btnUploadOABackup": //文件备份
          OnUploadOABackupClicked();
          break;
      case "btnInsertSeal": //插入印章
      yozo.showDialog(plugin,"selectSeal.html", "印章", 700, 440, false);
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
              yozo.setButtonEnable("btnOpenRevision",0);
              yozo.setButtonEnable("btnCloseRevision",1);
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
             yozo.setButtonEnable("btnOpenRevision",1);
              yozo.setButtonEnable("btnCloseRevision",0);
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
          OnInsertQRCode(plugin); //
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
        //   OnImportTemplate();
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
 * 打开插入书签页面
 */
 function OnInsertBookmarkToDoc(plugin) {
    if (!yozo.WApplication().getActiveDocument()) {
        return;
    }
    var url = window.location.href;
    console.log(url);
    
    yozo.showDialog(plugin,"selectBookmark.html", "自定义书签", 700, 440, false);
}

/**
 *  作用：插入二维码图片
 */
 function OnInsertQRCode(plugin) {
     //暂不支持插入base64图片
    yozo.showDialog(plugin,"QRCode.html", "插入二维码", 400, 400,false);
}
 
function GetImage(controlId,plugin){
  //console.log("js--controlId==="+controlId);
    switch (controlId) {
      case "btnShowTaskPane":
          return "./icon/yozo/右侧面板.png"; //打开右侧面板
      case "btnShowDebugPane": //打开网页调试窗口
          return "./icon/yozo/debug.png"
      case "btnOpenLocalWPSYUN": //导入文件
           return "./icon/yozo/导入文件.png"
      case "btnSaveToServer": //保存到OA后台服务端
          return "./icon/yozo/保存到OA.png";
      case "btnSaveAsFile": //另存为本地文件
          return "./icon/yozo/保存本地.png";
      case "btnChangeToPDF": //输出为PDF格式
          return "./icon/yozo/转PDF上传.png";
      case "btnChangeToUOT": //
          return "./icon/yozo/转UOT上传.png";
      case "btnChangeToOFD": //转OFD上传
          return "./icon/yozo/转OFD上传.png"; //
      case "btnInsertRedHeader": //套红头
          return "./icon/yozo/套红头.png";
      case "btnInsertSeal": //印章
          return "./icon/yozo/印章.png";
      case "btnClearRevDoc": //清稿
          return "./icon/yozo/清稿.png"
      case "btnUploadOABackup": //备份正文
          return "./icon/yozo/备份正文.png";
      case "btnOpenRevision": //打开 修订
          return "./icon/yozo/打开修订.png";
      case "btnShowRevision": //显示修订
      case "btnCloseRevision": //关闭修订
          return "./icon/yozo/关闭修订.png";
      case "btnAcceptAllRevisions": // 接受修订
          return "./icon/yozo/接受修订.png";
      case "btnRejectAllRevisions": // 拒绝修订
          return "./icon/yozo/拒绝修订.png";
      case "btnSaveAsFile":
          return "";
      case "btnInsertPic": //插入图片
          return "./icon/yozo/插入图片.png";
      case "btnOpenScan": //打开扫描仪
          return "./icon/yozo/打开扫描仪.png"; //
      case "btnPageSetup": //打开页面设置
          return "./icon/yozo/页面设置.png";
      case "btnInsertDate": //插入日期
          return "./icon/yozo/插入日期.png";
      case "btnQRCode": //二维码
          return "./icon/yozo/二维码.png";
      case "btnPrintDOC": // 打印设置
          return "./icon/yozo/打印设置.png";
      case "btnInsertBookmark":
          return "./icon/yozo/导入书签.png";
      case "btnImportTemplate":
          return "./icon/yozo/导入模板.png";
      case "btnSendMessage":
          return "./icon/yozo/3.svg"
      case "btnDebug":
          return "./icon/yozo/导入模板.png"
      default:
          ;
  }
  return "./icon/yozo/c_default.png";
}

