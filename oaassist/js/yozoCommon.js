/**
 * 新建文档
 */
function newDoc(){
    var app=yozo.WApplication()
    console.log(app)
   var doc = app.getDocuments().add();
  var tt = "测测测试测测";
//app.getSelection().setText(tt);
     app.getSelection().getRange().setText(tt);

    //  alert( app.getSelection().getRange().getText());
    // var doc=app.getActiveDocument()
    //var doc = app.getWorkbooks().getActiveWorkbook().getDocuments().getActiveDocument();
    //console.log(doc);
    
    //console.log('document name is',doc.getName())
}
/**
 * 打开文档
 */
function openDoc(){
  var app=yozo.WApplication()
  app.openDocumentRemote("http://localhost:8080/files/test.doc",false);
}

/**
	*保存远程接口有有两种保存方式：一是普通二进制流保存，oa服务端写一个保存的servlet，获取插件上传的二进制流保存；
	*							   二是表单上传方式，office插件提交到后台是一个“file”对象，oa后台获取file对象进行处理。
  * saveType --- 文档上传类型，为空默认类型上传
	**/
	
	/******************************************************************************************/
	//表单上传
	function formUpload(saveType)
	{
		try{
      if(saveType == "")
          saveType = ".doc"
      var app=yozo.WApplication()
      var doc = app.getActiveDocument();
			var tmpSavePath = app.getTempPath() + doc.getName() +saveType;//文件本地路径，控制文件保存类型，如存ofd就改成ofd
			console.log(tmpSavePath);
			
			doc.saveAs(tmpSavePath);
      console.log(doc.getFullName())
      var l_uploadPath = "http://192.168.61.72:8080/springdemo/upload?dir=upload-dir&name=remote001"+saveType;
			var result = app.uploadFile(l_uploadPath, tmpSavePath);
      alert("上传成功");
			console.log("UploadFile返回值："+result);
		}catch(e){
			console.log(e);
		}
	}


/**
 *  打开本地文档，并插入文档
 */
 function OpenLocalFile() {
  var l_FileName = "";

  //msoFileDialogFilePicker = 3
  var l_FileDialog = yozo.WApplication().getFileDialog(3);
  if (l_FileDialog.show()) {
      l_FileName = l_FileDialog.getSelectedItems();
      if (l_FileName.getCount() > 0) {
        
          yozo.WApplication().getSelection().insertFile(l_FileName.item(1),null,null,null,null);
      }
  }
}

/**
 *  执行另存为本地文件操作
 */
 function OnBtnSaveAsLocalFile() {

  //初始化临时状态值
  //wps.PluginStorage.setItem(constStrEnum.OADocUserSave, false);
  //wps.PluginStorage.setItem(constStrEnum.IsInCurrOADocSaveAs, false);

  //检测是否有文档正在处理
  var l_doc = yozo.WApplication().getActiveDocument();
  if (!l_doc) {
      alert("WPS当前没有可操作文档！");
      return;
  }

  // 设置WPS文档对话框 2 FileDialogType:=msoFileDialogSaveAs
  var app = yozo.WApplication(); 
  var l_ksoFileDialog = app.getFileDialog(2);
  l_ksoFileDialog.setInitialFileName(l_doc.getName()); //文档名称
  if (l_ksoFileDialog.show() == -1) { // -1 代表确认按钮
      //alert("确认");
     // wps.PluginStorage.setItem(constStrEnum.OADocUserSave, true); //设置保存为临时状态，在Save事件中避免OA禁止另存为对话框
      //l_ksoFileDialog.Execute(); //会触发保存文档的监听函数

     // pSetNoneOADocFlag(l_doc);

     var savePath = l_ksoFileDialog.getSelectedItems().item(2);
     console.log("savePath=="+l_ksoFileDialog.getSelectedItems().item(2));
     // wps.ribbonUI.Invalidate(); //刷新Ribbon的状态
     l_doc.saveAs(savePath);
  };
}

/**
 * 套红头
*/
function OnInsertRedHeaderClick(){
  var strFile = "C:/Users/lenovo/Desktop/红头文件.docx";
  var bookmark = "Content";
  var app = yozo.WApplication();
  var doc = app.getActiveDocument();
  alert('111')
  var bookMarks = doc.getBookmarks();
  if (bookMarks.item("quanwen")) { // 当前文档存在"quanwen"书签时候表示已经套过红头
    alert("当前文档已套过红头，请勿重复操作!");
    return;
  }

  
  var activeDoc = app.getActiveDocument();
  var selection = app.getActiveWindow().getSelection();
  // 准备以非批注的模式插入红头文件(剪切/粘贴等操作会留有痕迹,故先关闭修订)
  activeDoc.setTrackRevisions(false);
  selection.wholeStory(); //选取全文
  bookMarks.add("quanwen", selection.getRange())
  selection.cut();
  selection.insertFile(strFile,null,null,null,null);
  if (bookMarks.exists(bookmark)) {
    var bookmark1 = bookMarks.item(bookmark);
    bookmark1.getRange().select(); //获取指定书签位置
    var s = activeDoc.getActiveWindow().getSelection();
    s.paste();
  } else {
    alert("套红头失败，您选择的红头模板没有对应书签：" + bookmark);
  }

  // 恢复修订模式(根据传入参数决定)
  activeDoc.setTrackRevisions(true);

}


/**
 * 作用：执行清稿按钮操作
 * 业务功能：清除所有修订痕迹和批注
 */
 function OnBtnClearRevDoc() {
  var doc = yozo.WApplication().getActiveDocument();
  if (!doc) {
      alert("尚未打开文档，请先打开文档再进行清稿操作！");
  }

  //执行清稿操作前，给用户提示
  // if (!wps.confirm("清稿操作将接受所有的修订内容，关闭修订显示。请确认执行清稿操作？")) {
  //     return;
  // }

  //接受所有修订
  if (doc.getRevisions().getCount() >= 1) {
      doc.acceptAllRevisions();
  }
  //去除所有批注
  if (doc.getComments().getCount() >= 1) {
      doc.deleteAllComments();
  }

  //删除所有ink墨迹对象
  //pDeleteAllInkObj(doc);

  doc.setTrackRevisions(false); //关闭修订模式
  //wps.ribbonUI.InvalidateControl("btnOpenRevision");

  return;
}

/**
 *  作用：把当前文档修订模式关闭
 */
 function OnCloseRevisions() {
  //获取当前文档对象
  var l_Doc = yozo.WApplication().getActiveDocument();
  OnRevisionsSwitch(l_Doc, false);
}


/**
*  作用：把当前文档修订模式打开
*/
function OnOpenRevisions() {
  //获取当前文档对象
  var l_Doc = yozo.WApplication().getActiveDocument();
  OnRevisionsSwitch(l_Doc, true);
}

function OnRevisionsSwitch(doc, openRevisions) {
  if (!doc) {
      return;
  }
  var l_activeWindow = doc.getActiveWindow();
  if (l_activeWindow) {
      var v = l_activeWindow.getView();
      if (v) {
          //WPS 显示使用“修订”功能对文档所作的修订和批注
          v.setShowRevisionsAndComments(openRevisions);
          //WPS 显示从文本到修订和批注气球之间的连接线
          v.setRevisionsBalloonShowConnectingLines(openRevisions);
      }
      if (openRevisions == true) {
          //去掉修改痕迹信息框中的接受修订和拒绝修订勾叉，使其不可用
         // wps.WpsApplication().CommandBars.ExecuteMso("KsoEx_RevisionCommentModify_Disable");
      }

      //RevisionsMode:
      //wdBalloonRevisions	0	在左边距或右边距的气球中显示修订。
      //wdInLineRevisions	1	在正文中显示修订，使用删除线表示删除，使用下划线表示插入。 
      //                      这是 Word 早期版本的默认设置。
      //wdMixedRevisions	2	不支持。
      doc.setTrackRevisions(openRevisions); // 开关修订
      l_activeWindow.getActivePane().getView().setRevisionsMode(2); //2为不支持气球显示。

  }
}

/**
 * 作用：接受所有修订内容
 * 
 */
 function OnAcceptAllRevisions() {
  //获取当前文档对象
  var l_Doc = yozo.WApplication().getActiveDocument();
  if (!l_Doc) {
      return;
  }
  if (l_Doc.getRevisions().getCount() >= 1) {
      // if (!wps.confirm("目前有" + l_Doc.Revisions.Count + "个修订信息，是否全部接受？")) {
      //     return;
      // }
      l_Doc.acceptAllRevisions();
  }
}


/**
* 作用：拒绝当前文档的所有修订内容
*/
function OnRejectAllRevisions() {
  var l_Doc = yozo.WApplication().getActiveDocument();
  if (!l_Doc) {
      return;
  }
  if (l_Doc.getRevisions().getCount() >= 1) {
       // if (!wps.confirm("目前有" + l_Doc.Revisions.Count + "个修订信息，是否全部拒绝？")) {
      //     return;
      // }
      l_Doc.rejectAllRevisions();
  }
}

/**
 *  作用：在文档的当前光标处插入从前端传递来的图片
 *  OA参数中 picPath 是需要插入的图片路径
 *  图片插入的默认版式是在浮于文档上方
 */
 function DoInsertPicToDoc() {
  console.log("DoInsertPicToDoc...");

  var l_doc; //文档对象
  l_doc = yozo.WApplication().getActiveDocument();
  if (!l_doc) {
      return;
  }

  //获取当前文档对象对应的OA参数
  // var l_picPath = GetDocParamsValue(l_doc, constStrEnum.picPath); // 获取OA参数传入的图片路径
  var l_picPath = "http://localhost:8080/pic/111.jpg";
  if (l_picPath == "") {
      // alert("未获取到系统传入的图片URL路径，不能正常插入图片");
      // return;
      //如果没有传，则默认写一个图片地址
      l_picPath="http://127.0.0.1:3888/file/OA模板公章.png"
  }

  //var l_picHeight = GetDocParamsValue(l_doc, constStrEnum.picHeight); //图片高
 // var l_picWidth = GetDocParamsValue(l_doc, constStrEnum.picWidth); //图片宽
  var l_picHeight;
  var l_picWidth;
  if (l_picHeight == "") { //设定图片高度
      l_picHeight = 39.117798; //13.8mm=39.117798磅
  }
  if (l_picWidth == "") { //设定图片宽度
      l_picWidth = 72; //49.7mm=140.880768磅
  }

  var l_shape = l_doc.getShapes().addPicture(l_picPath, false, true,null,null,null,null,null);
  l_shape.select();
  // l_shape.WrapFormat.Type = wps.Enum&&wps.Enum.wdWrapBehind||5; //图片的默认版式为浮于文字上方，可通过此设置图片环绕模式
}


/**
 * 作用：打开当前文档的页面设置对话框
 */
 function OnPageSetupClicked() {
  var app = yozo.WApplication();
  var doc = app.getActiveDocument();
  if (!doc) {
      return;
  }
  app.getDialogs().item(178).show(null);
}

/**
* 作用：打开当前文档的打印设置对话框
*/
function OnPrintDocBtnClicked() {
  var app = yozo.WApplication();
  var doc = app.getActiveDocument();
  if (!doc) {
      return;
  }
  app.getDialogs().item(88).show(null);
}

/**
 *  作用：打开扫描仪
 */
 function OnOpenScanBtnClicked() {
  var app = yozo.WApplication();
  var doc = app.getActiveDocument();
  if (!doc) {
      return;
  }
  //打开扫描仪
  try {
    app.getWordBasic().insertImagerScan(); //打开扫描仪,暂不支持
  } catch (err) {
      alert("打开扫描仪的过程遇到问题。");
  }
}

/**
 * 插入时间
 * params参数结构
 * params:{
 *   
 * }
 */
 function OnInsertDateClicked() {
  var l_Doc = yozo.WApplication().getActiveDocument();
  if (l_Doc) {
      //打开插入日期对话框
     yozo.WApplication().getDialogs().item(165).show();//有问题
  }
}

function addBookmark(){
    var app=yozo.WApplication()
    var doc = app.getActiveDocument();
    var bmks = doc.getBookmarks();
    var range = app.getSelection().getRange();
    bmks.add("test",range);
}

