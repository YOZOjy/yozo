/**
 * 在这个js中，集中处理来自OA的传入的参数
 * 
 */

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
                // } else if (key === "OnlineEditDoc") { //在线方式打开文档，属于文档不落地的方式打开
                //     OnlineEditDoc(func[key]);
                } else if (key === "NewDoc") {
                    OpenDoc(func[key]);
                } 
                else if (key === "setMultiBookmarkValue") {
                    setMultiBookmarkValue(func[key]);
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
    console.log("========"+params);
    console.log("========-----------"+JSON.stringify(params));
    var l_strFileUrl = params.fileName; //来自OA网页端的OA文件下载路径
    var userName = params.userName;
    yozo.WApplication().setUserName(userName);
    if(l_strFileUrl != null){
        //打开远程文档
        yozo.WApplication().openDocumentRemote(l_strFileUrl,false);
    }else{
        //新建文档
        yozo.WApplication().createDocument("uot");
    }
   
    //将docId设置到文档中
    var doc = yozo.WApplication().getActiveDocument();
    //根据docId将json字符串存到map中
    yozo.getPluginStorage().setItem(doc.getID(),JSON.stringify(params));
}

//批量设置书签
function setMultiBookmarkValue(params){
    for (var key in params) {
        console.log(key+":",",value:"+params[key])
        setBookmarkValue(key,params[key])
    }
   return "success";
}