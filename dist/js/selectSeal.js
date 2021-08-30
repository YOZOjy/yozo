function addQz(){
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
}

function cancel(){
    
}