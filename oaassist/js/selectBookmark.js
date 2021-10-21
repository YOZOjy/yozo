function addBookMarkjy() {  // 加载下拉框数据
    alert($("#jy").val());
    // var app = yozo.WApplication();
    // var doc = app.getActiveDocument();
    // var range = app.getSelection().getRange();
    // doc.getBookmarks().add($("#jy").val(),range);

    var app=yozo.WApplication();
    doc =  app.getActiveDocument();
    var bmks = doc.getBookmarks();
    bmks.add($("#jy").val(),app.getSelection().getRange());
  
}


function delBookMarkjy() {  // 加载下拉框数据
    alert($("#jy").val());
    var doc = yozo.WApplication().getActiveDocument();
    //var bookmarkData = GetDocParamsValue(doc, "bookmarkData");
    doc.getBookmarks().remove($("#jy").val());
}

function cancel()
{
    window.close();  
}