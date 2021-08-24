function addBookMarkjy() {  // 加载下拉框数据
    alert($("#jy").val());
    var doc = yozo.WApplication().getActiveDocument();
    //var bookmarkData = GetDocParamsValue(doc, "bookmarkData");
    doc.getBookmarks().add($("#jy").val());

}



function delBookMarkjy() {  // 加载下拉框数据
    alert($("#jy").val());
    var doc = yozo.WApplication().getActiveDocument();
    //var bookmarkData = GetDocParamsValue(doc, "bookmarkData");
    yozo.bookmarks().remove($("#jy").val());
}


function cancel() {  // 加载下拉框数据
    window.close();
}