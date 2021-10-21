(function (global,factory) {
    global.YozoInvoke = factory();
}(this,function (global) {

    // const URL_WS = "ws://192.168.61.176:5937/ws";
    // const URL_HTTP_POST = "http://192.168.61.176:5937/post";
    let URL_WS = "ws://127.0.0.1:5937/ws";
    let URL_HTTP_POST = "http://127.0.0.1:5937/post";

    const FUNCTION_TYPE_AJAX = 1;
    const FUNCTION_TYPE_WS = 2;

    let websocketReg = null;
    let websocketIvk = null;
    let isEIOAlive = false;

    function getHttpObj() {
        var httpobj = null;
        if (IEVersion() < 10) {
            try {
                httpobj = new XDomainRequest();
            } catch (e1) {
                httpobj = new createXHR();
            }
        } else {
            httpobj = new createXHR();
        }
        return httpobj;
    }
    //兼容IE低版本的创建xmlhttprequest对象的方法
    function createXHR() {
        if (typeof XMLHttpRequest != 'undefined') { //兼容高版本浏览器
            return new XMLHttpRequest();
        } else if (typeof ActiveXObject != 'undefined') { //IE6 采用 ActiveXObject， 兼容IE6
            var versions = [ //由于MSXML库有3个版本，因此都要考虑
                'MSXML2.XMLHttp.6.0',
                'MSXML2.XMLHttp.3.0',
                'MSXML2.XMLHttp'
            ];

            for (var i = 0; i < versions.length; i++) {
                try {
                    return new ActiveXObject(versions[i]);
                } catch (e) {
                    //跳过
                }
            }
        } else {
            throw new Error('您的浏览器不支持XHR对象');
        }
    }
    //获取IE版本
    function IEVersion() {
        var userAgent = navigator.userAgent; //取得浏览器的userAgent字符串
        var isIE = userAgent.indexOf("compatible") > -1 && userAgent.indexOf("MSIE") > -1; //判断是否IE<11浏览器
        var isEdge = userAgent.indexOf("Edge") > -1 && !isIE; //判断是否IE的Edge浏览器
        var isIE11 = userAgent.indexOf('Trident') > -1 && userAgent.indexOf("rv:11.0") > -1;
        if (isIE) {
            var reIE = new RegExp("MSIE (\\d+\\.\\d+);");
            reIE.test(userAgent);
            var fIEVersion = parseFloat(RegExp["$1"]);
            if (fIEVersion == 7) {
                return 7;
            } else if (fIEVersion == 8) {
                return 8;
            } else if (fIEVersion == 9) {
                return 9;
            } else if (fIEVersion == 10) {
                return 10;
            } else {
                return 6; //IE版本<=7
            }
        } else if (isEdge) {
            return 20; //edge
        } else if (isIE11) {
            return 11; //IE11
        } else {
            return 30; //不是ie浏览器
        }
    }

    //向中间服务发送ajax请求
    function sendMsgAjax(params,callback) {
        //判断传参是否为string
        if (typeof params != 'string'){
            params = JSON.stringify(params);
        }
        let xhr = getHttpObj();
        xhr.open('POST',URL_HTTP_POST,false);
        xhr.setRequestHeader("Content-type","application/x-www-form-urlencoded");
        console.log("AjaxSend => "+params);
        xhr.onreadystatechange = function () {
            if (xhr.readyState === 4 && xhr.status === 200) {
                if (callback != null && callback != undefined) {
                    callback(xhr.responseText);
                }
            }else {
                throw new Error(`请求中间服务失败`);
            }
        };
        xhr.send(params);
    }
    //向中间服务发送socket消息
    function sendMsgSocket(params,callback) {
        //判断传参是否为string
        if (typeof params != 'string'){
            params = JSON.stringify(params);
        }
        if (websocketReg != null){
            websocketReg.send(params);
            console.log("WSSend => "+params);
            websocketReg.onmessage = function (event) {
                callback(event.data);
            };
        }
    }

    //初始化websocket方法
    function initWebsocket(params,callback,functionType) {
        let websocket = new ReconnectingWebSocket(URL_WS,null,{debug: true ,reconnectInterval: 1000});
        websocket.onopen = function(){
            isEIOAlive = true;
            if (functionType === FUNCTION_TYPE_AJAX){
                sendMsgAjax(params,callback);
            }else if (functionType === FUNCTION_TYPE_WS){
                sendMsgSocket(params,callback);
            }else {
                throw new Error(`未知functionType => `+ functionType);
            }
            websocket.onopen = function () {
                //重置
            };
        };
        websocket.onclose = function () {
            isEIOAlive = false;
            console.log("与EIO失去连接");
            websocket.onopen = function () {
                //重置
            }
        };
        websocket.onerror = function(){
            isEIOAlive = false;
            // throw new Error(`与EIO连接失败，正在尝试重连`);
        };
        return websocket;
    }

    //根据不同情况 发送ws/ajax消息,目前RegWebNotify方法用socket.send，invoke用ajax
    function sendMultiMsg(params,callback,functionType) {
        let sendFunction;
        let websocket;
        //根据不同方法使用不同的socket对象和发送方法
        if (functionType === FUNCTION_TYPE_WS){
            sendFunction = sendMsgSocket;
            websocket = websocketReg;
        }else if (functionType === FUNCTION_TYPE_AJAX){
            sendFunction = sendMsgAjax;
            websocket = websocketIvk;
        }else {
            throw new Error(`未知functionType =>` + functionType);
        }
        if (websocket == null){
            //若无websocket对象则通过init方法初始化
            websocket = initWebsocket(params,callback,functionType);
            if (functionType === FUNCTION_TYPE_WS){
                websocketReg = websocket;
            }else {
                websocketIvk = websocket;
            }
        }else if (websocket !== null && isEIOAlive === false){
            console.log('5');
            websocket.onopen = function () {
                isEIOAlive = true;
                sendFunction(params,callback);
                websocket.onopen = function () {
                    //重置
                };
            }
        }else {
            sendFunction(params,callback);
        }
    }

    function invoke(params,callback) {
        sendMultiMsg(params,callback,FUNCTION_TYPE_AJAX);
    }

    function regWebNotify(params,callback) {
        sendMultiMsg(params,callback,FUNCTION_TYPE_WS);
    }

    return {
        invoke: invoke,
        regWebNotify: regWebNotify
    }
}));