function onAction(controlId, plugin) {
    switch (controlId) {
        case 'btnShowTaskPane':
            // yozo.showTaskPane(plugin);
            var url = plugin.descriptor.localLocation + 'sign.html'
            yozo.showDialog(url, "签章", 600, 400, true)
            break
        case 'showSignDialog':
            var url = plugin.descriptor.localLocation + 'sign.html'
            yozo.showDialog(url, "签章", 600, 400, true)


            break

    }
}

function jy() {

    alert('aaaa');
}