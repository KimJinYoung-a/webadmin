var onCompleteUpload = function(uid, displayName, size) {
    try {
        var data = '{ uid: "' + uid + '", displayName: "' + displayName + '", size: "' + size + '" }';
        var wnd = window.opener.parent; 
        wnd.OnCompleteImageUpload(data);
    } catch(e) {
    } finally {
        self.close();
    } 
};