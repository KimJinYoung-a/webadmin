var onCompleteUpload = function(uid, displayName, size) {
    try {
        var data = '{ uid: "' + uid + '", displayName: "' + displayName + '", size: "' + size + '", isOldFile: "false" }';
        var wnd = window.opener.parent; 
        wnd.OnCompleteUpload(data);
    } catch(e) {
    } finally {
        self.close();
    } 
};

var g_fileListName = 'divFileList';
var g_newFilesName = 'newFiles';
var g_deletedFilesName = 'deletedFiles';

var g_fileCount = 0;
var g_newHash = new Hash();
var g_deletedHash = new Hash();

var fileListControl = {
    AddFile: function(element, fileInfo) {
        element = $(element);
        // Hash리스트에 첨부파일 정보를 기록한다. 
        var key = 'file_' + g_fileCount++;
        g_newHash.set(key, fileInfo); 
    
        // divFileList에 첨부파일 목록을 출력한다.        
        var div = document.createElement('div');
        var href = document.createElement('a'); 
        
        div.id = key;
        href.innerHTML = '<span style=""color: #cccccc;"">[삭제]</span>';
        href.style.cursor="hand"; 
        href.onclick = function() { 
            var parent = this.parentNode;
            if (parent.style.display != 'none') parent.style.display = 'none';

            // Hash리스트에 첨부파일 정보를 삭제한다.
            g_deletedHash.set(key, g_newHash.get(key));
            g_newHash.unset(key);
        }; 
        div.onmouseover = function() {
            this.style.backgroundColor = '#b2cbff'; 
        };
        div.onmouseout = function() {
            this.style.backgroundColor = '#ffffff'; 
        };
        div.innerHTML = fileInfo.displayName + ' (' + getFileSize(fileInfo.size) + ')';
        div.appendChild(href);
        element.appendChild(div);  
    },
    AddServerFile: function(element, uid, displayName, size) {
        element = $(element);       
        var fileInfo = ('{ "uid": "' + uid + '", "displayName": "' + displayName + '", "size": "' + size + '", "isOldFile": "true" }').evalJSON();
        element.AddFile(fileInfo); 
    }  
};

Element.addMethods($(g_fileListName), fileListControl);