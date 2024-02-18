var g_multimediaListName = 'divMultimediaList';
var g_multimediaCount = 0;
var g_newMultimediaHash = new Hash();
var g_deletedMultimediaHash = new Hash();

var multimediaListControl = {
    AddMultimedia: function(element, multimediaInfo, thumbnail) {
        element = $(element);
        // Hash리스트에 멀티미디어 파일 정보를 기록한다.  
        var key = 'multi_' + g_multimediaCount++;
        g_newMultimediaHash.set(key, multimediaInfo);
       
        // divMultimediaList에 멀티미디어 파일 목록을 출력한다.
        if (thumbnail) { 
            var tbl = document.createElement('table');
            var tr1 = tbl.insertRow(0);
            var tr0 = tbl.insertRow(0);
            tbl.id = key;
            tbl.width = '85px'; 
            tbl.style.display = 'inline'; 
            var img = document.createElement('img');
		    var url = (multimediaInfo.isOldFile == 'false') ? getCurrentPageFullURL('/Temp/') : getCurrentPageFullURL('/Upload/');
            img.src = url + multimediaInfo.thumbPath;
            img.width = '85';
            img.height = '70';  
            var td0 = tr0.insertCell(0);
            td0.appendChild(img);
            var td1 = tr1.insertCell(0);
            var span = document.createElement('span');
            span.innerHTML = '[삭제]';
            span.style.cursor="hand";  
            span.onclick = function() {
                // Thumbnail 이미지 삭제 
                $(key).remove();
               
                // 에디터 내용중 해당 이미지 태그 삭제             
                var oEditor = FCKeditorAPI.GetInstance('editor');
                var editor = oEditor.GetHTML();
               
                var html = ''; 
                if (multimediaInfo.isPhoto) {
                    var re = /<\s*img\b[^>]*\bsrc\s*=\s*("[\w\x2d\x2e\x2f\x3a]+"|'[\w\x2d\x2e\x2f\x3a]+')[^>]*>/g;
                 
	                html = editor.replace(re, 
	                    function($0, $1)
	                    {  
	                        if ($1.indexOf(multimediaInfo.url) != -1)
	                            return '';
	                        else
	                            return $0;   
	                    }
	                );
	            } else {
	                var re = /<\s*embed\b[^>]*\bsrc\s*=[^>]*>/g;
	                var r = editor.match(re);
	                if (r != null) {
	                    for (var i = 0; i < r.length; i++) {
	                        var subre = /filepath=[\w\x2d\x2e\x2f\x3a]+/;
	                        var sub = r[i].match(subre);
	                        if (sub.indexOf(multimediaInfo.url) != -1)
	                            html = editor.replace(r, '');
	                    } 
	                }    
	            }
	            oEditor.SetHTML(html);
    	        
                // Hash리스트에 멀티미디어 파일 정보를 삭제한다.
                g_deletedMultimediaHash.set(key, g_newMultimediaHash.get(key));
                g_newMultimediaHash.unset(key);
            };
            td1.innerHTML = getFileSize(multimediaInfo.filesize);
            td1.appendChild(span); 
            element.appendChild(tbl);
        } 
    },
    AddServerMediaFile: function(element, url, oriname, filesize, thumbPath, isPhoto, thumbnail) {
        element = $(element);
        var multimediaInfo = ('{ "url": "' + url + '", "oriname": "' + oriname + '", "filesize": "' + filesize + '", "thumbPath": "' + thumbPath + '", "isOldFile": "true", "isPhoto": ' + isPhoto + ' }').evalJSON();
        element.AddMultimedia(multimediaInfo, thumbnail);
    } 
};

Element.addMethods($(g_multimediaListName), multimediaListControl);