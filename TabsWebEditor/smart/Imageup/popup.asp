<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<%
'// 업체어드민 권한 여부로 인하여 인크루드 안하구 박아넣음
IF application("Svr_Info")="Dev" THEN
	uploadImgUrl    = "http://testupload.10x10.co.kr"
else
	uploadImgUrl    = "http://upload.10x10.co.kr"          '' upload.10x10.co.kr 통해서 Nas Server로 업로드
end if		
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>이미지 올리기</title>

<link rel="stylesheet" href="styles.css" />
<script type="text/javascript" src="../../fckeditor/prototype.js"></script>
<script type="text/javascript">
    function Submit() 
    {
        var form = $('uploadform');
        if (IsImage()) {
            if ($F('uploadFile') != "") { 
                form.submit();
            }
        } else {
            alert ('이미지 형식만 업로드가 가능합니다.');
        }  
    }
   
    function IsImage() {
        // 선택된 파일이 이미지인지의 여부를 검사한다.
        var path = $F('uploadFile');
        var temp = path.split('\\');
        var filename = temp[temp.length - 1];
       
        var exts = new Array();
        var isimg = false; 
        exts.push('.gif'); exts.push('.jpg'); exts.push('.png'); exts.push('.jpeg');
        exts.each(function(item) {
            var fname = filename.toLowerCase();
            if (fname.search(item) > -1) {
                isimg = true; 
            } 
        }); 
        return isimg; 
      } 
</script>
</head>

<body scroll="no" style="overflow: hidden;">
    <form name="uploadform" id="uploadform" action="<%=uploadImgUrl%>/linkweb/TabsWebEditor/editorUpload.asp" enctype="multipart/form-data" method="POST">
        <h1 class="head">
            이미지 올리기
        </h1> 
        <div class="body">
            <div class="content"> 
                <input type="file" id="uploadFile" name="uploadFile" style="width: 100%;" /><br /><br />
                <input type="button" id="uploadbutton" name="uploadbutton" value="확인" onclick="Submit()" /> 
            </div>
        </div>
    </form>
</body>
</html>