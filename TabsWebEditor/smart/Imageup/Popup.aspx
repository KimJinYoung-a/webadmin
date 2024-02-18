<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Popup.aspx.cs" Inherits="smart_Imageup_Popup" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>이미지 올리기</title>
    <link rel="stylesheet" href="styles.css" />
    <script type="text/javascript" src="../../fckeditor/prototype.js"></script>
    <script type="text/javascript" src="../../fckeditor/imageup.js"></script>  
    <script type="text/javascript"> 
        function BeforeSubmit() {
            if (IsImage()) {
                return true;
            } 
            alert ('이미지 형식만 업로드가 가능합니다.'); 
            return false;
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
<body>
    <form id="form1" runat="server">
        <h1 class="head">
            이미지 올리기
        </h1> 
        <div class="body">
            <div class="content"> 
                <asp:FileUpload ID="uploadFile" runat="server" Width="350px" /><br /><br />
                <asp:Button ID="uploadbutton" runat="server" Text="확인" OnClick="uploadbutton_Click" OnClientClick="return BeforeSubmit()" /> 
            </div>
        </div> 
    </form>
</body>
</html>
