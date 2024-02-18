<%@ CodePage=65001 Language="VBScript"%>
<!--#include file="../ASP/util.asp"-->

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <title>파일 업로드</title>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <script type="text/javascript" src="../../fckeditor/prototype.js"></script>
    <script type="text/javascript" src="../../fckeditor/imageup.js"></script> 
</head>

<body>
    <%
        On Error Resume Next
        Response.CharSet = "utf-8" 
        
        Dim Upload, UploadForm, TempPath, FileName

        Set Upload = Server.CreateObject("TABS.Upload")
        TempPath = Server.MapPath("../Temp")
        Upload.CodePage = 65001 
        Upload.Start TempPath
         
        Set UploadForm = Upload.Form("uploadFile")
        FileName = GetGuid() &"." &UploadForm.FileType
        UploadForm.SaveAs TempPath &"\" &FileName, True
    %>
    <script type="text/javascript">
        onCompleteUpload('<%=FileName %>', '<%=UploadForm.FileName %>', '<%=UploadForm.FileSize %>');
    </script>
    <%
        Set Uplaod = Nothing
    %> 
</body>
</html>




