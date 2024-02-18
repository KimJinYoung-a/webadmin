<%@ page contentType="text/html; charset=UTF-8" pageEncoding="UTF-8" %>
<%@ page import="com.oreilly.servlet.MultipartRequest, com.oreilly.servlet.multipart.DefaultFileRenamePolicy, java.util.*,java.io.*,util.RandomGUID" %> 

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
	request.setCharacterEncoding("UTF8");
	String tempPath = application.getRealPath("/") + "smart\\Temp";
	int sizeLimit = 10 * 1024 * 1024;

	MultipartRequest multi = new MultipartRequest(request, tempPath, sizeLimit, "utf-8", new DefaultFileRenamePolicy());
	Enumeration formNames = multi.getFileNames();
	String formName = (String)formNames.nextElement();
	String fileName = multi.getFilesystemName(formName);

	fileName = new String(fileName.getBytes("UTF-8"), "utf-8");
	File file = new File(tempPath + "\\" + fileName);
	Long fileSize = file.length();
	String extension = fileName.substring(fileName.lastIndexOf("."), fileName.length());
	RandomGUID myGUID = new RandomGUID();
	String fileName2 = myGUID.toString() + extension;
	file.renameTo(new File(tempPath + "\\" + fileName2));
%>
<script type="text/javascript">
	onCompleteUpload('<%=fileName2 %>', '<%=fileName %>', '<%=fileSize %>');
</script>
</body>
</html>