<%@ Language=VBScript %>
<%
	Option Explicit
	Response.Expires = -1440
	Response.CacheControl = "no-cache" 
	Response.AddHeader "Pragma", "no-cache" 

%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
dim sFileName, sFileLocation, sFilePath, sFileSize

sFileName = requestCheckVar(request("sfn"),100)
sFilePath = requestCheckVar(request("sfp"),100)
sFileLocation = requestCheckVar(request("sfl") ,50)
sFileSize= requestCheckVar(request("sfs") ,10)
if sFileName <> ""    then
%>
<script language="javascript">	
	var sFileName ="<%=sFileName%>";
	var sFilePath = "<%=sFilePath%>";
	var sFileLocation ="<%=sFileLocation%>";
	var sFileSize = "<%=sFileSize%>";
 	opener.jsSetFile(sFileName,sFilePath,sFileLocation,sFileSize);
	self.close();
</script>	
<%
else
%>
<script language="javascript">
	alert("���������ۿ� ������ �߻��Ͽ����ϴ�. �ٽ� �õ��� �ֽʽÿ�");
	self.close();
</script>
<%	
end if
%>