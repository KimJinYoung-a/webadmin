<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :   ���
' History : 2011.03.16 ������  ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" --> 
<!-- #include virtual="/lib/util/htmllib.asp"--> 
<%
	Dim arrFile,arrFilePath, i, arrName,sPosition 
	arrFile = ReplaceRequestSpecialChar(request("arrFN")) 
	arrFilePath= ReplaceRequestSpecialChar(request("arrFP")) 
	sPosition	= requestCheckVar(Request("sP"),10) 
	arrFile = split(arrFile,",")
	arrFilePath = split(arrFilePath,",")
%>
<div id="dAddFile">
	<%For i = 0 to UBound(arrFile)
		arrName = split(arrFile(i),".") 
	%>
	<div id="dF<%=arrName(0)%>"><%=arrFile(i)%>&nbsp;<input type="button" value="x" class="button" onclick=jsFileDel("<%=arrName(0)%>")></a>
	<input type="hidden" name="sFileP" value="<%=arrFilePath(i)%>"></div>
	<%Next%>
</div>
<script type="text/javascript" src="/js/jquery-1.7.2.min.js"></script>
<script language="javascript">
<!--
$(document).ready(function(){ 
	 var sValue = $("#dAddFile").html();    
	 $(opener.document).find("#dFile").append(sValue);   
	 self.close();
});
//-->
</script>
 
 