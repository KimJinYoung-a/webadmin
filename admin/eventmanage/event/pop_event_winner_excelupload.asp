<%@ language=vbscript %>
<% option explicit %>

<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
Dim vEventId
	vEventId		= requestCheckVar(Request("eventid"),10)
 	
 	If vEventId = "" Then
		Response.Write "<script>alert('�̺�Ʈ�ڵ尡 �����ϴ�.'); window.close();</script>"
		dbget.close()
		Response.End
	End If
	
Dim  iMaxLength	
	IF iMaxLength = "" THEN iMaxLength = 10
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript">
function jsSumbit(){
	var frm = document.frmFile;

	arrFileExt = new Array();
	arrFileExt[arrFileExt.length]  = "XLS";
	arrFileExt[arrFileExt.length]  = "XLSX";
	
	//���� Ȯ��
	if( frm.sFile.value =="") {
		alert("������ �Է��Ͻʽÿ�.");
		frm.sFile.focus();
		return;
	}

	//������ȿ�� üũ
	if (!fnChkFile(frm.sFile.value, <%=iMaxLength%>, arrFileExt)){
		alert("������ <%=iMaxLength%>MB������ XLS,XLSX ���ϸ� ���ε� �����մϴ�.");
		return;
	}
	
	frm.submit();
	
	$("#preProc").hide();
	$("#doingProc").show();
}

function fnChkFile(sFile, sMaxSize, arrExt){
	//���� ���ε� ����Ȯ��
	if (!sFile){
		return true;
	}
	
	var blnResult = false;
	
	//���� �뷮 Ȯ��
	var maxsize = sMaxSize * 1024 * 1024;
	
	//���� Ȯ���� Ȯ��
	var pPoint = sFile.lastIndexOf('.');
	var fPoint = sFile.substring(pPoint+1,sFile.length);
	var fExet = fPoint.toLowerCase();
	
	for (var i = 0; i < arrExt.length; i++)
	{
		if (arrExt[i].toLowerCase() == fExet) 
		{ 
			blnResult =  true;
		}
	}
	
	return blnResult;
}
</script>

<form name="frmFile" method="post" action="<%=uploadImgUrl%>/linkweb/event_admin/event_winner_excel_upload.asp"  enctype="MULTIPART/FORM-DATA">
<input type="hidden" name="iML" value="<%=iMaxLength%>">
<input type="hidden" name="sRID" value="<%=session("ssBctId")%>">
<input type="hidden" name="eventid" value="<%=vEventId%>">
<table width="100%" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr>
	<td width="70" align="center" bgcolor="<%= adminColor("tabletop") %>">���ϼ���</td>
	<td bgcolor="#FFFFFF">
		<input type="file" name="sFile" class="button">
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>"></td>
	<td bgcolor="#FFFFFF">
		<span id="preProc"><input type="button" class="button" value=" ��  �� " onClick="jsSumbit();"></span>
		<span id="doingProc" style="display:none;"><font color="red" size="3"><strong>* ���� �� �Դϴ�.<br>���ݸ� �״�� ��ٷ��ּ���!<br>Alertâ�� ��ϴ�!</strong></font></span>
	</td>
</tr>
</table>
</form>

<!-- #include virtual="/lib/db/dbclose.asp" -->