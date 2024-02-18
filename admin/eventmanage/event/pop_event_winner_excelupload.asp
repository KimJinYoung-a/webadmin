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
		Response.Write "<script>alert('이벤트코드가 없습니다.'); window.close();</script>"
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
	
	//파일 확인
	if( frm.sFile.value =="") {
		alert("파일을 입력하십시오.");
		frm.sFile.focus();
		return;
	}

	//파일유효성 체크
	if (!fnChkFile(frm.sFile.value, <%=iMaxLength%>, arrFileExt)){
		alert("파일은 <%=iMaxLength%>MB이하의 XLS,XLSX 파일만 업로드 가능합니다.");
		return;
	}
	
	frm.submit();
	
	$("#preProc").hide();
	$("#doingProc").show();
}

function fnChkFile(sFile, sMaxSize, arrExt){
	//파일 업로드 유무확인
	if (!sFile){
		return true;
	}
	
	var blnResult = false;
	
	//파일 용량 확인
	var maxsize = sMaxSize * 1024 * 1024;
	
	//파일 확장자 확인
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
	<td width="70" align="center" bgcolor="<%= adminColor("tabletop") %>">파일선택</td>
	<td bgcolor="#FFFFFF">
		<input type="file" name="sFile" class="button">
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>"></td>
	<td bgcolor="#FFFFFF">
		<span id="preProc"><input type="button" class="button" value=" 등  록 " onClick="jsSumbit();"></span>
		<span id="doingProc" style="display:none;"><font color="red" size="3"><strong>* 저장 중 입니다.<br>조금만 그대로 기다려주세요!<br>Alert창이 뜹니다!</strong></font></span>
	</td>
</tr>
</table>
</form>

<!-- #include virtual="/lib/db/dbclose.asp" -->