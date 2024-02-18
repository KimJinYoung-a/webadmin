<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/event_function.asp"-->
<%
Dim mode, rev, vol
mode = request("mode")
rev = request("rev")
vol = request("vol")
%>
<script language="javascript">
<!--
	document.domain ="10x10.co.kr";
	function jsUpload(){
		if(!document.regfrm.packageFile.value){
			alert("찾아보기 버튼을 눌러 업로드할 이미지를 선택해 주세요.");
			return false;
		}
	}
//-->
</script>
<div style="padding: 0 5 5 5"> <img src="/images/icon_arrow_link.gif" align="absmiddle"> ZIP파일 업로드 처리</div>
<table width="440" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<form name="regfrm" method="post" action="<%= staticUploadUrl %>/linkweb/appmanage/package_upload.asp?userid=<%=session("ssBctId")%>&mode=<%=mode%>&rev=<%=rev%>&vol=<%=vol%>" enctype="MULTIPART/FORM-DATA" onSubmit="return jsUpload();">
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">ZIP파일명</td>
	<td bgcolor="#FFFFFF"><input type="file" name="packageFile" size = "35"></td>
</tr>
<tr>
	<td colspan="2" bgcolor="#FFFFFF" align="right">
		<input type="image" src="/images/icon_confirm.gif">
		<a href="javascript:window.close();"><img src="/images/icon_cancel.gif" border="0"></a>
	</td>
</tr>
</form>
</table>
<!-- #include virtual="/lib/db/dbclose.asp" -->