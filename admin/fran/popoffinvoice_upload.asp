<%@ language=vbscript %>
<% option explicit %>

<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%

dim masteridx, ino

masteridx = requestCheckVar(request("idx"), 32)
ino= requestCheckVar(request("ino"), 10)
%>
<script language="javascript">
<!--
	function jsUpload(){
		if(!document.frmImg.exportdeclarefile.value) {
			alert("수출허가증을 선택하세요.");
			return false;
		}

		document.frmImg.submit();
	}
//-->
</script>
<div style="padding: 0 5 5 5"> <img src="/images/icon_arrow_link.gif" align="absmiddle"> 수출허가증 첨부</div>
<table width="380" border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmImg" method="post" action="<%= uploadImgUrl %>/linkweb/offinvoice/offinvoice_upload.asp" enctype="MULTIPART/FORM-DATA">
<input type="hidden" name="masteridx" value="<%= masteridx %>">
<input type="hidden" name="ino" value="<%= ino %>">
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>" align="center" valign="top" style="padding:6 0 0 2">파일명</td>
		<td bgcolor="#FFFFFF">
			<table cellpadding="0" cellspacing="0" border="0" id="div1">
			<tr>
				<td><input type="file" name="exportdeclarefile"></td>
			</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td colspan="2" bgcolor="#FFFFFF" align="right">
			<!--<input type="image" src="/images/icon_confirm.gif">//-->
			<img src="/images/icon_confirm.gif" style="cursor:pointer" onclick="jsUpload();">
			<a href="javascript:window.close();"><img src="/images/icon_cancel.gif" border="0"></a>
		</td>
	</tr>
</form>
</table>
<!-- #include virtual="/lib/db/dbclose.asp" -->