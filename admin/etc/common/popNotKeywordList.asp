<%@ language=vbscript %>
<% option explicit %>
<!DOCTYPE html>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/etc/common/commonCls.asp"-->
<!-- #include virtual="/admin/etc/common/inc_tabkeyword.asp"-->
<%
Dim arrRows, oCommon, i
SET oCommon = new CCommon
	arrRows = oCommon.getOutmallNotKeyWordsList
SET oCommon = nothing
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript">
function goSubmit(v){
	var text;
	text = $("#keywords"+v).val();
	$("#vidx").val(v);
	$("#vkeywords").val(text);
	if (confirm("저장 하시겠습니까?")){
		document.frm.target = "xLink";
		document.frm.mode.value = "nREG";
		document.frm.action = "/admin/etc/common/procKeywords.asp"
		document.frm.submit();
	}
}
</script>
<div style="width:100%;">
	<br />
	<form name="frm" method="post" onSubmit="return false;" action="" style="margin:0px;">
	<input type="hidden" id="vidx" name="vidx" value="" />
	<input type="hidden" id="vkeywords" name="vkeywords" value="" />
	<input type="hidden" name="mode" value="" />
	<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<th width="70">제휴몰</td>
		<th>제외 키워드</td>
	</tr>
<%
If IsArray(arrRows) Then
	For i = 0 To Ubound(arrRows, 2)
%>
	<tr align="center" bgcolor="#FFFFFF">
		<td><%= arrRows(1,i) %></td>
		<td align="left">
			<textarea class="textarea" id="keywords<%= arrRows(0,i) %>" name="keywords<%= arrRows(0,i) %>" cols="100" rows="5"><%= arrRows(2,i) %></textarea>
			<input type="button" class="button" value="저장" onclick="goSubmit('<%= arrRows(0,i) %>');">
		</td>
	</tr>
<%
	Next
End If
%>
	</table>
	</form>
</div>
<iframe name="xLink" id="xLink" frameborder="0" width="100%" height="<%= CHKIIF(request("auto") <> "Y",300,100) %>"></iframe>
<% SET oCommon = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->