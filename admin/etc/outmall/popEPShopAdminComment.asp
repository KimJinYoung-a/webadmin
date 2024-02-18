<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbCTopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/etc/outmallConfirmCls.asp"-->
<!-- #include virtual="/admin/etc/incOutMallCommonFunction.asp"-->
<%
Dim mallid, sellyn, makerid
mallid	= request("mallid")
makerid	= request("makerid")
sellyn	= request("sellyn")
%>
<script language='javascript'>
function frmsubmit(){
	var frm = document.frm;
	if(frm.sugiadminText.value == ''){
		alert('사유를 입력하세요');
		frm.sugiadminText.focus();
		return;
	}
	frm.submit();
}
</script>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="POST" action="/admin/etc/outmall/confirm_process.asp" style="margin:0px;">
<input type="hidden" name="cmdparam" value="epShop">
<input type="hidden" name="sugimallid" value="<%=mallid%>">
<input type="hidden" name="sugisellyn" value="<%=sellyn%>">
<input type="hidden" name="sugimakerid" value="<%=makerid%>">
<input type="hidden" name="sugiadminid" value="<%=session("ssBctID")%>">
<tr bgcolor="#FFFFFF" height="30">
	<td>
		[ 브랜드ID : <%= makerid %> ]<br><br>
		업체에게 전달할 코멘트 입력 후 저장 버튼<br>
		<input type="text" class="text" name="sugiadminText" size="100" onkeydown="if(event.keyCode==13){alert('저장 버튼을 클릭해주세요.');event.returnValue=false;}">
	</td>
</tr>
<tr bgcolor="#FFFFFF" height="30" align="center">
	<td>
		<input type="button" class="button" value="저장" onclick="frmsubmit();">&nbsp;&nbsp;
		<input type="button" class="button" value="취소" onclick="self.close();">
	</td>
</tr>
</form>
</table>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbCTclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->