<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/weekly_codi_qna_cls.asp" -->

<%
dim idx,i,masteridx,page
idx=request("idx")
masteridx=request("masteridx")
page=request("page")
dim wqna
set wqna = new WeeklyQna
wqna.getOneQna idx
%>
<script>
function Fnsubmit(){
	if (document.ansfrm.answer.value.length<1){
	
	alert('�亯������ �Է��� �ּ���.');
	
	return;
	}
	document.ansfrm.mode.value="add";
	document.ansfrm.submit();
	
}
function FnDel(){
	if (confirm('�����Ͻðڽ��ϱ�?.')) {
	
	document.ansfrm.mode.value="del";
	document.ansfrm.submit();
	}
}
	
	
</script>
<a href="/admin/board/weekly_codi_qna_list.asp?page=<%= page %>&masteridx=<%= masteridx %>"><font color="red">** ������� **</font></a>
<table width="560" border="0" cellpadding="5" cellspacing=1" class="a" bgcolor="000000">
	<tr bgcolor="#DDDDFF">
		<td width="80" align="center">UserId</td>
		<td width="200" align="center" bgcolor="#FFFFFF"><%= wqna.ouserid %></td>
		<td width="80" align="center">��� ��¥</td>
		<td width="200" align="center" bgcolor="#FFFFFF"><%= left(wqna.oRegdate,10) %></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td colspan="4" align="center"><a href="http://www.10x10.co.kr/guidebook/weekly_codi.asp?idx=<%= wqna.oMasteridx %>" target="_blank">��Ŭ�� �ڵ������ ����</a></td>
	</tr>
	<tr bgcolor="#DDDDFF">
		<td align="center" colspan="1">���� ����</td>
		<td colspan="3" bgcolor="#FFFFFF"><%= nl2br(db2html(wqna.oQuestion)) %></td>
	</tr>
	<form name="ansfrm" method="post" action="/admin/board/lib/doweekly_qna_answer.asp" onsubmit="return false;">
	<input type="hidden" name="page" value="<%= page %>">
	<input type="hidden" name="masteridx" value="<%= masteridx %>">
	<input type="hidden" name="idx" value="<%= wqna.oIdx %>">
	<input type="hidden" name="mode" value="">
	<tr bgcolor="#DDDDFF">
		<td align="center" colspan="1">�� ��</td>
		<td align="left" colspan="3" bgcolor="#FFFFFF"><textarea name="answer" cols="62" rows="10"><%= nl2br(wqna.oAnswer) %></textarea></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td colspan=2 align="center"><input type="button" value="Ȯ��" onclick="javascript:Fnsubmit();"></td>
		<td colspan=2 align="center"><input type="button" value="����" onclick="javascript:FnDel();"></td>
	</tr>
	</form>
	
	
</table>


<% set wqna=nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->