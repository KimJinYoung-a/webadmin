<%@ language=vbscript %>
<% option explicit %>

<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/test/tempdata/classes/eventcntcls.asp"-->

<%
dim oeventuserlist , i

	set oeventuserlist = new Ceventuserlist
	oeventuserlist.Feventuserlist3()
%>

<script language="javascript">
function excel()
{
var popup = window.open('/test/tempdata/40245_excel.asp','excel','width=1024,height=768,scrollbars=yes,resizable=yes');
popup.focus();
}
</script>

<table width="100%" align="center" cellpadding="0" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr>
		<td colspan=13 align="center">카카오톡_다이어트대작전(40245) 이벤트 응모건 수 데이터</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td align="center">날짜</td>
		<td align="center">2월13일</td>
		<td align="center">2월14일</td>
		<td align="center">2월15일</td>
		<td align="center">2월16일</td>
		<td align="center">2월17일</td>
		<td align="center">2월18일</td>
		<td align="center">2월19일</td>
		<td align="center">2월20일</td>
		<td align="center">2월21일</td>
		<td align="center">2월22일</td>
		<td align="center">2월23일</td>
		<td align="center">2월24일</td>
    </tr>
	<tr align="center" bgcolor="#FFFFFF">
		<td align="center">응모건수</td>
		<% for i= 0 to oeventuserlist.FResultCount-1 %>
		<td align="center"><%= oeventuserlist.flist(i) %></td>
		<% next %>
	</tr>

</table>
<br>

<table>
	<tr>
		<td>
			※ 총 기간내 순수 참여자 수: <%= oeventuserlist.Ftotalcount %>
		</td>
	</tr>
	<tr>
		<td>
			<input type="button" name="excelbox" value="엑셀파일로저장" class="button" onclick="excel();">
		</td>
	</tr>
</table>


<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->