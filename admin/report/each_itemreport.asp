<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/report/reportcls.asp"-->
<%

dim yyyy1,mm1,dd1,yyyy2,mm2,dd2

if (yyyy1="") then yyyy1 = Cstr(Year(now()))
if (mm1="") then mm1 = Cstr(Month(now()))
if (dd1="") then dd1 = Cstr(day(now()))
if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
if (dd2="") then dd2 = Cstr(day(now()))
%>

<table width="100%" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frm" method="post" action="each_itemreport_result.asp">
	  <input type="hidden" name="menupos" value="<%= request("menupos") %>">
	<tr>
		<td class="a" >
		&nbsp;�����ۼ��� : <input type="text" name="itemidlist" size="70">&nbsp;(��ǰ��ȣ ���̿� ,�� �־��ּ���)<br>
		&nbsp;�˻��Ⱓ :&nbsp;&nbsp;&nbsp;
		<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
		&nbsp;&nbsp;&nbsp;&nbsp;
        �ɼ�:
            <input type="radio" name="settle2" value="m" checked>����
            <input type="radio" name="settle2" value="w">�ֺ�
            <input type="radio" name="settle2" value="d">�Ϻ�
		</td>
		<td class="a" align="right">
			<input type="image" src="/images/search2.gif" width="74" height="22" border="0"></a>
		</td>
	</tr>
	</form>
</table>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
