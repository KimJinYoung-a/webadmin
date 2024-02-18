<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stockclass/monthlystockcls.asp"-->
<%

dim page, research, i
dim yyyy1, mm1, tmpDate


page       		= requestCheckvar(request("page"),10)
research		= requestCheckvar(request("research"),10)
yyyy1       	= requestCheckvar(request("yyyy1"),10)
mm1         	= requestCheckvar(request("mm1"),10)

if (page="") then page = 1
if (yyyy1="") then
	tmpDate = Left(DateAdd("m", -1, Now()), 7)
	yyyy1 = Left(tmpDate, 4)
	mm1 = Right(tmpDate, 2)
end if


'// ============================================================================
dim ojaego
set ojaego = new CMonthlyStock

ojaego.FPageSize = 200
ojaego.FCurrPage = page
ojaego.FRectYYYYMM = yyyy1 + "-" + mm1

ojaego.GetMonthlyMoveDiffList

%>

<script language='javascript'>

function NextPage(page){
	document.frm.page.value = page;
	document.frm.submit();
}
</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="page" value="">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="1" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
			&nbsp;
			<font color="#CC3333">��/�� :</font> <% DrawYMBox yyyy1,mm1 %> �� �̵�����
		</td>

		<td rowspan="1" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	</form>
</table>
<!-- �˻� �� -->
<p>

* �ִ� 200������ ǥ�õ˴ϴ�.

<p>

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="25">
		<td width="60">�̵���</td>
		<td width="120">�귣��</td>
		<td width="30">����</td>
		<td width=70>��ǰ�ڵ�</td>
		<td width=40>�ɼ�</td>
		<td width=40>����</td>
		<td>���</td>
	</tr>
<% if ojaego.FResultCount >0 then %>
	<% for i=0 to ojaego.FResultcount-1 %>
	<tr bgcolor="#FFFFFF" height=25>
		<td align=center><%= ojaego.FItemList(i).Fyyyymm %></td>
		<td align=center><%= ojaego.FItemList(i).Flastmakerid %></td>
		<td align=center><%= ojaego.FItemList(i).Fitemgubun %></td>
		<td align="right"><%= ojaego.FItemList(i).Fitemid %></td>
		<td align=center><%= ojaego.FItemList(i).Fitemoption %></td>
		<td align="right">
			<%= FormatNumber(ojaego.FItemList(i).FtotItemNo, 0) %>
		</td>
		<td>
	    </td>
	</tr>
	<% next %>
<% else %>
	<tr bgcolor="#FFFFFF" height=50>
		<td align=center colspan="17">������ �����ϴ�.</td>
	</tr>
<% end if %>
</table>

<%
set ojaego = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
