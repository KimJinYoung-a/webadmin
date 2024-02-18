<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  ù�����ڰ��� ���
' History : 2018.11.07 �̻� ����
'####################################################
%>
<!-- #include virtual="/admin/incSessionSTAdmin.asp" -->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/maechul/managementSupport/maechulCls.asp" -->
<!-- #include virtual="/lib/classes/maechul/statistic/statisticCls_analisys.asp" -->
<%

dim i, j, k
dim vSYear, vSMonth, vSDay, vEYear, vEMonth, vEDay, searchenddate

vSYear		= NullFillWith(request("syear"),Year(DateAdd("d",-7,now())))
vSMonth		= NullFillWith(request("smonth"),Month(DateAdd("d",-7,now())))
vSDay		= NullFillWith(request("sday"),Day(DateAdd("d",-7,now())))
vEYear		= NullFillWith(request("eyear"),Year(now))
vEMonth		= NullFillWith(request("emonth"),Month(now))
vEDay		= NullFillWith(request("eday"),Day(now))

searchenddate = DateAdd("d",+1,DateSerial(vEYear, vEMonth,vEDay))

dim cStatistic
Set cStatistic = New cStaticTotalClass_list
cStatistic.FRectStartdate = vSYear & "-" & TwoNumber(vSMonth) & "-" & TwoNumber(vSDay)
cStatistic.FRectEndDate = Year(searchenddate) & "-" & TwoNumber(Month(searchenddate)) & "-" & TwoNumber(Day(searchenddate))
cStatistic.fStatistic_firstOrder()

%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript">
function searchSubmit()
{
	$("#btnSubmit").prop("disabled", true);
	frm.submit();
}
</script>
<!-- �˻� ���� -->
<form name="frm" method="get" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="70" bgcolor="<%= adminColor("gray") %>">�˻� ����</td>
	<td align="left">
		�������� :
		<% DrawDateBoxdynamic vSYear, "syear", vEYear, "eyear", vSMonth, "smonth", vEMonth, "emonth", vSDay, "sday", vEDay, "eday" %>
	</td>
	<td width="50" bgcolor="<%= adminColor("gray") %>"><input type="button" id="btnSubmit" class="button_s" value="�˻�" onClick="javascript:searchSubmit();"></td>
</tr>
</table>
</form>
<!-- �˻� �� -->

<p />

<table width="100%" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="<%= adminColor("tabletop") %>">
	<td align="center" rowspan="2" width="100">����</td>
	<td align="center" colspan="7">�����</td>
	<td align="center" colspan="7">�ֹ��Ǽ�</td>
    <td align="center" rowspan="2">���</td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>">
	<td align="center" width="100">1ȸ</td>
	<td align="center" width="100">2ȸ</td>
	<td align="center" width="100">3ȸ</td>
	<td align="center" width="100">4ȸ</td>
	<td align="center" width="100">5ȸ</td>
	<td align="center" width="100">6ȸ</td>
	<td align="center" width="100">7ȸ �̻�</td>
	<td align="center" width="100">1ȸ</td>
	<td align="center" width="100">2ȸ</td>
	<td align="center" width="100">3ȸ</td>
	<td align="center" width="100">4ȸ</td>
	<td align="center" width="100">5ȸ</td>
	<td align="center" width="100">6ȸ</td>
	<td align="center" width="100">7ȸ �̻�</td>
</tr>
<%
For i = 0 To cStatistic.FResultCount -1
%>
<tr bgcolor="#FFFFFF">
	<td align="center"><%= cStatistic.FList(i).Fyyyymmdd %></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Fsubtotalprice1) %></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Fsubtotalprice2) %></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Fsubtotalprice3) %></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Fsubtotalprice4) %></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Fsubtotalprice5) %></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Fsubtotalprice6) %></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Fsubtotalprice7) %></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Fcnt1) %></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Fcnt2) %></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Fcnt3) %></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Fcnt4) %></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Fcnt5) %></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Fcnt6) %></td>
	<td align="right" style="padding-right:5px;"><%= NullOrCurrFormat(cStatistic.FList(i).Fcnt7) %></td>
	<td></td>
</tr>
<%
Next
%>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->