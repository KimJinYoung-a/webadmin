<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  ���� �԰���Ʈ
' History : 2007.01.01 �̻� ����
'			2017.01.06 �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/stock/newstoragecls.asp"-->
<%

dim yyyy1,yyyy2,mm1,mm2,dd1,dd2
dim fromDate, toDate
dim page, diffonly, research, i
research = request("research")
diffonly = request("diffonly")

yyyy1 = request("yyyy1")
yyyy2 = request("yyyy2")
mm1	  = request("mm1")
mm2	  = request("mm2")
dd1	  = request("dd1")
dd2	  = request("dd2")

if (yyyy1="") then yyyy1 = Cstr(Year(now()))
if (mm1="") then mm1 = Cstr(Month(now()))
if (dd1="") then dd1 = Cstr(day(now()))
''if (yyyy2="") then yyyy2 = Cstr(Year(now()))
''if (mm2="") then mm2 = Cstr(Month(now()))
''if (dd2="") then dd2 = Cstr(day(now()))

fromDate = CStr(DateSerial(yyyy1, mm1, dd1))
''toDate = CStr(DateSerial(yyyy2, mm2, dd2+1))

if research = "" then diffonly = "Y"

dim oipchul
set oipchul = new CIpChulStorage

oipchul.FRectYYYYMMDD = fromDate
oipchul.FRectDiffOnly = "Y"

oipchul.GetIpgoToAgvDiffList

%>
<script src="/cscenter/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript">

function SubmitFrm(frm) {
    frm.submit();
}

</script>

<!-- �˻� ���� -->
<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="page" value="">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		* �԰�����(��ǰ�Ϸ� �԰����� ����) : <% DrawOneDateBox yyyy1,mm1,dd1 %>
        &nbsp;
        <input type="checkbox" name="diffonly" value="Y" <%= CHKIIF(diffonly="Y", "checked", "") %>> ���������� ǥ��
	</td>

	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="SubmitFrm(frm)">
	</td>
</tr>
</table>
</form>
<!-- �˻� �� -->

<br />

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15" align="left">
		�˻���� : <b><%= oipchul.FResultCount %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="130">skuCd</td>
	<td width="30">����</td>
	<td width=80>��ǰ�ڵ�</td>
    <td width="40">�ɼ�</td>

    <td>�귣��</td>
    <td>��ǰ��</td>
    <td>�ɼǸ�</td>

    <td width="50">����<br />����</td>
    <td width="50">��ǰ<br />����</td>
    <td width="50">Ȯ��<br />����</td>
    <td width="50">AGV�԰�<br />����</td>

    <td width=70>���ڵ�</td>
    <td width=70>���ڵ�</td>
    <td width=70>���ڵ��</td>
	<td>���</td>
</tr>
<% if oipchul.FResultCount > 0 then %>
<% for i=0 to oipchul.FResultcount-1 %>

<tr bgcolor="#FFFFFF" height=24>
	<td align=center><%= oipchul.FItemList(i).Fskucd %></td>
    <td align=center><%= oipchul.FItemList(i).Fitemgubun %></td>
    <td align=center><%= oipchul.FItemList(i).Fitemid %></td>
    <td align=center><%= oipchul.FItemList(i).Fitemoption %></td>

    <td align=center><%= oipchul.FItemList(i).Fmakerid %></td>
    <td align=center><%= oipchul.FItemList(i).Fitemname %></td>
    <td align=center><%= oipchul.FItemList(i).Fitemoptionname %></td>

    <td align=center><%= oipchul.FItemList(i).Fbaljuitemno %></td>
    <td align=center><%= oipchul.FItemList(i).Fcheckitemno %></td>
    <td align=center><%= oipchul.FItemList(i).Frealitemno %></td>
    <td align=center><%= oipchul.FItemList(i).Fagvipgoitemno %></td>

    <td align=center>
        <%= oipchul.FItemList(i).FlocationCd1 %>
    </td>
    <td align=center>
        <% if Not IsNull(oipchul.FItemList(i).FlocationCd1) and Not IsNull(oipchul.FItemList(i).FlocationCd2) then %>
        <% if (oipchul.FItemList(i).FlocationCd1 <> oipchul.FItemList(i).FlocationCd2) then %>
        <%= oipchul.FItemList(i).FlocationCd2 %>
        <% end if %>
        <% end if %>
    </td>
    <td align=center>
        <%= oipchul.FItemList(i).FlocationCdCnt %>
    </td>
    <td align=center></td>
</tr>

<% next %>
<% else %>
<tr bgcolor="#FFFFFF">
	<td colspan="15" align=center>[ �˻������ �����ϴ�. ]</td>
</tr>
<% end if %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
