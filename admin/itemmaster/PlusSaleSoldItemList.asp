<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/PlusSaleCls.asp"-->
<%

dim buy_benefit_idx
dim i, j, k
dim page : page = 1

buy_benefit_idx = request("buy_benefit_idx")
if (buy_benefit_idx <> "") then
    if Not IsNumeric(buy_benefit_idx) then
        buy_benefit_idx = ""
    end if
end if

if (buy_benefit_idx = "") then
    response.write "�߸��� �����Դϴ�."
    dbget.close : response.end
end if


'// ===============================================
'// ���ñ׷�
'// ===============================================
dim oCBuyBenefitGroupItem
set oCBuyBenefitGroupItem = new CBuyBenefit
oCBuyBenefitGroupItem.FRectBuyBenefitIdx = buy_benefit_idx
oCBuyBenefitGroupItem.FPageSize = 100
oCBuyBenefitGroupItem.FCurrPage = page

oCBuyBenefitGroupItem.GetBuyBenefitSoldItemList

%>
<script language='javascript'>

</script>
<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="17">
        <img src="/images/icon_arrow_down.gif" align="absbottom">
        <font color="red"><strong>���ñ׷��ǰ ���</strong></font>
		�˻���� : <b><%= oCBuyBenefitGroupItem.FTotalCount %></b>
		&nbsp;
		������ : <b><%= page %> / <%= oCBuyBenefitGroupItem.FTotalPage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="25">
	<td width=100>��ǰ�ڵ�</td>
    <td width=60>�ɼ�</td>
    <td>��ǰ��</td>
    <td>�ɼǸ�</td>
    <td width=80>����</td>
    <td width=120>����</td>
    <td>���</td>
</tr>
<% for i=0 to oCBuyBenefitGroupItem.FResultcount-1 %>
<tr bgcolor="#FFFFFF" align="center" height="25">
    <td><%= oCBuyBenefitGroupItem.FItemList(i).Fitemid %></td>
    <td><%= oCBuyBenefitGroupItem.FItemList(i).Fitemoption %></td>
    <td><%= oCBuyBenefitGroupItem.FItemList(i).Fitemname %></td>
    <td><%= oCBuyBenefitGroupItem.FItemList(i).Fitemoptionname %></td>
    <td><%= oCBuyBenefitGroupItem.FItemList(i).Fitemno %></td>
    <td><%= oCBuyBenefitGroupItem.FItemList(i).Fmeachul %></td>
    <td>
    </td>
</tr>
<% next %>
</table>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/db3close.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
