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
    response.write "잘못된 접근입니다."
    dbget.close : response.end
end if


'// ===============================================
'// 혜택그룹
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
<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="17">
        <img src="/images/icon_arrow_down.gif" align="absbottom">
        <font color="red"><strong>혜택그룹상품 목록</strong></font>
		검색결과 : <b><%= oCBuyBenefitGroupItem.FTotalCount %></b>
		&nbsp;
		페이지 : <b><%= page %> / <%= oCBuyBenefitGroupItem.FTotalPage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="25">
	<td width=100>상품코드</td>
    <td width=60>옵션</td>
    <td>상품명</td>
    <td>옵션명</td>
    <td width=80>수량</td>
    <td width=120>매출</td>
    <td>비고</td>
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
