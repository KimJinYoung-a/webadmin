<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 매입상품원가관리 엑셀다운로드
' History : 2022.10.18 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/PurchasedProductCls.asp"-->
<%
dim i, research, page, ExcDel, productidx, makerid, purchasetype, codelist, reportIdx, itemid, arrLIst, menupos
	productidx = requestCheckVar(trim(getNumeric(request("productidx"))),8)
	makerid = requestCheckVar(trim(request("makerid")),32)
	purchasetype = requestCheckVar(request("purchasetype"),2)
	codelist = requestCheckVar(request("codelist"),32)
	reportIdx = requestCheckVar(trim(getNumeric(request("reportIdx"))),8)
	itemid      = requestCheckvar(request("itemid"),1500)
page = requestCheckVar(request("page"),8)
ExcDel = requestCheckVar(request("ExcDel"),1)
research = requestCheckVar(request("research"),1)
menupos = requestCheckVar(trim(getNumeric(request("menupos"))),10)

if page = "" then page = "1"
if ExcDel = "" and research="" then ExcDel = "Y"
if itemid<>"" then
	dim iA ,arrTemp,arrItemid
  itemid = replace(itemid,chr(13),"")
	arrTemp = Split(itemid,chr(10))

	iA = 0
	do while iA <= ubound(arrTemp)
		if Trim(arrTemp(iA))<>"" and isNumeric(Trim(arrTemp(iA))) then
			arrItemid = arrItemid & Trim(arrTemp(iA)) & ","
		end if
		iA = iA + 1
	loop

	if len(arrItemid)>0 then
		itemid = left(arrItemid,len(arrItemid)-1)
	else
		if Not(isNumeric(itemid)) then
			itemid = ""
		end if
	end if
end if

dim oCPurchasedProduct
set oCPurchasedProduct = new CPurchasedProduct
	oCPurchasedProduct.FCurrPage = page
	oCPurchasedProduct.Fpagesize = 1000000
    oCPurchasedProduct.FRectExcDel = ExcDel
	oCPurchasedProduct.FRectproductidx = productidx
	oCPurchasedProduct.FRectpurchasetype = purchasetype
	oCPurchasedProduct.FRectmakerid = makerid
	oCPurchasedProduct.FRectcodelist = codelist
	oCPurchasedProduct.FRectreportIdx = reportIdx
	oCPurchasedProduct.FRectItemid       = itemid
	oCPurchasedProduct.GetPurchasedProductMasterListNotPaging

if oCPurchasedProduct.FTotalCount>0 then
    arrLIst=oCPurchasedProduct.fArrLIst
end if

Response.Expires=0
response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=TENPurchasedProduct" & Left(CStr(now()),10) & "_" & session.sessionID & ".xls"
Response.CacheControl = "public"
Response.Buffer = true    '버퍼사용여부
%>
<html>
<head>
<style type='text/css'>
	.txt {mso-number-format:'\@'}
</style>
</head>
<body>
<table width="100%" align="center" cellpadding="3" cellspacing="1" border=1 bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="FFFFFF">
	<td colspan="14">
		검색결과 : <b><%= oCPurchasedProduct.FTotalCount %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>IDX</td>
	<td>적요</td>
	<td>브랜드ID</td>
	<td>주문코드</td>
    <td>품의번호</td>
    <td>품의금액</td>
    <td>주문수량</td>
    <td>주문금액</td>
    <td>입고수량</td>
    <td>입고금액</td>
	<td>결제진행중</td>
	<td>결제액</td>
    <td>등록자</td>
    <td>비고</td>
</tr>
<% if isarray(arrLIst) then %>
<%
for i=0 to ubound(arrLIst,2)
%>
<tr bgcolor="#FFFFFF" align="center">
    <td><%= arrLIst(0,i) %></td>
	<td align="left"><%= arrLIst(14,i) %></td>
	<td class="txt"><%= arrLIst(17,i) %></td>
    <td><%= arrLIst(1,i) %></td>
    <td><%= arrLIst(2,i) %></td>
    <td align="right"><%= FormatNumber(arrLIst(15,i), 0) %></td>
    <td align="right"><%= FormatNumber(arrLIst(5,i), 0) %></td>
    <td align="right"><%= FormatNumber(arrLIst(6,i), 0) %></td>
    <td align="right"><%= FormatNumber(arrLIst(7,i), 0) %></td>
    <td align="right"><%= FormatNumber(arrLIst(8,i), 0) %></td>
	<td align="right"><%= FormatNumber(arrLIst(19,i), 0) %></td>
	<td align="right"><%= FormatNumber(arrLIst(18,i), 0) %></td>
    <td><%= arrLIst(10,i) %></td>
    <td></td>
</tr>
<%
if i mod 1000 = 0 then
    Response.Flush		' 버퍼리플래쉬
end if
next
%>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="14" align="center">[검색결과가 없습니다.]</td>
	</tr>
<% end if %>
</table>
</body>
</html>
<%
set oCPurchasedProduct = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->