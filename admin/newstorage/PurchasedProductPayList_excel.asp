<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ������������Ʈ �����ٿ�ε�
' History : 2023.09.22 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/PurchasedProductCls.asp"-->
<!-- #include virtual="/lib/classes/approval/payrequestCls.asp"-->
<%

dim i, research, page, ExcDel, productidx, sheetidx, makerid, purchasetype, codelist, reportIdx, itemid, arrLIst, menupos
	productidx = requestCheckVar(trim(getNumeric(request("productidx"))),8)
	sheetidx = requestCheckVar(trim(getNumeric(request("sheetidx"))),8)
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

dim oCPurchasedProductPay
set oCPurchasedProductPay = new CPurchasedProduct
	oCPurchasedProductPay.FCurrPage = page
	oCPurchasedProductPay.Fpagesize = 1000000
    oCPurchasedProductPay.FRectExcDel = ExcDel
	oCPurchasedProductPay.FRectproductidx = productidx
    oCPurchasedProductPay.FRectSheetidx = sheetidx
	oCPurchasedProductPay.FRectpurchasetype = purchasetype
	oCPurchasedProductPay.FRectmakerid = makerid
	oCPurchasedProductPay.FRectcodelist = codelist
	oCPurchasedProductPay.FRectreportIdx = reportIdx
	oCPurchasedProductPay.FRectItemid       = itemid
	oCPurchasedProductPay.GetPurchasedProductItemAllPayListNotPaging

if oCPurchasedProductPay.FTotalCount>0 then
    arrLIst=oCPurchasedProductPay.fArrLIst
end if

Response.Expires=0
response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=TENPurchasedProductPayList" & Left(CStr(now()),10) & "_" & session.sessionID & ".xls"
Response.CacheControl = "public"
Response.Buffer = true    '���ۻ�뿩��
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
	<td colspan="11">
		�˻���� : <b><%= oCPurchasedProductPay.FTotalCount %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>������IDX</td>
    <td>ǰ�ǹ�ȣ</td>
    <td>ǰ�Ǳݾ�</td>
    <td>������û��IDX</td>
    <td>������û��</td>
    <td>������</td>
    <td>������û�ݾ�(��)</td>
    <td>�������</td>
    <td>�ڱݿ뵵</td>
    <td>�ŷ�ó</td>
    <td>����</td>
</tr>
<% if isarray(arrLIst) then %>
<%
for i=0 to ubound(arrLIst,2)
%>
<tr bgcolor="#FFFFFF" align="center">
    <td><%= arrLIst(0,i) %></td>
    <td align="center"><%= arrLIst(2,i) %></td>
    <td align="right"><%= FormatNumber(arrLIst(3,i), 0) %></td>
    <td align="center"><%= arrLIst(4,i) %></td>
    <td align="center"><%= arrLIst(5,i) %></td>
    <td align="center"><%= arrLIst(11,i) %></td>
    <td align="right"><%= FormatNumber(arrLIst(6,i), 0) %></td>
    <td align="center"><%= fnGetPayType(arrLIst(7,i)) %></td>
    <td align="center"><%= arrLIst(10,i) %></td>
    <td align="center"><%= arrLIst(8,i) %></td>
    <td align="center"><%= fnGetPayRequestState(arrLIst(9,i)) %></td>
</tr>
<%
if i mod 500 = 0 then
    Response.Flush		' ���۸��÷���
end if
next
%>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="11" align="center">[�˻������ �����ϴ�.]</td>
	</tr>
<% end if %>
</table>
</body>
</html>
<%
set oCPurchasedProductPay = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->