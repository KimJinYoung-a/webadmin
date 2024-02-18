<%@ language=vbscript %>
<% option explicit %>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
%>
<%
'####################################################
' Description :  �ؿܻ�ǰ �Ӽ� ���� ���� �ٿ�ε塩
' History : 2019.11.21 ������ �߰�
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/items/itemcls_2008.asp"-->

<%
dim itemid, itemname, makerid, sellyn, usingyn, mwdiv, limityn, overSeaYn, weightYn, itemrackcode,research
dim cdl, cdm, cds, sortDiv, page, limitrealstock, stocktype, i, pojangok, itemManageType, sizeYn, arrlist
dim itemdivNotexists
	itemid		= request("itemid")
	itemname	= requestCheckVar(request("itemname"),128)
	makerid		= requestCheckVar(request("makerid"),32)
	sellyn		= requestCheckVar(request("sellyn"),1)
	usingyn		= requestCheckVar(request("usingyn"),1)
	mwdiv		= requestCheckVar(request("mwdiv"),32)
	limityn		= requestCheckVar(request("limityn"),1)
	overSeaYn	= requestCheckVar(request("overSeaYn"),1)
	weightYn	= requestCheckVar(request("weightYn"),1)
	itemrackcode= requestCheckVar(request("itemrackcode"),32)
	sortDiv		= requestCheckVar(request("sortDiv"),32)
	research	=requestCheckVar(Request("research"),1)
	pojangok	=requestCheckVar(Request("pojangok"),1)
	cdl = requestCheckVar(request("cdl"),32)
	cdm = requestCheckVar(request("cdm"),32)
	cds = requestCheckVar(request("cds"),32)
	page = requestCheckVar(request("page"),32)
	limitrealstock = requestCheckVar(request("limitrealstock"),32)
	stocktype = requestCheckVar(request("stocktype"),32)
	itemManageType = requestCheckVar(request("itemManageType"),32)
	sizeYn = requestCheckVar(request("sizeYn"),32)
	itemdivNotexists = requestCheckVar(request("itemdivNotexists"),32)

'�⺻��
if (page="") then page=1
if sortDiv="" then sortDiv="new"
if research="" then
	if mwdiv="" then mwdiv="MW"
	if overSeaYn="" then overSeaYn="Y"
	if weightYn="" then weightYn="Y"
	'if pojangok="" then pojangok="Y"
	itemManageType = "I"
end if
if research="" and itemdivNotexists="" then
	itemdivNotexists="on"
end if
if (stocktype = "") then
	stocktype = "sys"
end if

if itemid<>"" then
	dim iA ,arrTemp,arrItemid

	arrTemp = Split(itemid,",")

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

dim oitem
set oitem = new CItem
	oitem.FPageSize         = 200000
	oitem.FCurrPage         = 1
	oitem.FRectMakerid      = makerid
	oitem.FRectItemid       = itemid
	oitem.FRectItemName     = itemname
	oitem.FRectSellYN       = sellyn
	oitem.FRectIsUsing      = usingyn
	oitem.FRectLimityn      = limityn
	oitem.FRectMWDiv        = mwdiv
	oitem.FRectIsOversea	= overSeaYn
	oitem.FRectIsWeight		= weightYn
	oitem.FRectRackcode		= itemrackcode
	oitem.FRectCate_Large   = cdl
	oitem.FRectCate_Mid     = cdm
	oitem.FRectCate_Small   = cds
	oitem.FRectSortDiv		= sortDiv
	oitem.FRectlimitrealstock = limitrealstock
	oitem.FRectStockType = stocktype
	oitem.FRectpojangok = pojangok
	oitem.FRecItemManageType = itemManageType
	oitem.FRectSizeYn = sizeYn

	if itemdivNotexists="on" then
		oitem.frectitemdivNotexists="'08','21'"
	end if

	oitem.GetItemAboardList_notpaging

if oitem.FTotalCount>0 then
	arrlist = oitem.farrlist
end if

Response.Buffer = true    '���ۻ�뿩��
Response.Expires=0
response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=TEN_" & Left(CStr(now()),10) & "_item_aboard_list.xls"
Response.CacheControl = "public"
%>

<html>
<head>
<% '<meta http-equiv="Content-Type" content="text/html; charset="euc-kr"> %>
<meta http-equiv='Content-Type' content='text/html;charset=euc-kr' />
<title>�ؿܻ�ǰ�Ӽ���</title>
<style type='text/css'>
	.txt {mso-number-format:'\@'}
</style>
</head>
<body>

<table width="100%" align="center" border="1" cellpadding="3" cellspacing="0" bgcolor="#c1c1c1">
<tr bgcolor="#c1c1c1">
	<td>Rack</td>
	<td>No.</td>
	<td>�귣��ID</td>
	<td>��ǰ��</td>
	<td>�ǸŰ�</td>
	<td>��౸��</td>
	<td>�Ǹſ���</td>
	<td>��뿩��</td>
	<td>�ؿܿ���</td>
	<td>���尡�ɿ���</td>
	<td>��ǰ����</td>
	<td>��ǰ������</td>
</tr>

<% if isarray(arrlist) then %>
<% for i=0 to ubound(arrlist,2) %>
<tr bgcolor="#ffffff">
    <td><%= arrlist(38,i) %></td>
    <td><%= arrlist(0,i) %></td>
    <td><%= arrlist(1,i) %></td>
    <td><%= arrlist(7,i) %></td>
    <td>
    <%
        Response.Write FormatNumber(arrlist(10,i),0)
        '���ΰ�
        if arrlist(21,i)="Y" then
            Response.Write "(��)" & FormatNumber(arrlist(12,i),0)
        end if
        '������
        if arrlist(50,i)="Y" then
            Select Case arrlist(52,i)
                Case "1"
                    Response.Write "(��)" & FormatNumber(arrlist(10,i)*((100-arrlist(53,i))/100),0)
                Case "2"
                    Response.Write "(��)" & FormatNumber(arrlist(10,i)-arrlist(53,i),0)
            end Select
        end if
    %>
    </td>
    <td><%= fnColor(arrlist(24,i),"mw") %></td>
    <td><%= fnColor(arrlist(18,i),"yn") %></td>
    <td><%= fnColor(arrlist(22,i),"yn") %></td>
    <td><%= fnColor(arrlist(56,i),"yn") %></td>
    <td><%= fnColor(arrlist(32,i),"yn") %></td>
    <td><%= FormatNumber(arrlist(55,i),0) %>g</td>
    <td>
        <%= arrlist(84,i) %> * <%= arrlist(85,i) %> * <%= arrlist(86,i) %> cm
    </td>
</tr>
<%
if i mod 1000 = 0 then
	Response.Flush		' ���۸��÷���
end if
next
%>

<% end if %>
</table>
</body>
</html>

<%
set oitem = nothing
 %>
<!-- #include virtual="/lib/db/dbclose.asp" -->
