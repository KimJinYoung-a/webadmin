<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/company/incSessionCompany.asp" -->
<!-- #include virtual="/company/ch/incGlobalVariable.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/overseas/overseasCls.asp"-->


<%

dim itemid, itemname, makerid, sellyn, usingyn, mwdiv, limityn, overSeaYn, weightYn, itemrackcode, vRegUserID, vIsReg
dim cdl, cdm, cds, sortDiv
dim page

itemid		= request("itemid")
itemname	= request("itemname")
makerid		= request("makerid")
sellyn		= request("sellyn")
usingyn		= request("usingyn")
mwdiv		= request("mwdiv")
limityn		= request("limityn")
overSeaYn	= request("overSeaYn")
weightYn	= request("weightYn")
itemrackcode= request("itemrackcode")
sortDiv		= request("sortDiv")
vRegUserID	= request("reguserid")
vIsReg		= request("isreg")

cdl = request("cdl")
cdm = request("cdm")
cds = request("cds")


'기본값
if mwdiv="" then mwdiv="MW"
if overSeaYn="" then overSeaYn="Y"
if weightYn="" then weightYn="Y"
if sortDiv="" then sortDiv="new"


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


'==============================================================================
dim oitem, arrlist

set oitem = new COverSeasItem

oitem.FPageSize         = 1000000
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

oitem.FRectRegUserID	= vRegUserID
oitem.FRectIsReg		= vIsReg

oitem.GetOverSeasTargetItemListXLS

if oitem.FResultCount > 0 then
	arrlist = oitem.farrlist
end if

dim i, vBody

Response.Buffer=False
Response.Expires=0
'response.ContentType = "application/vnd.ms-excel"
'Response.AddHeader "Content-Disposition", "attachment; filename=상품리스트.csv"
'Response.CacheControl = "public"

Response.ContentType = "application/x-msexcel"
Response.CacheControl = "public"
Response.AddHeader "Content-Disposition", "attachment;filename=상품리스트.xls"


if isarray(arrlist) then
%>
<html>
<body>
<table>
<% for i=0 to ubound(arrlist,2) %>
<tr>
	<td></td>
	<td><%=arrlist(0,i)%></td>
	<td><%=arrlist(1,i)%></td>
	<td></td>
	<td><%=db2html(arrlist(2,i))%></td>
	<td></td>
	<td><%=db2html(arrlist(3,i))%></td>
	<td></td>
	<td><%=db2html(arrlist(4,i))%></td>
	<td></td>
	<td><%=db2html(arrlist(5,i))%></td>
	<td></td>
	<td><%=db2html(arrlist(6,i))%></td>
	<td></td>
	<td><%=db2html(arrlist(18,i))%></td>
	<td><%=db2html(arrlist(17,i))%></td>
	<td></td>
	<td></td>
	<td></td>
	<td><%=arrlist(7,i)%></td>
	<td><%=arrlist(8,i)%></td>
	<td><%=arrlist(9,i)%></td>
	<td></td>
	<td><%=db2html(arrlist(12,i))%></td>
	<td></td>
	<td><%=arrlist(10,i)%></td>
	<td></td>
	<td></td>
	<td></td>
	<td></td>
	<td><%=arrlist(11,i)%></td>
	<td></td>
	<td></td>
	<td></td>
	<td style="mso-number-format:'\@'"><%=db2html(arrlist(15,i))%></td>
	<td><%=db2html(arrlist(16,i))%></td>
	<td><%=db2html(arrlist(13,i))%></td>
	<td><%=db2html(arrlist(14,i))%></td>
</tr>
<% next %>
</table>
</body>
</html>
<%
end if

set oitem = nothing

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->