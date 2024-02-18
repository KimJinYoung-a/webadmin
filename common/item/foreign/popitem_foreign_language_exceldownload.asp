<%@ language=vbscript %>
<% option explicit %>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
%>
<%
'####################################################
' Description :  오프라인 주문서
' History : 2016.09.05 한용민 생성
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/items/overseas/overseasCls.asp"-->

<%
Dim oitem, i, page, itemid, itemname, makerid, cdl, cdm, cds, vCountryCd, sellyn, usingyn, limityn, danjongyn, sitecountrylang
Dim sitename, vpriceArrv, vpriceArrk, v, k, reloading, vcountryLangCDArrv, vcountryLangCDArrk, mwdiv, overSeaYn
dim weightYn, sellcash1, sellcash2, sortDiv
	page = requestCheckvar(request("page"),10)
	vCountryCd	= requestCheckvar(request("countrycd"),2)
	itemid      = requestCheckvar(request("itemid"),255)
	itemname	= requestCheckvar(request("itemname"),64)
	makerid		= requestCheckvar(request("makerid"),32)
	sellyn		= requestCheckvar(request("sellyn"),1)
	usingyn		= requestCheckvar(request("usingyn"),1)
	cdl = requestCheckvar(request("cdl"),3)
	cdm = requestCheckvar(request("cdm"),3)
	cds = requestCheckvar(request("cds"),3)
	limityn = requestCheckvar(request("limityn"),1)
    sitename = requestCheckvar(request("sitename"),32)
    reloading		= requestCheckvar(request("reloading"),2)
	danjongyn   = requestCheckvar(request("danjongyn"),10)
	mwdiv		= requestCheckvar(request("mwdiv"),2)
	overSeaYn	= requestCheckvar(request("overSeaYn"),1)
	weightYn	= requestCheckvar(request("weightYn"),1)
	sellcash1	= requestCheckvar(request("sellcash1"),10)
	sellcash2	= requestCheckvar(request("sellcash2"),10)
	sortDiv		= requestCheckvar(request("sortDiv"),16)

if (page = "") then page = 1
if sitename="" then sitename="WSLWEB"
if (vCountryCd = "") then vCountryCd = "o"
if reloading="" and sellyn="" then sellyn="YS"
if reloading="" and usingyn="" then usingyn="Y"
if reloading<>"ON" and overSeaYn="" then overSeaYn="Y"
if sortDiv="" then sortDiv="new"

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

set oitem = new COverSeasItem
	oitem.FPageSize         = 10000
	oitem.FCurrPage         = 1
	oitem.FRectCountryCd	= vCountryCd
	oitem.FRectMakerid      = makerid
	oitem.FRectSellYN       = sellyn
	oitem.FRectIsUsing      = usingyn
	oitem.FRectItemid       = itemid
	oitem.FRectItemName     = itemname
	oitem.FRectCate_Large   = cdl
	oitem.FRectCate_Mid     = cdm
	oitem.FRectCate_Small   = cds
	oitem.FRectLimitYN		= limityn
    oitem.FRectSitename = Sitename
	oitem.FRectDanjongyn    = danjongyn
	oitem.FRectMWDiv        = mwdiv
	oitem.FRectIsOversea	= overSeaYn
	oitem.FRectIsWeight		= weightYn
	oitem.FRectSellcash1	= sellcash1
	oitem.FRectSellcash2	= sellcash2
	oitem.FRectsortDiv	= sortDiv

	If sitename <> "" Then
		oitem.GetOverSeasItemList_excel
	End If

Response.Expires=0
response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=TEN_" & Left(CStr(now()),10) & "_item_foreign_language.xls"
Response.CacheControl = "public"
Response.Buffer = true    '버퍼사용여부
%>

<html>
<head>
<% '<meta http-equiv="Content-Type" content="text/html; charset="euc-kr"> %>
<meta http-equiv='Content-Type' content='text/html;charset=euc-kr' />
<title>텐바이텐주문서</title>
<style type='text/css'>
	.txt {mso-number-format:'\@'}
</style>
</head>
<body>

<table width="100%" align="center" border="1" cellpadding="3" cellspacing="0" bgcolor="#c1c1c1">
<tr bgcolor="#c1c1c1">
	<td>Item Code</td>
	<td>Item Name[EN]</td>
	<td>Option delimitation[EN]</td>
	<td>Option Name[EN]</td>
	<td>ItemCopy[EN]</td>
	<td>Material[EN]</td>
	<td>Size[EN]</td>
	<td>Manufacturer[EN]</td>
	<td>Origin[EN]</td>
	<td>Keyword[EN]</td>
	<td>Brand</td>
	<td>Item Name[KR]</td>
	<td>Option delimitation[KR]</td>
	<td>Option Name[KR]</td>
	<td>ItemCopy[KR]</td>
	<td>Material[KR]</td>
	<td>Size[KR]</td>
	<td>Manufacturer[KR]</td>
	<td>Origin[KR]</td>
	<td>Keyword[KR]</td>
</tr>

<% if oitem.FresultCount > 0 then %>
<% for i=0 to oitem.FresultCount-1 %>
	<tr bgcolor="#ffffff">
		<td class='txt' bgcolor="#e1e1e1">
			<%= oitem.FItemList(i).fitemgubun %>
			<%= CHKIIF(oitem.FItemList(i).Fitemid>=1000000,Format00(8,oitem.FItemList(i).Fitemid),Format00(6,oitem.FItemList(i).Fitemid)) %>
			<%= oitem.FItemList(i).fitemoption %>
		</td>
		<td class='txt'><%= oitem.FItemList(i).fitemname_en %></td>
		<td class='txt'><%= oitem.FItemList(i).foptiontypename_en %></td>
		<td class='txt'><%= oitem.FItemList(i).foptionname_en %></td>
		<td class='txt'><%= oitem.FItemList(i).fitemcopy_en %></td>
		<td class='txt'><%= oitem.FItemList(i).fitemsource_en %></td>
		<td class='txt'><%= oitem.FItemList(i).fitemsize_en %></td>
		<td class='txt'><%= oitem.FItemList(i).fmakername_en %></td>
		<td class='txt'><%= oitem.FItemList(i).fsourcearea_en %></td>
		<td class='txt'><%= oitem.FItemList(i).fkeywords_en %></td>
		<td class='txt' bgcolor="#e1e1e1"><%= oitem.FItemList(i).fmakerid %></td>
		<td class='txt' bgcolor="#e1e1e1"><%= oitem.FItemList(i).fitemname_10x10 %></td>
		<td class='txt' bgcolor="#e1e1e1"><%= oitem.FItemList(i).foptiontypename_10x10 %></td>
		<td class='txt' bgcolor="#e1e1e1"><%= oitem.FItemList(i).foptionname_10x10 %></td>
		<td class='txt' bgcolor="#e1e1e1"><%= oitem.FItemList(i).fitemcopy_10x10 %></td>
		<td class='txt' bgcolor="#e1e1e1"><%= oitem.FItemList(i).fitemsource_10x10 %></td>
		<td class='txt' bgcolor="#e1e1e1"><%= oitem.FItemList(i).fitemsize_10x10 %></td>
		<td class='txt' bgcolor="#e1e1e1"><%= oitem.FItemList(i).fmakername_10x10 %></td>
		<td class='txt' bgcolor="#e1e1e1"><%= oitem.FItemList(i).fsourcearea_10x10 %></td>
		<td class='txt' bgcolor="#e1e1e1"><%= oitem.FItemList(i).fkeywords_10x10 %></td>
	</tr>
<%
	if i mod 3000 = 0 then
		Response.Flush		' 버퍼리플래쉬
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
