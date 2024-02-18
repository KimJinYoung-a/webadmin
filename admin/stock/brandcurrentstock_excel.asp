<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  브랜드별재고현황 엑셀다운로드
' History : 2021.02.16 한용민 생성
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stock/summary_itemstockcls.asp"-->
<!-- #include virtual="/lib/BarcodeFunction.asp"-->
<%
dim arrlist

dim makerid, onoffgubun, mwdiv, research, sellyn, usingyn, danjongyn, osummarystockbrand, centermwdiv
dim returnitemgubun, itemname, itemidArr, cdl, cdm, cds, page, i, BasicMonth, limitrealstock
dim totsysstock, totavailstock, totrealstock, totjeagosheetstock, totmaystock, IsSheetPrintEnable
dim stocktype, useoffinfo, itemgubun, startMon, endMon, excits, pagesize, ordby, vPurchasetype
dim limityn, itemgrade, itemrackcode, bulkstockgubun, warehouseCd, agvstockgubun
Dim dispCate : dispCate = RequestCheckvar(Request("disp"),12)
	makerid         = requestCheckvar(request("makerid"),32)
	onoffgubun      = requestCheckvar(request("onoffgubun"),9)
	research        = requestCheckvar(request("research"),9)
	sellyn          = requestCheckvar(request("sellyn"),9)
	usingyn         = requestCheckvar(request("usingyn"),9)
	danjongyn       = requestCheckvar(request("danjongyn"),9)
	mwdiv           = requestCheckvar(request("mwdiv"),9)
	returnitemgubun = requestCheckvar(request("returnitemgubun"),9)
	itemname        = requestCheckvar(request("itemname"),64)
	itemidArr       = Trim(requestCheckvar(request("itemidArr"),255))
	page            = requestCheckvar(request("page"),9)
	cdl = requestCheckvar(request("cdl"),3)
	cdm = requestCheckvar(request("cdm"),3)
	cds = requestCheckvar(request("cds"),3)
	limitrealstock 	= requestCheckvar(request("limitrealstock"),10)
    centermwdiv    	= requestCheckvar(request("centermwdiv"),10)
	stocktype    	= requestCheckvar(request("stocktype"),32)
	itemgubun     	= RequestCheckVar(request("itemgubun"),32)
	startMon     	= RequestCheckVar(request("startMon"),32)
	endMon     		= RequestCheckVar(request("endMon"),32)
	useoffinfo = request("useoffinfo")
	excits  		= requestCheckvar(request("excits"),2)
	pagesize  		= requestCheckvar(request("pagesize"),4)
	ordby    		= requestCheckvar(request("ordby"),64)
	vPurchasetype 	= request("purchasetype")
    limityn  		= requestCheckvar(request("limityn"),2)
    itemgrade     	= RequestCheckVar(request("itemgrade"),32)
    itemrackcode    = RequestCheckVar(request("itemrackcode"),32)
    bulkstockgubun  = RequestCheckVar(request("bulkstockgubun"),32)
    warehouseCd  	= RequestCheckVar(request("warehouseCd"),32)
    agvstockgubun  	= RequestCheckVar(request("agvstockgubun"),32)

if (stocktype = "") then stocktype = "sys"
if (pagesize = "") then pagesize = 25

'///////////////// 바코드 프린트기 설정 ///////////////////////
dim printername, printpriceyn, titledispyn, isforeignprint, makeriddispyn, useforeigndata, currencyunit, currencyChar
	printername = requestCheckVar(request("printername"),32)
	printpriceyn 	= requestCheckVar(request("printpriceyn"),1)
	titledispyn = requestCheckVar(request("titledispyn"),1)
	isforeignprint 	= requestCheckVar(request("isforeignprint"),32)
	makeriddispyn 			= requestCheckVar(request("makeriddispyn"),32)

if printpriceyn = "" then printpriceyn = "Y"	' R
if printername = "" then printername = "TEC_B-FV4_80x50"	' TEC_B-FV4_45x22
if makeriddispyn = "" then makeriddispyn = "Y"
if titledispyn = "" then titledispyn = "Y"
useforeigndata = "N"
currencyunit = "KRW"
currencyChar = "￦"
'/////////////////'////////////////////////////////////////

'//상품코드 유효성 검사
if itemidArr<>"" then
	dim iA ,arrTemp,arrItemid
  itemidArr = replace(itemidArr,chr(13),"")
	arrTemp = Split(itemidArr,chr(10))

	iA = 0
	do while iA <= ubound(arrTemp)
		if Trim(arrTemp(iA))<>"" and isNumeric(Trim(arrTemp(iA))) then
			arrItemid = arrItemid & Trim(arrTemp(iA)) & ","
		end if
		iA = iA + 1
	loop

	if len(arrItemid)>0 then
		itemidArr = left(arrItemid,len(arrItemid)-1)
	else
		if Not(isNumeric(itemidArr)) then
			itemidArr = ""
		end if
	end if
end if


if (request("research") = "") then
	excits = "Y"
end if

if (page="") then page=1
''if onoffgubun="" then onoffgubun="on"
''if itemgubun="" then itemgubun="10"
BasicMonth = Left(CStr(DateSerial(Year(now()),Month(now())-1,1)),7)


'// onoffgubun => itemgubun, skyer9, 2016-06-21
if (onoffgubun = "") and (itemgubun = "") then
	itemgubun="10"
elseif (onoffgubun <> "") and (itemgubun = "") then
	if (onoffgubun = "on") then
		itemgubun="10"
	elseif (onoffgubun = "off") then
		itemgubun="exc10"
	else
		itemgubun = Right(onoffgubun,2)
	end if
end if
if itemgubun="" then itemgubun="10"

if itemgubun = "10" then
	onoffgubun = "on"
elseif (itemgubun = "exc10") then
	onoffgubun = "off"
elseif (itemgubun <> "10") then
	onoffgubun = "off" & itemgubun
end if


set osummarystockbrand = new CSummaryItemStock
	osummarystockbrand.FPageSize = 1000000
	osummarystockbrand.FCurrPage = 1
	osummarystockbrand.FRectCD1   = cdl
	osummarystockbrand.FRectCD2   = cdm
	osummarystockbrand.FRectCD3   = cds
	osummarystockbrand.FRectItemIdArr = itemidArr
	osummarystockbrand.FRectItemName = itemname
	osummarystockbrand.FRectMakerid = makerid
	osummarystockbrand.FRectOnlySellyn = sellyn
	osummarystockbrand.FRectOnlyIsUsing = usingyn
	osummarystockbrand.FRectDanjongyn =danjongyn
	osummarystockbrand.FRectMWDiv = mwdiv
	osummarystockbrand.FRectCenterMWDiv = centermwdiv
	osummarystockbrand.FRectReturnItemGubun = returnitemgubun
	osummarystockbrand.FRectlimitrealstock = limitrealstock
	osummarystockbrand.FRectStockType = stocktype
	osummarystockbrand.FRectDispCate = dispCate
	osummarystockbrand.FRectUseOffInfo = useoffinfo
	osummarystockbrand.FRectExcIts = excits
	osummarystockbrand.FRectPurchasetype = vPurchasetype
    osummarystockbrand.FRectLimitYN = limityn
    osummarystockbrand.FRectItemGrade = itemgrade
    osummarystockbrand.FRectRackCode = itemrackcode
    osummarystockbrand.FRectBulkStockGubun = bulkstockgubun
    osummarystockbrand.FRectWarehouseCd = warehouseCd
    osummarystockbrand.FRectAgvStockGubun = agvstockgubun

	if (ordby = "1") then
		osummarystockbrand.FRectOrderBy = "T.itemid desc"
	elseif (ordby = "2") then
		osummarystockbrand.FRectOrderBy = "T.itemrackcode asc,T.itemid desc"
	end if

	if IsNumeric(startMon) then
		osummarystockbrand.FRectStartDate = startMon
	elseif (startMon <> "") then
		response.write "<script>alert('월령은 숫자만 가능합니다. " & startMon & "')</script>"
	end if
	if IsNumeric(endMon) then
		osummarystockbrand.FRectEndDate = endMon
	elseif (endMon <> "") then
		response.write "<script>alert('월령은 숫자만 가능합니다. " & endMon & "')</script>"
	end if

	if (itemgubun = "10") and ((itemidArr<>"") or (itemname<>"") or (makerid<>"") or (cdl<>"") or (mwdiv<>"")) then
		''osummarystockbrand.GetCurrentStockByOnlineBrand
		osummarystockbrand.GetCurrentStockByOnlineBrandNEW_notpaping
	elseif itemgubun <> "10" then
		if itemgubun <> "exc10" then
			osummarystockbrand.FRectItemGubun =  itemgubun
		end if
		osummarystockbrand.GetCurrentStockByOfflineBrand_notpaping
	end if

IsSheetPrintEnable = (osummarystockbrand.FResultCount>0)
if osummarystockbrand.FTotalCount>0 then
    arrlist = osummarystockbrand.farrlist
end if
dim bulkrealstock, buycash
buycash=0
Response.Expires=0
response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=brandcurrentStock_itemlist_" & GetCurrentTimeFormat & ".xls"
Response.CacheControl = "public"
Response.Buffer = true    '버퍼사용여부
%>
<html xmlns:o="urn:schemas-microsoft-com:office:office"
xmlns:x="urn:schemas-microsoft-com:office:excel"
xmlns="http://www.w3.org/TR/REC-html40">

<head>
<meta http-equiv=Content-Type content="text/html; charset=ks_c_5601-1987">
<meta name=ProgId content=Excel.Sheet>
<meta name=Generator content="Microsoft Excel 12">
<style type="text/css">
 td {font-size:8.0pt;}
 .txt {mso-number-format:"\@";}
 .num {mso-number-format:"0";}
 .prc {mso-number-format:"\#\,\#\#0";}
</style>
</head>

<body>
<!--[if !excel]>　　<![endif]-->
<div align=center x:publishsource="Excel">
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="37">
		검색결과 : <b><%= osummarystockbrand.FTotalCount %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td>랙코드</td>
    <td>구분</td>
	<td>상품코드</td>
	<td>옵션<br />코드</td>
	<td>브랜드ID</td>
	<td>상품명<br>[옵션명]</td>
	<td>소비자가</td>
	<td>매입가(현)</td>
	<td>매입<br>구분</td>
	<td>센터<br>매입<br>구분</td>
	<td>총<br>입고<br>반품</td>
	<td>ON총<br>판매<br>반품</td>
    <td>OFF총<br>출고<br>반품</td>
    <td>기타<br>출고<br>반품</td>
    <td>CS<br>출고<br>반품</td>
    <td bgcolor="F4F4F4"><b>시스템<br>총재고</b></td>

	<td>총<br>실사<br>오차</td>
	<td>실사<br>재고</td>
	<td>총<br>불량</td>
	<td bgcolor="F4F4F4"><b>유효<br>재고</b></td>

    <td>총<br>상품<br>준비</td>
    <td bgcolor="F4F4F4"><b>재고<br>파악<br>재고</b></td>
    <td>발주<br>이전<br>주문</td>
    <td bgcolor="F4F4F4">예상<br>재고</td>
    <td width="30">판매<br>여부</td>
    <td width="30">한정<br>여부</td>
    <td>단종<br>여부</td>
	<td width="60">마지막<br>입고월</td>
	<td width="40">전월<br />판매<br />(물류)</td>
    <td width="35">상품<br />등급</td>
    <td>벌크<br />실사</td>
    <td>벌크<br />재고</td>
    <td>AGV<br />재고</td>
</tr>
<% 
if isarray(arrlist) then
for i=0 to ubound(arrlist,2)
if (itemgubun = "10") and ((itemidArr<>"") or (itemname<>"") or (makerid<>"") or (cdl<>"") or (mwdiv<>"")) then
    buycash = arrlist(44,i)
else
    buycash = arrlist(44,i)
    if (buycash = 0) and (arrlist(73,i) <> 0) then
        buycash = CLng(arrlist(43,i) * (100 - arrlist(73,i)) / 100)
    end if
end if
%>
<%
totsysstock	= totsysstock + arrlist(19,i)
totavailstock = totavailstock + arrlist(20,i)
totrealstock = totrealstock + arrlist(21,i)
totjeagosheetstock = totjeagosheetstock + arrlist(21,i) + arrlist(24,i) + arrlist(27,i)
totmaystock = totmaystock + arrlist(21,i) + arrlist(24,i) + arrlist(27,i) + arrlist(25,i) + arrlist(26,i) + arrlist(28,i)

%>
<% if arrlist(45,i)="Y" then %>
<tr bgcolor="#FFFFFF" align="center">
<% else %>
<tr bgcolor="#EEEEEE" align="center">
<% end if %>
    <td><%= arrlist(53,i) %></td>
    <td><%= arrlist(0,i) %></td>
	<td>
	    <% if arrlist(0,i)="10" then %>
	    <%= arrlist(1,i) %>
	    <% else %>
	    <%= arrlist(1,i) %>
	    <% end if %>
	</td>
    <td class="txt"><%= arrlist(2,i) %></td>
	<td><%= arrlist(41,i) %></td>
	<td align="left">
      	<%= arrlist(42,i) %>
      	<% if arrlist(36,i) <>"" then %>
      		<font color="blue">[<%= arrlist(36,i) %>]</font>
      	<% end if %>
    </td>
	<td align="right"><%= FormatNumber(arrlist(60,i),0) %></td>
	<td align="right"><%= FormatNumber(buycash,0) %></td>
    <td><%= fnColor(arrlist(51,i),"mw") %></td>
    <td>
		<%= fnColor(arrlist(61,i),"mw") %>
		<% if IsOffContractExist(arrlist(65,i)) then %>
		<br />
			<% if arrlist(60,i)<>0 then %>
			<%= 100-(CLng(buycash/arrlist(60,i)*10000)/100) %> %
			<% end if %>
			<br>-&gt;<font color="blue"><%= arrlist(65,i) %>%</font>
		<% end if %>
	</td>
	<td align="right"><%= arrlist(5,i) %></td>
	<td align="right"><%= -1*arrlist(13,i) %></td>
	<td align="right"><%= arrlist(6,i) + arrlist(7,i) %></td>
	<td align="right"><%= arrlist(8,i) + arrlist(9,i) %></td>
	<td align="right"><%= arrlist(14,i) %></td>
	<td align="right" bgcolor="F4F4F4"><b><%= arrlist(19,i) %></b></td>

	<td align="right"><%= arrlist(16,i) %></td>
	<td align="right"><%= getErrAssignStock(arrlist(19,i),arrlist(16,i)) %></td>
	<td align="right"><%= arrlist(15,i) %></td>
	<td align="right" bgcolor="F4F4F4"><b><%= arrlist(21,i) %></td>

	<td align="right"><%= arrlist(24,i) + arrlist(27,i) %></td>
	<td align="right" bgcolor="F4F4F4"><b><%= arrlist(21,i) + arrlist(24,i) + arrlist(27,i) %></b></td>
	<td align="right"><%= arrlist(25,i) + arrlist(26,i) + arrlist(28,i) %></td>
	<td align="right" bgcolor="F4F4F4">
		<% if arrlist(48,i)="Y" then %>
			<font color="#FF0000"><%= arrlist(21,i) + arrlist(24,i) + arrlist(27,i) + arrlist(25,i) + arrlist(26,i) + arrlist(28,i) %></font>
		<% else %>
      		<b><%= arrlist(21,i) + arrlist(24,i) + arrlist(27,i) + arrlist(25,i) + arrlist(26,i) + arrlist(28,i) %></b>
    	<% end if %>
    </td>
	<td><%= fnColor(arrlist(47,i),"yn") %></td>
	<td>
		<%= fnColor(arrlist(48,i),"yn") %>
		<% if arrlist(48,i)="Y" then %>
		(<%= GetLimitStr(arrlist(2,i),arrlist(48,i),arrlist(49,i),arrlist(50,i),arrlist(57,i),arrlist(58,i),arrlist(59,i)) %>)
		<% end if %>
	</td>
	<td><%= fnColor(arrlist(52,i),"dj") %></td>
	<td>
		<%= arrlist(62,i) %>
	</td>
	<td>
		<%= arrlist(64,i) %>
	</td>
    <td>
        <% if (arrlist(40,i) = "A") then %><font color="red"><% end if %>
		<%= arrlist(40,i) %>
	</td>
	<td>
        <%
        bulkrealstock = NULL
        if Not IsNull(arrlist(67,i)) and arrlist(67,i) <> "" and IsNumeric(arrlist(67,i)) then
            bulkrealstock = arrlist(67,i) + arrlist(66,i)
        end if
        %>
		<%= bulkrealstock %>
	</td>
    <td>
		<%= arrlist(67,i) %>
	</td>
    <td>
		<%= arrlist(66,i) %>
	</td>
</tr>
<%
if i mod 1000 = 0 then
    Response.Flush		' 버퍼리플래쉬
end if

next
%>
<% else %>
    <tr bgcolor="#FFFFFF">
        <td colspan="37" align="center">[검색결과가 없습니다.]</td>
    </tr>
<% end if %>
</table>
</div>
</body>
</html>
<%
set osummarystockbrand = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
