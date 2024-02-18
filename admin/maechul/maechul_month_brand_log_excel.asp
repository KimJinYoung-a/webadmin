<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 매출로그 Excel 받기
' Hieditor : 2023.06.21 허진원 생성
'###########################################################
%>
<%	'엑셀 출력시작
Response.ContentType = "application/x-msexcel"
Response.CacheControl = "public"
Response.AddHeader "Content-Disposition", "attachment;filename="+"maechul_month_brand_log_excel_"+replace(date,"-","")+".xls"
%>
<!-- #include virtual="/admin/incSessionSTAdmin.asp" -->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/db/dbSTSopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/maechul/pgdatacls.asp"-->
<!-- #include virtual="/lib/classes/maechul/maechulLogCls.asp"-->
<%
dim research
Dim i, yyyy1,mm1,yyyy2,mm2, dd1, dd2, fromDate ,toDate ,oCMaechulLog, page, vatinclude, targetGbn, mwdiv_beasongdiv
dim searchfield, searchtext, dategbn, actDivCode, makerid, excptdlv, exceptSite
dim excTPL
dim searchGbn
dim yyyy3, mm3, yyyy4, mm4, dd3, dd4, fromDate2, toDate2
dim vPurchasetype
dim useNewDB , nxmonthfixed

	research = requestCheckvar(request("research"),10)

	yyyy2   = requestcheckvar(request("yyyy2"),10)
	mm2     = requestcheckvar(request("mm2"),10)
	dd2     = requestcheckvar(request("dd2"),10)
	yyyy4   = requestcheckvar(request("yyyy4"),10)
	mm4     = requestcheckvar(request("mm4"),10)
	dd4     = requestcheckvar(request("dd4"),10)
	vatinclude     = requestcheckvar(request("vatinclude"),1)
	targetGbn     = requestcheckvar(request("targetGbn"),16)
	mwdiv_beasongdiv     = requestcheckvar(request("mwdiv_beasongdiv"),10)
	searchfield 	= request("searchfield")
	searchtext 		= Replace(Replace(request("searchtext"), "'", ""), Chr(34), "")
	dategbn     = requestCheckvar(request("dategbn"),32)
	actDivCode = requestCheckvar(request("actDivCode"),10)
	makerid   = requestcheckvar(request("makerid"),32)
    excptdlv  = requestcheckvar(request("excptdlv"),10)
    exceptSite = requestcheckvar(request("exceptSite"),10)
	searchGbn = requestcheckvar(request("searchGbn"),10)
	vPurchasetype = requestcheckvar(request("purchasetype"),10)

	excTPL 	= request("excTPL")
    useNewDB 	= request("useNewDB")
	nxmonthfixed= request("nxmonthfixed")
    
if dategbn="" then dategbn="ActDate"

if (research = "") then
	excTPL = "Y"
	excptdlv = "on"
	useNewDB = "Y"
end if

if (yyyy2="") then yyyy2 = Cstr(Year( dateadd("m",-1,date()) ))
if (mm2="") then mm2 = Cstr(Month( dateadd("m",-1,date()) ))
if (dd2="") then dd2 = "01"
if (yyyy4="") then yyyy4 = Cstr(Year( dateadd("m",-1,date()) ))
if (mm4="") then mm4 = Cstr(Month( dateadd("m",-1,date()) ))
if (dd4="") then dd4 = "01"

yyyy1=yyyy2
mm1=mm2
dd1=dd2
yyyy3=yyyy4
mm3=mm4
dd3=dd4

fromDate = DateSerial(yyyy2, mm2, dd2)
toDate = DateSerial(yyyy4, mm4, dd4+1)

set oCMaechulLog = new CMaechulLog
	oCMaechulLog.FPageSize = 4000
	oCMaechulLog.FCurrPage = 1
	oCMaechulLog.FRectStartDate = fromDate
	oCMaechulLog.FRectEndDate = toDate
	oCMaechulLog.FRectvatinclude = vatinclude
	oCMaechulLog.FRecttargetGbn = targetGbn
	oCMaechulLog.FRectmwdiv_beasongdiv = mwdiv_beasongdiv
	oCMaechulLog.FRectSearchField = searchfield
	oCMaechulLog.FRectSearchText = searchtext
	oCMaechulLog.FRectDategbn = dategbn
	oCMaechulLog.FRectActDivCode = actDivCode
	oCMaechulLog.FRectmakerid = makerid
	oCMaechulLog.FRectExceptDlv = excptdlv
	oCMaechulLog.FRectExceptSite = exceptSite

	oCMaechulLog.FRectExcTPL = excTPL

	oCMaechulLog.FRectGrpBy = "brand2"
    oCMaechulLog.FRectUseNewDB = useNewDB
	oCMaechulLog.FRectPurchaseType = vPurchasetype
	oCMaechulLog.FRectNextMonthJungsanFixed = nxmonthfixed
	if (oCMaechulLog.FRectPurchaseType = "") then
		oCMaechulLog.FRectPurchaseType = 0
	end if

    oCMaechulLog.GetMaechul_month_item_Log

dim ToTitemno

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
<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">

	<% if (dategbn <> "actDateAndChulgoDate") then %>
	<td rowspan="2">기준월</td>
	<% else %>
	<td rowspan="2">결제월<br>(처리월)</td>
	<td rowspan="2">출고월</td>
	<% end if %>

	<td rowspan="2">매출<br>구분</td>
	<td rowspan="2">브랜드</td>
	<td rowspan="2">(현재)<br />구매유형</td>
	<td rowspan="2">과세구분</td>
	<td rowspan="2">매입<Br>구분</td>
	<td rowspan="2">매출<Br>구분</td>
	<td rowspan="2">판매수량</td>
	<% if (C_InspectorUser = False) then %>
	<td rowspan="2">소비자가<br>합계</td>
	<td rowspan="2">판매가<br>(할인가)</td>
	<td rowspan="2">상품쿠폰<br>적용가</td>
	<td colspan="3">보너스쿠폰</td>
	<td rowspan="2">기타할인<br>(올앳)</td>
	<% end if %>
	<td rowspan="2">매출총액</td>
	<td rowspan="2"><b>공급가액</b></td>
	<td rowspan="2">세액</td>
	<td rowspan="2">업체<Br>정산액</td>
	<td rowspan="2"><b>회계매출</b></td>
	<td rowspan="2">구매<Br>마일리지</td>
	<td rowspan="2">평균<br>매입가</td>
	<td rowspan="2">재고<br>충당금</td>
	<td rowspan="2">비고</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<% if (C_InspectorUser = False) then %>
	<td width="45">비율<br>쿠폰</td>
	<td width="45">정액<br>쿠폰</td>
	<td width="45">배송비<br>쿠폰</td>
	<% end if %>
</tr>
<%
Dim ttl_orgTotalPrice,ttl_subtotalpriceCouponNotApplied, ttl_totalsum
Dim ttl_proCpnDiscount, ttl_totalPriceBonusCouponDiscount, ttl_totalBeasongBonusCouponDiscount, ttl_allatdiscountprice
Dim ttl_totalMaechulPrice,ttl_totalMileage ,ttl_totalBuycash, ttl_totalUpcheJungsanCash
dim ttl_avgipgoPrice, ttl_overValueStockPrice
%>
<% if oCMaechulLog.FresultCount >0 then %>
<% for i=0 to oCMaechulLog.FresultCount -1 %>
<%
ttl_orgTotalPrice=ttl_orgTotalPrice+oCMaechulLog.FItemList(i).forgTotalPrice
ttl_subtotalpriceCouponNotApplied=ttl_subtotalpriceCouponNotApplied+oCMaechulLog.FItemList(i).fsubtotalpriceCouponNotApplied
ttl_totalsum=ttl_totalsum+oCMaechulLog.FItemList(i).ftotalsum

ttl_proCpnDiscount=ttl_proCpnDiscount+(oCMaechulLog.FItemList(i).FtotalBonusCouponDiscount - oCMaechulLog.FItemList(i).FtotalBeasongBonusCouponDiscount)
ttl_totalPriceBonusCouponDiscount=ttl_totalPriceBonusCouponDiscount+oCMaechulLog.FItemList(i).FtotalPriceBonusCouponDiscount
ttl_totalBeasongBonusCouponDiscount=ttl_totalBeasongBonusCouponDiscount+oCMaechulLog.FItemList(i).FtotalBeasongBonusCouponDiscount
ttl_allatdiscountprice=ttl_allatdiscountprice+oCMaechulLog.FItemList(i).fallatdiscountprice

ttl_totalMaechulPrice=ttl_totalMaechulPrice+oCMaechulLog.FItemList(i).ftotalMaechulPrice

ttl_totalMileage=ttl_totalMileage+oCMaechulLog.FItemList(i).ftotalMileage
ttl_totalBuycash=ttl_totalBuycash+oCMaechulLog.FItemList(i).ftotalBuycash
ttl_totalUpcheJungsanCash=ttl_totalUpcheJungsanCash+oCMaechulLog.FItemList(i).ftotalUpcheJungsanCash

ToTitemno = ToTitemno + oCMaechulLog.FItemList(i).Fitemno

ttl_avgipgoPrice = ttl_avgipgoPrice + oCMaechulLog.FItemList(i).FavgipgoPrice
ttl_overValueStockPrice = ttl_overValueStockPrice + CLng(oCMaechulLog.FItemList(i).FoverValueStockPrice)

%>
<tr align="center" bgcolor="FFFFFF" onmouseover=this.style.background="F1F1F1"; onmouseout=this.style.background="FFFFFF";>

	<% if (dategbn <> "actDateAndChulgoDate") then %>
	<td class="txt"><%= oCMaechulLog.FItemList(i).fyyyymm %></td>
	<% else %>
	<td class="txt"><%= oCMaechulLog.FItemList(i).fyyyymm %></td>
	<td class="txt"><%= oCMaechulLog.FItemList(i).fyyyymm2 %></td>
	<% end if %>

	<td class="txt"><%= oCMaechulLog.FItemList(i).FtargetGbn %></td>
	<td class="txt"><%= oCMaechulLog.FItemList(i).Fmakerid %></td>
	<td class="txt"><%= oCMaechulLog.FItemList(i).fpurchasetypename %></td>
	<td class="txt"><%= fnColor(oCMaechulLog.FItemList(i).fvatinclude,"tx") %></td>
	<td class="txt"><%= getmwdiv_beasongdivname(oCMaechulLog.FItemList(i).fmwdiv_beasongdiv) %></td>
	<td class="txt"><%=oCMaechulLog.FItemList(i).getMeaChulGubunName%></td>
	<td align="right" class="prc"><%= FormatNumber(oCMaechulLog.FItemList(i).Fitemno,0) %></td>
	<% if (C_InspectorUser = False) then %>
	<td align="right" class="prc"><%= FormatNumber(oCMaechulLog.FItemList(i).forgTotalPrice, 0) %></td>
	<td align="right" class="prc"><%= FormatNumber(oCMaechulLog.FItemList(i).fsubtotalpriceCouponNotApplied, 0) %></td>
	<td align="right" class="prc"><%= FormatNumber(oCMaechulLog.FItemList(i).ftotalsum, 0) %></td>
	<td align="right" class="prc"><%= FormatNumber((oCMaechulLog.FItemList(i).FtotalBonusCouponDiscount - oCMaechulLog.FItemList(i).FtotalBeasongBonusCouponDiscount), 0) %></td>
	<td align="right" class="prc"><%= FormatNumber(oCMaechulLog.FItemList(i).FtotalPriceBonusCouponDiscount, 0) %></td>
	<td align="right" class="prc"><%= FormatNumber(oCMaechulLog.FItemList(i).FtotalBeasongBonusCouponDiscount, 0) %></td>
	<td align="right" class="prc"><%= FormatNumber(oCMaechulLog.FItemList(i).fallatdiscountprice, 0) %></td>
	<% end if %>
	<td align="right" class="prc"><%= FormatNumber(oCMaechulLog.FItemList(i).ftotalMaechulPrice, 0) %></td>
	<td align="right" class="prc"><%= FormatNumber(oCMaechulLog.FItemList(i).ftotalBuycash, 0) %></td>
	<td align="right" class="prc"><%= FormatNumber(oCMaechulLog.FItemList(i).ftotalMaechulPrice-oCMaechulLog.FItemList(i).ftotalBuycash, 0) %></td>
	<td align="right" class="prc"><%= FormatNumber(oCMaechulLog.FItemList(i).ftotalUpcheJungsanCash, 0) %></td>
	<td align="right" class="prc"><%= FormatNumber((oCMaechulLog.FItemList(i).FtotalMaechulPrice - oCMaechulLog.FItemList(i).FtotalUpcheJungsanCash), 0) %></td>
	<td align="right" class="prc"><%= FormatNumber(oCMaechulLog.FItemList(i).ftotalMileage, 0) %></td>
	<td align="right" class="prc"><%= FormatNumber(oCMaechulLog.FItemList(i).FavgipgoPrice, 0) %></td>
	<td align="right" class="prc"><%= FormatNumber(oCMaechulLog.FItemList(i).FoverValueStockPrice, 0) %></td>
	<td></td>
</tr>
<%
'' ASP 페이지를 실행하여 Response 버퍼의 구성된 제한이 초과되었습니다.
 if (i mod 500)=1 then response.flush
%>
<% next %>
<tr bgcolor="FFFFFF" >

	<% if (dategbn <> "actDateAndChulgoDate") then %>
	<td align="center">합계</td>
	<% else %>
	<td align="center" colspan="2">합계</td>
	<% end if %>

    <td></td>
	<td></td>
	<td></td>
	<td></td>
    <td></td>
    <td></td>
	<td align="right"  class="prc"><%= FormatNumber(ToTitemno,0) %></td>
	<% if (C_InspectorUser = False) then %>
    <td align="right" class="prc"><%=FormatNumber(ttl_orgTotalPrice,0)%></td>
    <td align="right" class="prc"><%=FormatNumber(ttl_subtotalpriceCouponNotApplied,0)%></td>
    <td align="right" class="prc"><%=FormatNumber(ttl_totalsum,0)%></td><!-- 상품쿠폰적용가 -->
    <td align="right" class="prc"><%=FormatNumber(ttl_proCpnDiscount,0)%></td>
    <td align="right" class="prc"><%=FormatNumber(ttl_totalPriceBonusCouponDiscount,0)%></td>
    <td align="right" class="prc"><%=FormatNumber(ttl_totalBeasongBonusCouponDiscount,0)%></td>
    <td align="right" class="prc"><%=FormatNumber(ttl_allatdiscountprice,0)%></td>
	<% end if %>
    <td align="right" class="prc"><%=FormatNumber(ttl_totalMaechulPrice,0)%></td>
    <td align="right" class="prc"><%=FormatNumber(ttl_totalBuycash,0)%></td>
    <td align="right" class="prc"><%=FormatNumber(ttl_totalMaechulPrice-ttl_totalBuycash,0)%></td>
    <td align="right" class="prc"><%=FormatNumber(ttl_totalUpcheJungsanCash,0)%></td>
    <td align="right" class="prc"><%=FormatNumber(ttl_totalMaechulPrice-ttl_totalUpcheJungsanCash,0)%></td>
    <td align="right" class="prc"><%=FormatNumber(ttl_totalMileage,0)%></td>
	<td align="right" class="prc"><%= FormatNumber(ttl_avgipgoPrice, 0) %></td>
	<td align="right" class="prc"><%= FormatNumber(ttl_overValueStockPrice, 0) %></td>
    <td></td>
</tr>
</table>
<% else %>
<tr align="center" bgcolor="#FFFFFF">
	<td colspan="30">검색된 내용이 없습니다.</td>
</tr>
<% end if %>
</body>
</html>
<%
set oCMaechulLog = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
<!-- #include virtual="/lib/db/dbSTSclose.asp" -->