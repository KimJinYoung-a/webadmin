<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' [OFF]오프_통계관리>>오프라인 MD별 매출통계
' 백우현 이사님의 지시로 신규 메뉴 생성. (2012-05-10)
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopsellreportmd_cls.asp"-->
<%
dim page,shopid ,yyyymmdd1,yyymmdd2 ,offgubun ,oldlist ,fromDate,toDate ,yyyy1,mm1,dd1,yyyy2,mm2,dd2
dim i, sum1, sum2, sum3 ,makerid ,datefg , parameter ,CurrencyUnit, CurrencyChar, ExchangeRate ,FmNum, vOffCateCode, vOffMDUserID
dim dategubun
	dategubun = requestCheckVar(request("dategubun"),1)
	shopid = requestCheckVar(request("shopid"),32)
	page = requestCheckVar(request("page"),10)
	if page="" then page=1
	yyyy1 = requestCheckVar(request("yyyy1"),4)
	mm1 = requestCheckVar(request("mm1"),2)
	dd1 = requestCheckVar(request("dd1"),2)
	yyyy2 = requestCheckVar(request("yyyy2"),4)
	mm2 = requestCheckVar(request("mm2"),2)
	dd2 = requestCheckVar(request("dd2"),2)
	offgubun = requestCheckVar(request("offgubun"),10)
	oldlist = requestCheckVar(request("oldlist"),2)
	makerid = requestCheckVar(request("makerid"),32)
	datefg = requestCheckVar(request("datefg"),32)
	vOffCateCode = requestCheckVar(request("offcatecode"),32)
	vOffMDUserID = requestCheckVar(request("offmduserid"),32)

if datefg = "" then datefg = "maechul"
if dategubun = "" then dategubun = "G"	
	
sum1 =0
sum2 =0
sum3 =0

if (yyyy1="") then
	fromDate = DateSerial(Cstr(Year(now())), Cstr(Month(now())), Cstr(day(now()))-7)
else
	fromDate = DateSerial(yyyy1, mm1, dd1)
end if

if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
if (dd2="") then dd2 = Cstr(day(now()))

toDate = DateSerial(yyyy2, mm2, dd2+1)

yyyy1 = left(fromDate,4)
mm1 = Mid(fromDate,6,2)
dd1 = Mid(fromDate,9,2)

'/매장
if (C_IS_SHOP) then
	
	'/어드민권한 점장 미만
	if getlevel_sn("",session("ssBctId")) > 6 then
		shopid = C_STREETSHOPID
	end if
else
	'/업체
	if (C_IS_Maker_Upche) then
		makerid = session("ssBctID")
	else
		if (Not C_ADMIN_USER) then
		else
		end if
	end if
end if

if shopid<>"" then offgubun=""

dim ooffsell
set ooffsell = new COffShopSellReportMD
	ooffsell.FRectShopID = shopid
	ooffsell.FRectNormalOnly = "on"
	ooffsell.FRectStartDay = fromDate
	ooffsell.FRectEndDay = toDate
	ooffsell.FRectOffgubun = offgubun
	ooffsell.FRectOldData = oldlist
	ooffsell.frectmakerid = makerid
	ooffsell.frectdatefg = datefg
	ooffsell.frectdategubun = dategubun
	ooffsell.frectoffcatecode = vOffCateCode
	ooffsell.frectoffmduserid = vOffMDUserID
	ooffsell.FCurrPage = page
	ooffsell.Fpagesize=5000
	ooffsell.GetMDSellSumList

'Call fnGetOffCurrencyUnit(shopid,CurrencyUnit, CurrencyChar, ExchangeRate)
'FmNum = CHKIIF(CurrencyUnit="WON" or CurrencyUnit="KRW",0,2)

parameter = "&datefg="& datefg &"&shopid="& shopid &"&offgubun="& offgubun &"&oldlist="& oldlist &"&yyyy1="&yyyy1&"&mm1="&mm1&"&dd1="&dd1&"&yyyy2="&yyyy2&"&mm2="&mm2&"&dd2="&dd2&"offcatecode="&vOffCateCode&"&offmduserid="&vOffMDUserID&"&makerid="&makerid&""

	Response.ContentType = "application/x-msexcel"
	Response.CacheControl = "public"
	Response.AddHeader "Content-Disposition", "attachment;filename=오프라인_MD별_매출통계.xls"
%>

<html>
<head></head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<body>
<table width="100%" border="1" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
	<td>담당자</td>
	<td>브랜드</td>
	<td>매출액</td>
	<td>매입가합</td>
	<td>마진율</td>
	<td>합계매출액</td>
</tr>
<%
	Dim vBody, vTotalTD, v, vTmpMDname, vTmpCnt, vTotalSum
	v = vbCrLf
	vTmpMDname = ""
	vTotalSum = 0
	For i=0 To ooffsell.FresultCount-1
		
		vBody = vBody & "<tr bgcolor=""#FFFFFF"" align=""center"">" & v

		If i = 0 Then
			vTotalSum = vTotalSum + ooffsell.FItemList(i).FSum
		End If

		If vTmpMDname <> ooffsell.FItemList(i).Fmdname Then
			vBody = Replace(vBody,"id='nametd'","rowspan="""&vTmpCnt&"""")
			vBody = Replace(vBody,"id='totaltd'>","rowspan="""&vTmpCnt&""">"&FormatNumber(vTotalSum,0)&"")
			
			vBody = vBody & "	<td id='nametd'>" & ooffsell.FItemList(i).Fmdname & "</td>" & v
			vTotalTD = "	<td id='totaltd'></td>" & v
			vTmpCnt = "1"
			If i <> 0 Then
				vTotalSum = ooffsell.FItemList(i).FSum
			End IF
		Else
			vTotalTD = ""
			vTmpCnt = vTmpCnt + 1
			vTotalSum = vTotalSum + ooffsell.FItemList(i).FSum
		End If
		
		If ooffsell.FItemList(i).FChargeDiv = "6" Then
			vBody = vBody & "	<td><b><font color=""#3333CC"">" & ooffsell.FItemList(i).FMakerid & "</font></b></td>" & v
		Else
			vBody = vBody & "	<td>" & ooffsell.FItemList(i).FMakerid & "</td>" & v
		End If
		
		vBody = vBody & "	<td style=""padding-right:5px;"" align=""right"">" & FormatNumber(ooffsell.FItemList(i).FSum,0) & "</td>" & v
		vBody = vBody & "	<td style=""padding-right:5px;"" align=""right"">" & FormatNumber(ooffsell.FItemList(i).fsuplyprice,0) & "</td>" & v
		vBody = vBody & "	<td style=""padding-right:5px;"" align=""right"">"
		
		If ooffsell.FItemList(i).fsuplyprice > 0 and ooffsell.FItemList(i).FSum > 0 Then
			vBody = vBody & "" & FormatNumber(100-ooffsell.FItemList(i).fsuplyprice/ooffsell.FItemList(i).FSum*100,0) & "%"
		Else
			vBody = vBody & "0%"
		End If
		
		vBody = vBody & "	</td>" & v
		vBody = vBody & vTotalTD
		vBody = vBody & "</tr>" & v
		
		vTmpMDname = ooffsell.FItemList(i).Fmdname
		
		If i = ooffsell.FresultCount-1 Then
			vBody = Replace(vBody,"id='nametd'","rowspan="""&vTmpCnt&"""")
			vBody = Replace(vBody,"id='totaltd'>","rowspan="""&vTmpCnt&""">"&FormatNumber(vTotalSum,0)&"")
		End IF
	Next
	
	Response.Write vBody
%>
</table>


<%
set ooffsell = Nothing
%>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->