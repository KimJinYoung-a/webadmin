<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"

Server.ScriptTimeOut = 60*30
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stockclass/monthlyMaeipLedgeCls.asp"-->
<!-- #include virtual="/lib/BarcodeFunction.asp"-->
<%
Dim CCADMIN : CCADMIN = C_ADMIN_AUTH
dim showItemDetailPopup : showItemDetailPopup = (left(now, 10) <= "2014-08-09") or CCADMIN


dim i, totalpage, totalCount
dim yyyy1, yyyymm1, makerid, showsuply, meaipTp, showShopid, showDiff
dim stockPlace, shopid
dim targetGbn, itemgubun
dim bPriceGbn, showItem
dim stype       '' S:재고, J:정산

dim page : page=1
dim PageSize : PageSize = 1000

stype       = requestCheckvar(request("stype"),10)
shopid    	= requestCheckvar(request("shopid"),32)
yyyy1       = requestCheckvar(request("yyyy1"),10)
stockPlace  = requestCheckvar(request("stockPlace"),10)
makerid     = requestCheckvar(request("makerid"),32)
showsuply   = requestCheckvar(request("showsuply"),10)
showShopid  = requestCheckvar(request("showShopid"),10)
meaipTp     = requestCheckvar(request("meaipTp"),10)
itemgubun   = requestCheckvar(request("itemgubun"),10)
targetGbn   = requestCheckvar(request("targetGbn"),10)
showDiff   	= requestCheckvar(request("showDiff"),10)
bPriceGbn   = requestCheckvar(request("bPriceGbn"),10)
showItem	= requestCheckvar(request("showItem"),2)

if (totalpage="") then totalpage=1
if (stype="") then stype="S"
if (stockPlace="") then stockPlace="L"
if (bPriceGbn="") then bPriceGbn = "V"
if (showsuply="") then showsuply = "on"
if (yyyy1="") then yyyy1 = "2012"

dim stockPlaceName
Select Case stockPlace
	Case "L" : stockPlaceName = "물류"
	Case "S" : stockPlaceName = "매장"
	Case "T" : stockPlaceName = "띵소"
	Case "O" : stockPlaceName = "온라인매입정산"
	Case "N" : stockPlaceName = "온라인매입정산_공제불가"
	Case "F" : stockPlaceName = "오프매입정산"
	Case "A" : stockPlaceName = "핑거스매입정산"
	Case "R" : stockPlaceName = "렌탈"
	Case "E" : stockPlaceName = "에러"
End Select

Dim AdmPath : AdmPath = "/admin/newreport/xldwn/" & year(now)
Dim appPath : appPath = server.mappath(AdmPath) + "/"

Dim FileName: FileName = "YearMaeipLedge_" & yyyy1 & "_" & stockPlaceName & "_" & GetCurrentTimeFormat() &".csv"
dim fso, tFile, headLine, bodyLine
dim FTotCnt : FTotCnt=0

response.write "<div style=""text-align:center;"">"
response.write "<h3 style=""margin:40px 0 20px 0;"">다운로드 파일 <strong>["&FileName&"]</strong> <span id=""statusHead"">생성 중...</span></h3>"
response.write "작성에 시간이 걸립니다. 창을 닫지마시고 다운로드 창이 나타날 때까지 기다려주세요.<br /><br />"
response.write "<div style=""width:400px; padding:1px; background-color:#EEE; margin:auto;""><div id=""statusBar"" style=""width:0px; height:9px; background-image: linear-gradient(to right, #fFf6f0, #E53C3C);""></div></div>"
response.write "<h4 id=""statusText""></h4>"
response.write "</div>"
response.flush

'-----------------------------------------------------------------------
'// 파일헤더 생성
Set fso = CreateObject("Scripting.FileSystemObject")
	If NOT fso.FolderExists(appPath) THEN
		fso.CreateFolder(appPath)
	END If
Set tFile = fso.CreateTextFile(appPath & FileName )

headLine = "년월,매장ID,브랜드ID,상품구분,상품코드,옵션코드,바코드,매입구분,재고위치,구매유형,상품명,옵션명,"
headLine = headLine & "기초재고수량,기초재고금액,매입수량,매입금액,이동수량,이동금액,판매수량,판매금액,매장출고수량,매장출고금액,"
headLine = headLine & "기타출고수량,기타출고금액,로스출고수량,로스출고금액,CS출고수량,CS출고금액,오차수량,오차금액,기말재고수량,기말재고금액"
tFile.WriteLine headLine


'//본내용 생성 및 LOOP
'-----------------------------------------------------------------------
dim oCMonthlyMaeipLedge
	
do until (totalpage < page)
	set oCMonthlyMaeipLedge = new CMonthlyMaeipLedge

	oCMonthlyMaeipLedge.FRectYYYY = yyyy1
	oCMonthlyMaeipLedge.FRectStockPlace = stockPlace
	oCMonthlyMaeipLedge.FRectShopid = shopid
	oCMonthlyMaeipLedge.FRectMakerid = makerid
	oCMonthlyMaeipLedge.FRectBySuplyPrice = CHKIIF(showsuply="on",1,0)
	oCMonthlyMaeipLedge.FRectMeaipTp = meaipTp
	oCMonthlyMaeipLedge.FRectItemgubun = itemgubun
	oCMonthlyMaeipLedge.FRectTargetGbn = targetGbn
	oCMonthlyMaeipLedge.FRectShowShopid = showShopid
	oCMonthlyMaeipLedge.FRectShowItem = showItem
	oCMonthlyMaeipLedge.FRectPriceGubun = bPriceGbn
	oCMonthlyMaeipLedge.FRectShowPurchaseType = "on"	''구매유형표시

	oCMonthlyMaeipLedge.FRectShowDiff = showDiff

	oCMonthlyMaeipLedge.FPageSize = PageSize
	oCMonthlyMaeipLedge.FCurrPage = page

	if (stype="S") then
		oCMonthlyMaeipLedge.GetMaeipLedgeSUMSubDetail
	else
		oCMonthlyMaeipLedge.GetMaeipJungsanSumSubDetail
	end if

	'전체 페이지 반영(loop용)
	if page=1 and totalpage<>oCMonthlyMaeipLedge.FtotalPage then
		totalpage = oCMonthlyMaeipLedge.FtotalPage
		totalCount = oCMonthlyMaeipLedge.FTotalCount
	end if

	if oCMonthlyMaeipLedge.FResultCount>0 then
		FTotCnt = FTotCnt + oCMonthlyMaeipLedge.FResultCount
		for i=0 to oCMonthlyMaeipLedge.FResultCount-1 
			'본문목록 출력
			bodyLine = oCMonthlyMaeipLedge.FItemList(i).Fyyyymm & ","			'대상연도
			bodyLine = bodyLine & chkIIF(showShopid<>"",oCMonthlyMaeipLedge.FItemList(i).Fshopid,"") & ","		'매장ID
			bodyLine = bodyLine & oCMonthlyMaeipLedge.FItemList(i).FMakerid & ","			'브랜드ID
			bodyLine = bodyLine & oCMonthlyMaeipLedge.FItemList(i).Fitemgubun & ","			'상품구분
			bodyLine = bodyLine & oCMonthlyMaeipLedge.FItemList(i).Fitemid & ","			'상품코드
			bodyLine = bodyLine & oCMonthlyMaeipLedge.FItemList(i).Fitemoption & ","		'옵션코드
			bodyLine = bodyLine & BF_MakeTenBarcode(oCMonthlyMaeipLedge.FItemList(i).Fitemgubun, oCMonthlyMaeipLedge.FItemList(i).Fitemid, oCMonthlyMaeipLedge.FItemList(i).Fitemoption) & ","		'바코드
			bodyLine = bodyLine & oCMonthlyMaeipLedge.FItemList(i).getMeaipTypeName & ","	'매입구분
			bodyLine = bodyLine & oCMonthlyMaeipLedge.FItemList(i).FstockPlace & ","		'재고위치
			bodyLine = bodyLine & oCMonthlyMaeipLedge.FItemList(i).FpurchaseTypeName & ","	'구매유형
			bodyLine = bodyLine & oCMonthlyMaeipLedge.FItemList(i).FitemName & ","			'상품명
			bodyLine = bodyLine & oCMonthlyMaeipLedge.FItemList(i).FitemoptionName & ","	'옵션명
			bodyLine = bodyLine & oCMonthlyMaeipLedge.FItemList(i).FprevSysStockNo & ","	'기초재고수량
			bodyLine = bodyLine & oCMonthlyMaeipLedge.FItemList(i).FprevSysStockSum & ","	'기초재고금액
			bodyLine = bodyLine & oCMonthlyMaeipLedge.FItemList(i).getIpgoNo & ","			'매입수량
			bodyLine = bodyLine & oCMonthlyMaeipLedge.FItemList(i).getIpgoSum & ","			'매입금액
			bodyLine = bodyLine & oCMonthlyMaeipLedge.FItemList(i).getMoveNo & ","			'이동수량
			bodyLine = bodyLine & oCMonthlyMaeipLedge.FItemList(i).getMoveSum & ","			'이동금액
			bodyLine = bodyLine & oCMonthlyMaeipLedge.FItemList(i).FSellNo & ","			'판매수량
			bodyLine = bodyLine & oCMonthlyMaeipLedge.FItemList(i).FSellSum & ","			'판매금액
			bodyLine = bodyLine & oCMonthlyMaeipLedge.FItemList(i).FOffChulNo & ","			'매장출고수량
			bodyLine = bodyLine & oCMonthlyMaeipLedge.FItemList(i).FOffChulSum & ","		'매장출고금액
			bodyLine = bodyLine & oCMonthlyMaeipLedge.FItemList(i).FEtcChulNo & ","			'기타출고수량
			bodyLine = bodyLine & oCMonthlyMaeipLedge.FItemList(i).FEtcChulSum & ","		'기타출고금액
			bodyLine = bodyLine & oCMonthlyMaeipLedge.FItemList(i).FLossChulNo & ","		'로스출고수량
			bodyLine = bodyLine & oCMonthlyMaeipLedge.FItemList(i).FLossChulSum & ","		'로스출고금액
			bodyLine = bodyLine & oCMonthlyMaeipLedge.FItemList(i).FCsNo & ","				'CS출고수량
			bodyLine = bodyLine & oCMonthlyMaeipLedge.FItemList(i).FCsSum & ","				'CS출고금액
			bodyLine = bodyLine & oCMonthlyMaeipLedge.FItemList(i).getTotErrNo & ","		'오차수량
			bodyLine = bodyLine & oCMonthlyMaeipLedge.FItemList(i).getTotErrSum & ","		'오차금액
			bodyLine = bodyLine & oCMonthlyMaeipLedge.FItemList(i).FcurSysStockNo & ","		'기말재고수량
			bodyLine = bodyLine & oCMonthlyMaeipLedge.FItemList(i).FcurSysStockSum & ""		'기말재고금액

			tFile.WriteLine bodyLine
		next
	end if
	set oCMonthlyMaeipLedge = Nothing
	page = page + 1

	if totalCount>0 then
		response.write "<script>" &_
				"document.getElementById(""statusText"").innerHTML=""" & formatNumber(totalCount,0) & "건 중 " & formatNumber(FTotCnt,0) & "건 [" & formatNumber(FTotCnt/totalCount*100,1) & "%]""; " &_
				"document.getElementById(""statusBar"").style.width=""" & round(FTotCnt/totalCount*400,0) & "px"";" &_
				"</script>"
		response.flush
	end if
Loop

tFile.Close
Set tFile = Nothing
Set fso = Nothing
response.write "<script>document.getElementById(""statusHead"").innerHTML=""생성 완료!"";</script>"
response.write "<p style=""text-align:center;"">" & FTotCnt&"건 생성 완료 ["&FileName&"]</p>"
%>
<script>
self.location.replace("<%=AdmPath&"/"&FileName%>");
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->
