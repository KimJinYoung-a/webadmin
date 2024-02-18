<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"

Server.ScriptTimeOut = 60*20		' 20분
%>
<%
'###########################################################
' Description : 재고자산
' History : 이상구 생성
'			2023.05.04 한용민 수정(검색조건추가, 재고 계산이 복잡해서 전체셀에 코맨트 넣음)
'###########################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/stockclass/monthlystockcls.asp"-->
<%

Const MaxPage   = 40
Const PageSize = 5000

Const isOnlySys = FALSE
Const isViewWonga =FALSE

dim yyyy1,mm1,isusing,sysorreal, research, newitem, vatyn, minusinc, bPriceGbn, i
dim mwgubun, buseo, itemgubun, stplace, purchasetype, showsuply, dtype, makerid, shopid, etcjungsantype, showDiff
dim brandUseYN
	yyyy1       = requestCheckvar(request("yyyy1"),10)
	mm1         = requestCheckvar(request("mm1"),10)
	isusing     = requestCheckvar(request("isusing"),10)
	sysorreal   = requestCheckvar(request("sysorreal"),10)
	research    = requestCheckvar(request("research"),10)
	newitem     = requestCheckvar(request("newitem"),10)
	mwgubun     = requestCheckvar(request("mwgubun"),10)
	vatyn       = requestCheckvar(request("vatyn"),10)
	minusinc   = requestCheckvar(request("minusinc"),10)
	bPriceGbn   = requestCheckvar(request("bPriceGbn"),10)
	buseo       = requestCheckvar(request("buseo"),10)
	itemgubun   = requestCheckvar(request("itemgubun"),10)
	purchasetype   = requestCheckvar(request("purchasetype"),10)
	stplace     = requestCheckvar(request("stplace"),10)
	showsuply   = requestCheckvar(request("showsuply"),10)
	dtype       = requestCheckvar(request("dtype"),10)
	makerid     = requestCheckvar(request("makerid"),32)
	shopid     = requestCheckvar(request("shopid"),32)
	etcjungsantype      = requestCheckvar(request("etcjungsantype"),10)
	showDiff      = requestCheckvar(request("showDiff"),10)
	brandUseYN      = requestCheckvar(request("brandUseYN"),10)

if (makerid<>"") then dtype=""
if (sysorreal="") then sysorreal="sys"
if (research="") and (bPriceGbn = "") then
    bPriceGbn="V"
end if
if (stplace="") then
    stplace="L"
	showDiff = "Y"
end if
if (research="") then
	if (itemgubun = "") then
		'itemgubun = "AA"
	end if
	if (buseo = "") then
		buseo = "3X"
	end if
end if

dim nowdate
if yyyy1="" then
	nowdate = dateserial(year(Now),month(now)-1,1)
	yyyy1 = Left(CStr(nowdate),4)
	mm1 = Mid(CStr(nowdate),6,2)
end if

dim totno, totbuy, subTotno, subTotbuy '', totavgBuy, offtotavgBuy
dim totPreno, totPrebuy     , subPreno, subPrebuy
dim totIpno,totIpBuy        , subIpno, subIpBuy
dim totLossno, totLossBuy   , subLossno, subLossBuy
dim totSellno, totSellBuy   , subSellno, subSellBuy
dim totOffChulno, totOffChulBuy  , subOffChulno, subOffChulBuy
dim totEtcChulno, totEtcChulBuy  , subEtcChulno, subEtcChulBuy
dim totCsChulno, totCsChulBuy    , subCsChulno, subCsChulBuy
dim iURL, iURLEtc, nBusiName, diffStock, diffStockPrc, diffStockW
DIM isGroupByBrand : isGroupByBrand = (dtype="mk")
Dim isItemList : isItemList = (makerid<>"")

dim totErrBadItemno, totErrBadItemBuy, subErrBadItemno, subErrBadItemBuy
dim totMoveItemno, totMoveItemBuy, subMoveItemno, subMoveItemBuy
dim totErrRealCheckno, totErrRealCheckBuy, subErrRealCheckno, subErrRealCheckBuy
dim totRealStockno, totRealStockBuy, subRealStockno, subRealStockBuy
dim totErrRealCheckBuyPlus, totErrRealCheckBuyMinus

'Dim AdmPath : AdmPath = "/admin/newreport/xldwn/" & Replace(yyyymm, "-", "")
Dim AdmPath : AdmPath = "/admin/newreport/xldwn/" & yyyy1 & mm1
Dim appPath : appPath = server.mappath(AdmPath) + "/"

Dim sNow, sY, sM, sD, sH, sMi, sS, sDateName
	sNow = now()
	sY= Year(sNow)
	sM = Format00(2,Month(sNow))
	sD = Format00(2,Day(sNow))
	sH = Format00(2,Hour(sNow))
	sMi = Format00(2,Minute(sNow))
	sS = Format00(2,Second(sNow))
	sDateName = sY&sM&sD&sH&sMi&sS

Dim FileName: FileName = "MonthlyStockAsset_"&sDateName&".csv"
dim fso, tFile

Function WriteMakeFile(tFile, arrList)
    Dim intLoop,iRow
    Dim bufstr, tmpPrice
    Dim itemid,deliverytype, deliv
    iRow = UBound(arrList,2)
    For intLoop=0 to iRow
    	bufstr = ""

		bufstr = "'" & arrList(1,intLoop) & "'"		' YYYY-MM
		bufstr = bufstr & "," & arrList(2,intLoop)		' 재고위치
		bufstr = bufstr & "," & arrList(3,intLoop)		' 상품구분
		bufstr = bufstr & "," & arrList(4,intLoop)		' 상품코드
		bufstr = bufstr & "," & "'" & arrList(5,intLoop) & "'"		' 옵션코드
        bufstr = bufstr & "," & arrList(44,intLoop)				'바코드

		bufstr = bufstr & "," & arrList(6,intLoop)		' 부서구분

		bufstr = bufstr & "," & arrList(52,intLoop)				'브랜드ID
		bufstr = bufstr & "," & arrList(53,intLoop)				'구매유형

		bufstr = bufstr & "," & arrList(7,intLoop)				'재고매입구분
		bufstr = bufstr & "," & arrList(55,intLoop)				'ON매입구분
		bufstr = bufstr & "," & arrList(8,intLoop)		' 과세구분
		bufstr = bufstr & "," & arrList(10,intLoop)		' 기초재고(SYS)

		if (bPriceGbn = "P") then
			bufstr = bufstr & "," & arrList(11,intLoop)
			bufstr = bufstr & "," & arrList(12,intLoop)		' 매입(입고)
			bufstr = bufstr & "," & arrList(13,intLoop)
			bufstr = bufstr & "," & arrList(56,intLoop)		' 이동
			bufstr = bufstr & "," & arrList(57,intLoop)
			bufstr = bufstr & "," & arrList(14,intLoop)		' 판매
			bufstr = bufstr & "," & arrList(15,intLoop)
			' arrList(16,intLoop) 오프출고수량 arrList(17,intLoop)
			bufstr = bufstr & "," & arrList(16,intLoop)+(arrList(56,intLoop)*-1)		' 당월출고1(오프출고수량+이동수량)
			bufstr = bufstr & "," & arrList(17,intLoop)+(arrList(57,intLoop)*-1)
			bufstr = bufstr & "," & arrList(18,intLoop)-arrList(20,intLoop)		' 당월출고2(물류는 기타출고에 로스가 포함되어 있음)
			bufstr = bufstr & "," & arrList(19,intLoop)-arrList(21,intLoop)
			bufstr = bufstr & "," & arrList(20,intLoop)		' 당월기타출고(로스출고)
			bufstr = bufstr & "," & arrList(21,intLoop)
			bufstr = bufstr & "," & arrList(22,intLoop)		' CS출고
			bufstr = bufstr & "," & arrList(23,intLoop)
			' 오차( (기초재고(SYS)+매입(입고)+이동+판매+당월출고1+당월출고2+CS출고+당월기타출고)-기말재고 )
			bufstr = bufstr & "," & ((arrList(10,intLoop) + arrList(12,intLoop) + arrList(56,intLoop) + arrList(14,intLoop) + (arrList(16,intLoop)+(arrList(56,intLoop)*-1)) + (arrList(18,intLoop)-arrList(20,intLoop)) + arrList(22,intLoop) + arrList(20,intLoop))-arrList(24,intLoop))*-1
			bufstr = bufstr & "," & ((arrList(11,intLoop) + arrList(13,intLoop) + arrList(57,intLoop) + arrList(15,intLoop) + (arrList(17,intLoop)+(arrList(57,intLoop)*-1)) + (arrList(19,intLoop)-arrList(21,intLoop)) + arrList(23,intLoop) + arrList(21,intLoop))-arrList(25,intLoop))*-1
			bufstr = bufstr & "," & arrList(24,intLoop)		' 시스템재고
			bufstr = bufstr & "," & arrList(25,intLoop)
			bufstr = bufstr & "," & arrList(28,intLoop)		' 누적오차
			bufstr = bufstr & "," & arrList(29,intLoop)
			bufstr = bufstr & "," & arrList(30,intLoop)		' 실사재고
			bufstr = bufstr & "," & arrList(31,intLoop)
			bufstr = bufstr & "," & arrList(59,intLoop)		' 누적불량
			bufstr = bufstr & "," & arrList(60,intLoop)
		else
			bufstr = bufstr & "," & arrList(32,intLoop)
			bufstr = bufstr & "," & arrList(12,intLoop)		' 매입(입고)
			bufstr = bufstr & "," & arrList(33,intLoop)
			bufstr = bufstr & "," & arrList(56,intLoop)		' 이동
			bufstr = bufstr & "," & arrList(58,intLoop)
			bufstr = bufstr & "," & arrList(14,intLoop)		' 판매
			bufstr = bufstr & "," & arrList(34,intLoop)
			' arrList(16,intLoop) 오프출고수량 arrList(35,intLoop)
			bufstr = bufstr & "," & arrList(16,intLoop)+(arrList(56,intLoop)*-1)		' 당월출고1(오프출고수량+이동수량)
			bufstr = bufstr & "," & arrList(35,intLoop)+(arrList(58,intLoop)*-1)
			bufstr = bufstr & "," & arrList(18,intLoop)-arrList(20,intLoop)		' 당월출고2(물류는 기타출고에 로스가 포함되어 있음)
			bufstr = bufstr & "," & arrList(36,intLoop)-arrList(37,intLoop)
			bufstr = bufstr & "," & arrList(20,intLoop)		' 당월기타출고(로스출고)
			bufstr = bufstr & "," & arrList(37,intLoop)
			bufstr = bufstr & "," & arrList(22,intLoop)		' CS출고
			bufstr = bufstr & "," & arrList(38,intLoop)
			' 오차( (기초재고(SYS)+매입(입고)+이동+판매+당월출고1+당월출고2+CS출고+당월기타출고)-기말재고 )
			bufstr = bufstr & "," & ((arrList(10,intLoop) + arrList(12,intLoop) + arrList(56,intLoop) + arrList(14,intLoop) + (arrList(16,intLoop)+(arrList(56,intLoop)*-1)) + (arrList(18,intLoop)-arrList(20,intLoop)) + arrList(22,intLoop) + arrList(20,intLoop))-arrList(24,intLoop))*-1
			bufstr = bufstr & "," & ((arrList(32,intLoop) + arrList(33,intLoop) + arrList(58,intLoop) + arrList(34,intLoop) + (arrList(35,intLoop)+(arrList(58,intLoop)*-1)) + (arrList(36,intLoop)-arrList(37,intLoop)) + arrList(38,intLoop) + arrList(37,intLoop))-arrList(39,intLoop))*-1
			bufstr = bufstr & "," & arrList(24,intLoop)		' 시스템재고
			bufstr = bufstr & "," & arrList(39,intLoop)
			bufstr = bufstr & "," & arrList(28,intLoop)		' 누적오차
			bufstr = bufstr & "," & arrList(41,intLoop)
			bufstr = bufstr & "," & arrList(30,intLoop)		' 실사재고
			bufstr = bufstr & "," & arrList(42,intLoop)
			bufstr = bufstr & "," & arrList(59,intLoop)		' 누적불량
			bufstr = bufstr & "," & arrList(61,intLoop)
		end if

		bufstr = bufstr & "," & "'" & arrList(43,intLoop) & "'"		' 최종입고월

        bufstr = bufstr & "," & arrList(45,intLoop)		' 대카테고리코드
        bufstr = bufstr & "," & arrList(46,intLoop)		' 대카테고리명
        bufstr = bufstr & "," & arrList(47,intLoop)		' 중간카테고리코드
        bufstr = bufstr & "," & arrList(48,intLoop)		' 중간카테고리명
        bufstr = bufstr & "," & arrList(49,intLoop)		' 사이즈X
        bufstr = bufstr & "," & arrList(50,intLoop)		' 사이즈Y
        bufstr = bufstr & "," & arrList(51,intLoop)		' 사이즈Z
		bufstr = bufstr & "," & arrList(54,intLoop)		'판매여부

		if (bPriceGbn = "P") then
			bufstr = bufstr & "," & arrList(26,intLoop)		' 당월오차
			bufstr = bufstr & "," & arrList(27,intLoop)
			bufstr = bufstr & "," & arrList(62,intLoop)		' 당월불량
			bufstr = bufstr & "," & arrList(63,intLoop)
		else
			bufstr = bufstr & "," & arrList(26,intLoop)		' 당월오차
			bufstr = bufstr & "," & arrList(40,intLoop)
			bufstr = bufstr & "," & arrList(62,intLoop)		' 당월불량
			bufstr = bufstr & "," & arrList(64,intLoop)
		end if

        tFile.WriteLine bufstr
    Next
End function

Dim sqlStr
Dim FTotCnt, FTotPage, FCurrPage

''rw "수정중"
''response.end

sqlStr = " [db_datamart].[dbo].[sp_Ten_monthlystock_Asset_MakeEXL_Count] ('" & yyyy1 & "-" & mm1 & "','" & stplace & "','" & shopid & "','"&buseo&"','"&itemgubun&"','"&mwgubun&"','"&vatyn&"','"&purchasetype&"','"&CHKIIF(showsuply="on",1,0)&"','"&CHKIIF(dtype="mk",1,0)&"','"&etcjungsantype&"','" & brandUseYN & "','') "

response.write sqlStr & "<br>"
db3_rsget.Open sqlStr, db3_dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
db3_dbget.CommandTimeout = 60*10   ' 10분
IF Not (db3_rsget.EOF OR db3_rsget.BOF) THEN
	FTotCnt = db3_rsget(0)
END IF
db3_rsget.close

response.write "FTotCnt:" & FTotCnt & "<br>"

Dim ArrRows
Dim headLine
IF FTotCnt > 0 THEN
	FTotPage =  CInt(FTotCnt\PageSize)
	If (FTotCnt\PageSize) <> (FTotCnt/PageSize) Then
		FTotPage = FTotPage + 1
	End If
    IF (FTotPage>MaxPage) THEn FTotPage=MaxPage

    Set fso = CreateObject("Scripting.FileSystemObject")
		If NOT fso.FolderExists(appPath) THEN
			fso.CreateFolder(appPath)
		END If
	Set tFile = fso.CreateTextFile(appPath & FileName )

	headLine = "YYYY-MM,재고위치,상품구분,상품코드,옵션코드,바코드,부서구분,브랜드ID,구매유형,재고매입구분,ON매입구분,과세구분,기초재고(SYS),,매입,,이동,,판매,,당월출고1,,당월출고2,,당월기타출고,,CS출고,,오차,,시스템재고,,누적오차,,실사재고,,누적불량,,최종입고월,대카테고리코드,대카테고리명,중간카테고리코드,중간카테고리명,사이즈X,사이즈Y,사이즈Z,판매여부,당월오차,,당월불량,,"

	tFile.WriteLine headLine

    For i=0 to FTotPage-1
    	ArrRows = ""
		sqlStr ="[db_datamart].[dbo].[sp_Ten_monthlystock_Asset_MakeEXL_List] ('" & yyyy1 & "-" & mm1 & "','" & stplace & "','" & shopid & "','"&buseo&"','"&itemgubun&"','"&mwgubun&"','"&vatyn&"','"&purchasetype&"','"&CHKIIF(showsuply="on",1,0)&"','"&CHKIIF(dtype="mk",1,0)&"','"&etcjungsantype&"','" & brandUseYN & "',''," & (i+1) & "," & PageSize & ")"

		response.write sqlStr & "<br>"
		db3_rsget.Open sqlStr, db3_dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
        db3_dbget.CommandTimeout = 60*10   ' 10분
        IF Not (db3_rsget.EOF OR db3_rsget.BOF) THEN
        	ArrRows = db3_rsget.getRows()
        END IF
        db3_rsget.close
       	CALL WriteMakeFile(tFile,ArrRows)
    NExt
    tFile.Close
	Set tFile = Nothing
	Set fso = Nothing
END IF

response.write FTotCnt&"건 생성 ["&FileName&"]"
IF FTotCnt > 0 THEN
    response.redirect AdmPath&"/"&FileName
end if
''response.end
''response.write appPath & FileName
%>
<!-- #include virtual="/lib/db/db3close.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
