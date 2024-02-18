<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"

Server.ScriptTimeOut = 60*5
%>
<%
'###########################################################
' Description : 제휴몰 클래스
' Hieditor : 2011.04.22 이상구 생성
'			 2020.02.04 한용민 수정
'###########################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%

Const MaxPage   = 40
Const PageSize = 5000  ''건수 수정.. (5000=>1000), 전시카테고리 뺌 => 프로시져 수정

dim makerid, startDate, endDate, vatinclude, mwdiv
dim yyyymm, Dategbn, indexmSqlStr, indexdSqlStr
Dategbn = requestCheckvar(request("Dategbn"),32)
makerid = request("makerid")
startDate = request("startDate")
endDate = request("endDate")
vatinclude = request("vatinclude")
mwdiv = request("mwdiv")

yyyymm = Left(startDate, 7)

if Dategbn="" then Dategbn="chulgoDate"

Dim AdmPath : AdmPath = "/admin/newreport/xldwn/" & Replace(yyyymm, "-", "")
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

Dim FileName: FileName = "maechul_detail_log_"&sDateName&".csv"
dim fso, tFile

function GetActDivCodeName(actDivCode)
	Select Case actDivCode
		Case "A"
			GetActDivCodeName = "원주문"
		Case "C"
			GetActDivCodeName = "취소주문"
		Case "H"
			GetActDivCodeName = "상품변경"
		Case "E"
			GetActDivCodeName = "교환주문"
		Case "M"
			GetActDivCodeName = "반품주문"
		Case "CC"
			GetActDivCodeName = "취소정상화"
		Case "HH"
			GetActDivCodeName = "상품변경취소"
		Case "EE"
			GetActDivCodeName = "교환취소"
		Case "MM"
			GetActDivCodeName = "반품취소"
		Case Else
			GetActDivCodeName = actDivCode
	End Select
end function

function GetFullOrderSerial(orderserial, suborderserial)
	GetFullOrderSerial = orderserial & "-" & Format00(3, suborderserial)
end function

function GetVatIncludeName(vatinclude)
	Select Case vatinclude
		Case "N"
			GetVatIncludeName = "면세"
		Case Else
			GetVatIncludeName = "과세"
	End Select
end function

function GetOMWdivName(omwdiv, itemid)
	if (CStr(itemid) = "0") then
		if (omwdiv="UU") then
			GetOMWdivName = "업배"
		elseif (omwdiv="TT") then
			GetOMWdivName = "텐배"
		else
			GetOMWdivName = omwdiv
		end if
	else
		Select Case omwdiv
			Case "M"
				GetOMWdivName = "매입"
			Case "W"
				GetOMWdivName = "위탁"
			Case "U"
				GetOMWdivName = "업체"

			Case "B000"
				GetOMWdivName = "미지정"
			Case "B011"
				GetOMWdivName = "위탁판매"
			Case "B012"
				GetOMWdivName = "업체위탁"
			Case "B013"
				GetOMWdivName = "출고위탁"
			Case "B021"
				GetOMWdivName = "오프매입"
			Case "B022"
				GetOMWdivName = "매장매입"
			Case "B023"
				GetOMWdivName = "가맹점매입"
			Case "B031"
				GetOMWdivName = "출고매입"
			Case "B032"
				GetOMWdivName = "센터매입"
			Case "B999"
				GetOMWdivName = "기타보정"
			Case "PP"
				GetOMWdivName = "포장"
			Case Else
				GetOMWdivName = omwdiv
		End Select
	end if
end function

Function WriteMakeFile(tFile, arrList)
    Dim intLoop,iRow
    Dim bufstr, tmpPrice
    Dim itemid,deliverytype, deliv, itemname, itemoptionname
    iRow = UBound(arrList,2)
    For intLoop=0 to iRow
    	bufstr = ""

		bufstr = "" & GetActDivCodeName(arrList(0,intLoop)) & ""
		bufstr = bufstr & "," & arrList(1,intLoop)
		bufstr = bufstr & "," & GetFullOrderSerial(arrList(2,intLoop), arrList(3,intLoop))
		bufstr = bufstr & "," & arrList(4,intLoop)
		bufstr = bufstr & "," & Left(arrList(5,intLoop), 10)
		bufstr = bufstr & "," & Left(arrList(6,intLoop), 10)
		bufstr = bufstr & "," & GetVatIncludeName(arrList(7,intLoop))
		bufstr = bufstr & "," & arrList(8,intLoop)

		bufstr = bufstr & "," & GetOMWdivName(arrList(9,intLoop), arrList(11,intLoop))
		bufstr = bufstr & "," & arrList(10,intLoop)
		bufstr = bufstr & "," & arrList(11,intLoop)
		bufstr = bufstr & ",'" & arrList(12,intLoop)

		itemname = db2html(arrList(13,intLoop))
		itemoptionname = db2html(arrList(14,intLoop))
		if (itemoptionname <> "") then
			itemname = itemname & "[" & itemoptionname & "]"
		end if
		itemname = Replace(itemname, Chr(34), "")
		itemname = Chr(34) & itemname & Chr(34)


		bufstr = bufstr & "," & itemname
		bufstr = bufstr & "," & arrList(15,intLoop)
		bufstr = bufstr & "," & arrList(16,intLoop)
		bufstr = bufstr & "," & arrList(17,intLoop)
		bufstr = bufstr & "," & arrList(18,intLoop)
		bufstr = bufstr & "," & arrList(19,intLoop)
		bufstr = bufstr & "," & arrList(20,intLoop)
		bufstr = bufstr & "," & arrList(21,intLoop)
		bufstr = bufstr & "," & arrList(22,intLoop)
		bufstr = bufstr & "," & arrList(23,intLoop)
		bufstr = bufstr & "," & arrList(24,intLoop)
		bufstr = bufstr & "," & arrList(25,intLoop)
		bufstr = bufstr & "," & arrList(26,intLoop)
		bufstr = bufstr & "," & arrList(27,intLoop)
		bufstr = bufstr & "," & arrList(28,intLoop)
		bufstr = bufstr & "," & arrList(29,intLoop)
		bufstr = bufstr & "," & arrList(30,intLoop)

        tFile.WriteLine bufstr
    Next
End function

dim sqlStr, addSqlStr
dim FTotCnt, FTotPage

indexmSqlStr = ""
indexdSqlStr = ""
if Dategbn="ActDate" then
	indexmSqlStr = indexmSqlStr + " with (NOLOCK,index(IX_tbl_order_master_log_actDate))"
elseif Dategbn="chulgoDate" then
	indexdSqlStr = indexdSqlStr + " with (NOLOCK,index(IX_tbl_order_detail_log_beasongdate))"
elseif Dategbn="jFixedDt" then
	indexdSqlStr = indexdSqlStr + " with (NOLOCK)"
else
	indexmSqlStr = indexmSqlStr + " with (NOLOCK,index(IX_tbl_order_master_log_ipkumdate))"
end if
if (application("Svr_Info")="Dev") then indexmSqlStr=" "
if (application("Svr_Info")="Dev") then indexdSqlStr=" "

addSqlStr = ""
' 정산확정일자
if Dategbn="jFixedDt" Then
	addSqlStr = addSqlStr + " and d.DTLjFixedDt>='" + CStr(startDate) + "'"
	addSqlStr = addSqlStr + " and d.DTLjFixedDt<'" + CStr(endDate) + "'"

' 결제일자
elseif Dategbn="ActDate" Then
	addSqlStr = addSqlStr + " and m.actDate>='" + CStr(startDate) + "'"
	addSqlStr = addSqlStr + " and m.actDate<'" + CStr(endDate) + "'"

' 원결제일자
elseif Dategbn="orgPay" Then
	addSqlStr = addSqlStr + " and m.ipkumdate>='" + CStr(startDate) + "'"
	addSqlStr = addSqlStr + " and m.ipkumdate<'" + CStr(endDate) + "'"


' 출고일자
else
	addSqlStr = addSqlStr + " and d.beasongdate>='" + CStr(startDate) + "'"
	addSqlStr = addSqlStr + " and d.beasongdate<'" + CStr(endDate) + "'"
end if

addSqlStr = addSqlStr + " and d.vatinclude='" + vatinclude + "'"
if mwdiv="M" or mwdiv="W" or mwdiv="U" then
	addSqlStr = addSqlStr + " and d.itemid<>0 and d.omwdiv='" + mwdiv + "'"
elseif mwdiv="TT" then
	addSqlStr = addSqlStr + " and d.itemid=0 and left(d.itemoption, 1)<>'9'"
elseif mwdiv="UU" then
	addSqlStr = addSqlStr + " and d.itemid=0 and left(d.itemoption, 1)='9'"
elseif (Len(mwdiv) = 4) then
	addSqlStr = addSqlStr + " and d.omwdiv='" + mwdiv + "' "
end if
addSqlStr = addSqlStr + " and d.makerid='" + makerid + "'"

sqlStr = "select count(*) as cnt , CEILING(CAST(Count(*) AS FLOAT)/" + CStr(PageSize) + ") as totPg "
sqlStr = sqlStr + " from "
sqlStr = sqlStr + "		db_datamart.dbo.tbl_order_master_log m " & indexmSqlStr
sqlStr = sqlStr + "		join db_datamart.dbo.tbl_order_detail_log d " & indexdSqlStr
sqlStr = sqlStr + "		on "
sqlStr = sqlStr + "			1 = 1 "
sqlStr = sqlStr + "			and m.orderserial = d.orderserial "
sqlStr = sqlStr + "			and m.suborderserial = d.suborderserial "
sqlStr = sqlStr + " where "
sqlStr = sqlStr + "		1 = 1 "
sqlStr = sqlStr + addSqlStr

db3_rsget.CursorLocation = adUseClient
db3_rsget.Open sqlStr, db3_dbget, adOpenForwardOnly, adLockReadOnly'', adCmdStoredProc
IF Not (db3_rsget.EOF OR db3_rsget.BOF) THEN
	FTotCnt = db3_rsget(0)
END IF
db3_rsget.close

Dim i, ArrRows
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

	headLine = "구분,매출처,주문번호,원주문번호,원결제일,결제일(처리일),과세구분,상품귀속,매입구분,브랜드,상품코드,옵션코드,상품명[옵션명],수량,소비자가합계,판매가(할인가),상품쿠폰적용가,비율쿠폰,정액쿠폰,배송비쿠폰,기타할인(올앳),매출총액,업체정산액,회계매출,출고일,정산일,구매마일리지,등록일,평균매입가,비고"

	tFile.WriteLine headLine

    For i=0 to FTotPage-1
    	ArrRows = ""

		sqlStr = "select top " + CStr(PageSize*(i+1)) + " "
		sqlStr = sqlStr + " m.actDivCode "
		sqlStr = sqlStr + " , m.sitename "
		sqlStr = sqlStr + " , d.orderserial, d.suborderserial "
		sqlStr = sqlStr + " , (select top 1 linkorderserial from [db_order].[dbo].[tbl_order_master] o where o.orderserial = m.orderserial) as orgorderserial "
		sqlStr = sqlStr + " , m.ipkumdate "
		sqlStr = sqlStr + " , m.actDate "
		sqlStr = sqlStr + " , d.vatinclude "
		sqlStr = sqlStr + " , IsNull(m.targetGbn, 'ON') as targetGbn "

		sqlStr = sqlStr + " , d.omwdiv "
		sqlStr = sqlStr + " , d.makerid "
		sqlStr = sqlStr + " , d.itemid "
		sqlStr = sqlStr + " , d.itemoption "
		sqlStr = sqlStr + " , d.itemname, d.itemoptionname "			'// 14
		sqlStr = sqlStr + " , d.itemno "
		sqlStr = sqlStr + " , d.orgitemcost*d.itemno "
		sqlStr = sqlStr + " , d.itemcostCouponNotApplied*d.itemno "
		sqlStr = sqlStr + " , d.itemcost*d.itemno "

		sqlStr = sqlStr + " , (case when d.itemid <> 0 then (d.itemcost - d.reducedPrice)*d.itemno else 0 end) - d.anbunCouponPriceDetailSUM - allAtDiscount "
		sqlStr = sqlStr + " , d.anbunCouponPriceDetailSUM "
		sqlStr = sqlStr + " , (case when d.itemid = 0 then (d.itemcost - d.reducedPrice)*d.itemno else 0 end) "
		sqlStr = sqlStr + " , d.allAtDiscount "

		sqlStr = sqlStr + " , d.reducedPrice*d.itemno "
		sqlStr = sqlStr + " , d.upcheJungsanCash*d.itemno "
		sqlStr = sqlStr + " , (d.reducedPrice - d.upcheJungsanCash)*d.itemno "
		sqlStr = sqlStr + " , d.beasongdate "
		sqlStr = sqlStr + " , d.DTLjFixedDt"
		sqlStr = sqlStr + " , d.mileage*d.itemno "
		sqlStr = sqlStr + " , '' "
		sqlStr = sqlStr + " , IsNull((case "
		sqlStr = sqlStr + " 	when d.omwdiv in ('M', 'B031') then s.avgipgoPrice*d.itemno "
		sqlStr = sqlStr + "     else 0 end),0) as avgipgoPrice "
		sqlStr = sqlStr + " , '' "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " 	db_datamart.dbo.tbl_order_master_log m "
		sqlStr = sqlStr + " 	join db_datamart.dbo.tbl_order_detail_log d " & indexdSqlStr
		sqlStr = sqlStr + " 	on "
		sqlStr = sqlStr + " 		1 = 1 "
		sqlStr = sqlStr + " 		and m.orderserial = d.orderserial "
		sqlStr = sqlStr + " 		and m.suborderserial = d.suborderserial "
		sqlStr = sqlStr + "		Left Join db_summary.dbo.tbl_monthly_accumulated_logisstock_summary as s with(noLock) "
		sqlStr = sqlStr + "		on s.yyyymm=convert(varchar(7),m.actDate,21) "
		sqlStr = sqlStr + "			and s.itemgubun=d.itemgubun "
		sqlStr = sqlStr + "			and s.itemid=d.itemid "
		sqlStr = sqlStr + "			and s.itemoption=d.itemoption "
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + " 	1 = 1 "
		sqlStr = sqlStr + addSqlStr
    	sqlStr = sqlStr + " order by m.actDate desc, m.orderserial, m.suborderserial, d.itemgubun, d.itemid, d.itemoption"

		'response.write sqlStr & "<Br>"
		db3_rsget.CursorLocation = adUseClient
	    db3_rsget.pagesize = PageSize
		db3_rsget.Open sqlStr,db3_dbget,adOpenForwardOnly, adLockReadOnly
		db3_rsget.absolutepage = (i+1)

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
response.redirect AdmPath&"/"&FileName

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
