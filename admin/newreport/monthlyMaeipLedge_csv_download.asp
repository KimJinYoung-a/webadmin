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
' Description : 재고자산(월별) FIX csv다운로드
' Hieditor : 이상구 생성
'			 2023.10.11 한용민 수정(파일명 변경)
'###########################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%


Const MaxPage   = 50
''Const PageSize = 2000  ''건수 수정..
dim PageSize : PageSize = 2000

dim yyyymm, placeGubun, PriceGbn
dim ver

yyyymm = request("yyyymm")
placeGubun = request("placeGubun")
PriceGbn = request("PriceGbn")
ver = request("ver")

if (ver = "") then
	ver = "V2"
end if

if (ver = "DW") then
    '// 5만개는 서버오류 남.
	PageSize = 2500
end if

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

Dim FileName: FileName = "MonthlyMaeipLedge_"&sDateName&".csv"
dim fso, tFile

Function WriteMakeFile(tFile, arrList, placeGubun)
    Dim intLoop,iRow
    Dim bufstr, tmpPrice
    Dim itemid,deliverytype, deliv
	if isarray(arrList) then
    iRow = UBound(arrList,2)
    For intLoop=0 to iRow
    	bufstr = ""

		bufstr = "'" & arrList(1,intLoop) & "'"
		bufstr = bufstr & "," & trim(arrList(2,intLoop))
		bufstr = bufstr & "," & arrList(3,intLoop)
		bufstr = bufstr & "," & arrList(4,intLoop)
        ''브랜드
        bufstr = bufstr & "," & arrList(26,intLoop)
		''상품구분
		bufstr = bufstr & "," & arrList(5,intLoop)
		''상품코드
		bufstr = bufstr & "," & arrList(6,intLoop)
		''옵션코드
		bufstr = bufstr & "," & "'" & arrList(7,intLoop) & "'"
        ''바코드
		if (ver = "DW") then
            bufstr = bufstr & "," & "'" & arrList(44,intLoop) & "'"
        else
            bufstr = bufstr & "," & "'" & arrList(42,intLoop) & "'"
        end if

		'단가(평균)
		bufstr = bufstr & "," & arrList(28,intLoop) ''arrList(30,intLoop)
		''기초재고(SYS)
		bufstr = bufstr & "," & arrList(8,intLoop)
		bufstr = bufstr & "," & arrList(9,intLoop)
		''입고
		bufstr = bufstr & "," & arrList(10,intLoop)
		bufstr = bufstr & "," & arrList(11,intLoop)
		''이동
		bufstr = bufstr & "," & arrList(12,intLoop)
		bufstr = bufstr & "," & arrList(13,intLoop)
		''판매
		bufstr = bufstr & "," & arrList(14,intLoop)
		bufstr = bufstr & "," & arrList(15,intLoop)
        ''오프출고
		bufstr = bufstr & "," & arrList(16,intLoop)
		bufstr = bufstr & "," & arrList(17,intLoop)
		''기타출고(구:로스출고)
		bufstr = bufstr & "," & arrList(20,intLoop)
		bufstr = bufstr & "," & arrList(21,intLoop)
		''CS출고
		bufstr = bufstr & "," & arrList(22,intLoop)
		bufstr = bufstr & "," & arrList(23,intLoop)

		''오차
		bufstr = bufstr & "," & (arrList(8,intLoop) + arrList(10,intLoop)+ arrList(12,intLoop)+ arrList(14,intLoop)+arrList(16,intLoop)+ arrList(18,intLoop)+arrList(20,intLoop) +arrList(22,intLoop)- arrList(24,intLoop))*-1
		bufstr = bufstr & "," & (arrList(9,intLoop) + arrList(11,intLoop)+ arrList(13,intLoop)+ arrList(15,intLoop)+arrList(17,intLoop)+ arrList(19,intLoop)+arrList(21,intLoop) +arrList(23,intLoop)- arrList(25,intLoop))*-1

		''기말재고(시스템재고)
		bufstr = bufstr & "," & arrList(24,intLoop)
		bufstr = bufstr & "," & arrList(25,intLoop)

		''최종입고월
		if placeGubun <> "S" then
			bufstr = bufstr & ",'" & arrList(29,intLoop) & "'"
		end if

		''최종입고월(매입구분별)
		if placeGubun <> "L" and placeGubun <> "T" and placeGubun <> "O" and placeGubun <> "N" and placeGubun <> "F" and placeGubun <> "A" and placeGubun <> "R" then
			bufstr = bufstr & ",'" & arrList(30,intLoop) & "'"
		end if

		''기타출고처(로스출고처)
		''bufstr = bufstr & "," & Replace(arrList(33,intLoop), ",", "__")

        ''사업자번호
        bufstr = bufstr & "," & arrList(32,intLoop) ''arrList(28,intLoop)

		''재고구분
        bufstr = bufstr & "," & arrList(27,intLoop) ''arrList(29,intLoop)

        ''전시카테고리
        bufstr = bufstr & "," & arrList(31,intLoop) ''arrList(27,intLoop)

        ''관리카테1  //2016/03/22
        bufstr = bufstr & "," & arrList(34,intLoop)
        ''관리카테2
        bufstr = bufstr & "," & arrList(35,intLoop)

		'// 구매유형
		bufstr = bufstr & "," & arrList(36,intLoop)
		bufstr = bufstr & "," & arrList(37,intLoop)
		bufstr = bufstr & "," & arrList(38,intLoop)

		'// 소비자가, 현재판매가, 현재판매여부
		bufstr = bufstr & "," & arrList(39,intLoop)
		bufstr = bufstr & "," & arrList(40,intLoop)
		bufstr = bufstr & "," & arrList(41,intLoop)

		if (ver = "DW") then
			bufstr = bufstr & "," & arrList(42,intLoop)	'취급액(보너스쿠폰적용가)
			bufstr = bufstr & "," & arrList(43,intLoop)	'상품명
			bufstr = bufstr & "," & arrList(47,intLoop)	'옵션명
			bufstr = bufstr & "," & arrList(45,intLoop)	'상품단종여부
			bufstr = bufstr & "," & arrList(46,intLoop)	'옵션단종여부
		end if

        tFile.WriteLine bufstr
    Next
	end if
End function

Dim sqlStr
Dim FTotCnt, FTotPage, FCurrPage
dim otime
''otime=Timer()

if (ver = "V2") then
	sqlStr = "exec [db_summary].[dbo].[sp_Ten_monthlyMaeipLedge_MakeEXL_Count_V2] '" + CStr(yyyymm) + "', '" + CStr(placeGubun) + "' "
elseif (ver = "DW") then
	sqlStr = "exec [db_summary].[dbo].[sp_Ten_monthlyMaeipLedge_MakeEXL_Count_V2] '" + CStr(yyyymm) + "', '" + CStr(placeGubun) + "' "
else
	sqlStr = "exec [db_summary].[dbo].[sp_Ten_monthlyMaeipLedge_MakeEXL_Count] '" + CStr(yyyymm) + "', '" + CStr(placeGubun) + "' "
end if
rsget.CursorLocation = adUseClient
rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly'', adCmdStoredProc
IF Not (rsget.EOF OR rsget.BOF) THEN
	FTotCnt = rsget(0)
END IF
rsget.close

''rw FormatNumber(Timer()-otime,4)
''response.write FTotCnt


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

	headLine = "YYYY-MM,재고위치,부서,과세구분,브랜드,상품구분,상품코드,옵션코드,바코드,단가(평균),기초수량,기초금액,입고수량,입고금액,이동수량,이동금액,판매수량,판매금액,오프출고수량,오프출고금액,기타출고수량,기타출고금액,CS출고수량,CS출고금액,오차수량,오차금액,기말수량,기말금액"

	''최종입고월
	if placeGubun <> "S" then
		headLine = headLine & ",최종입고월"
	end if
	''최종입고월(매입구분별)
	if placeGubun <> "L" and placeGubun <> "T" and placeGubun <> "O" and placeGubun <> "N" and placeGubun <> "F" and placeGubun <> "A" and placeGubun <> "R" then
		headLine = headLine & ",최종입고월(매입구분별)"
	end if

	headLine = headLine & ",사업자번호,재고구분,전시카테고리"


	headLine = headLine & ",관리카테1,관리카테2"
	''headLine = ",,상품구분,상품코드,옵션코드"

	headLine = headLine & ",구매유형"
	headLine = headLine & ",센터매입구분"
	headLine = headLine & ",상품매입구분"
	headLine = headLine & ",소비자가"
	headLine = headLine & ",현재판매가"
	headLine = headLine & ",현재판매여부"
	headLine = headLine & ",보너스쿠폰적용가"
	if (ver = "DW") then
		headLine = headLine & ",상품명"
		headLine = headLine & ",옵션명"
		headLine = headLine & ",상품단종여부"
		headLine = headLine & ",옵션단종여부"
	end if

	tFile.WriteLine headLine

    For i=0 to FTotPage-1
    	ArrRows = ""
        otime=Timer()

		if (ver = "V2") then
		    '' sp_Ten_monthlyMaeipLedge_MakeEXL_List_V2 => sp_Ten_monthlyMaeipLedge_MakeEXL_List_V2_1 ''임시변경 2015/01/12
			sqlStr ="exec [db_summary].[dbo].[sp_Ten_monthlyMaeipLedge_MakeEXL_List_V2_1] '" & yyyymm & "','" & placeGubun & "'," & (i+1) & "," & PageSize & ",'" + CStr(PriceGbn) + "'"  ''위치변경 2015/04/13
		elseif (ver = "DW") then
			sqlStr ="exec [db_datamart].[dbo].[sp_Ten_monthlyMaeipLedge_MakeEXL_List_DW] '" & yyyymm & "','" & placeGubun & "'," & (i+1) & "," & PageSize & ",'" + CStr(PriceGbn) + "'"
		else
			sqlStr ="exec [db_summary].[dbo].[sp_Ten_monthlyMaeipLedge_MakeEXL_List] '" & yyyymm & "','" & placeGubun & "'," & (i+1) & "," & PageSize & ",'" + CStr(PriceGbn) + "'"
		end if

        ''response.write "1111111<br />"
        ''response.write sqlStr
        ''dbget.close : db3_dbget.close : response.end

		if (ver = "DW") then
    		db3_rsget.CursorLocation = adUseClient
			db3_rsget.Open sqlStr, db3_dbget, adOpenForwardOnly, adLockReadOnly', adCmdStoredProc
			IF Not (db3_rsget.EOF OR db3_rsget.BOF) THEN
				ArrRows = db3_rsget.getRows()
			END IF
			db3_rsget.close
		else
    		rsget.CursorLocation = adUseClient
			rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly', adCmdStoredProc
			IF Not (rsget.EOF OR rsget.BOF) THEN
				ArrRows = rsget.getRows()
			END IF
			rsget.close
		end if



        CALL WriteMakeFile(tFile,ArrRows, placeGubun)

        ''rw FormatNumber(Timer()-otime,4)
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
