<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"

Server.ScriptTimeOut = 40
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%

Const MaxPage   = 40
Const PageSize = 1000

dim startYYYYMM, endYYYYMM

startYYYYMM = requestCheckvar(request("startYYYYMM"),10)
endYYYYMM = requestCheckvar(request("endYYYYMM"),10)

Dim AdmPath : AdmPath = "/admin/newreport/xldwn/" & Replace(Left(Now(),7), "-", "")
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

Dim FileName: FileName = "preDailyOpExp_"&sDateName&".csv"
dim fso, tFile

Function EscapeString(str)
	dim resultStr : resultStr = str

	resultStr = db2html(resultStr)
	resultStr = Replace(resultStr, "&nbsp;", " ")
	resultStr = Replace(resultStr, Chr(34), "")		'// 이중 따옴표 제거
	resultStr = Chr(34) & resultStr & Chr(34)

	EscapeString = resultStr
end Function

Function WriteMakeFile(tFile, arrList)
    Dim intLoop,iRow
    Dim bufstr, tmpPrice
    Dim itemid,deliverytype, deliv
    iRow = UBound(arrList,2)
    For intLoop=0 to iRow
    	bufstr = ""

		headLine = "청구월,승인일,운영비사용처,수지항목,업체명,사업자번호,적요(상세내역),사용액,공급가액,부가세,봉사료,승인번호,과세유형,국내/외,사용부서,공제여부,처리,useYN"

		bufstr = "'" & arrList(1,intLoop)
		bufstr = bufstr & "," & Left(arrList(2,intLoop),10)
		bufstr = bufstr & "," & EscapeString(arrList(15,intLoop))
		bufstr = bufstr & "," & EscapeString(arrList(5,intLoop))
		bufstr = bufstr & "," & EscapeString(arrList(11,intLoop))
		bufstr = bufstr & "," & arrList(18,intLoop)
		bufstr = bufstr & "," & EscapeString(arrList(12,intLoop))
		bufstr = bufstr & "," & arrList(6,intLoop)
		bufstr = bufstr & "," & arrList(7,intLoop)
		bufstr = bufstr & "," & arrList(8,intLoop)
		bufstr = bufstr & "," & arrList(9,intLoop)
		bufstr = bufstr & "," & arrList(10,intLoop)
		bufstr = bufstr & "," & arrList(16,intLoop)
		bufstr = bufstr & "," & CHKIIF(arrList(19,intLoop)=1,"국내","국외")
		bufstr = bufstr & "," & arrList(14,intLoop)
		bufstr = bufstr & "," & CHKIIF(arrList(17,intLoop)=TRUE,"Y","N")
		bufstr = bufstr & "," & arrList(21,intLoop)
		bufstr = bufstr & "," & CHKIIF(arrList(22,intLoop)=TRUE,"Y","N")

        tFile.WriteLine bufstr
    Next
End function

Dim sqlStr
Dim FTotCnt, FTotPage, FCurrPage
dim otime
''otime=Timer()

sqlStr = "exec [db_partner].[dbo].[sp_Ten_OpExpDailyCard_getNoSetListCnt] '" & startYYYYMM & "','" & endYYYYMM & "',0 ,0 ,0,'','','',0,NULL "
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

	headLine = "청구월,승인일,운영비사용처,수지항목,업체명,사업자번호,적요(상세내역),사용액,공급가액,부가세,봉사료,승인번호,과세유형,국내/외,사용부서,공제여부,처리,useYN"

	tFile.WriteLine headLine

    For i=0 to FTotPage-1
    	ArrRows = ""
        otime=Timer()

		sqlStr =" exec [db_partner].[dbo].sp_Ten_OpExpDailyCard_getNoSetList '" & startYYYYMM & "','" & endYYYYMM & "',0 , 0,0,'','','',0," & (i+1) & "," & PageSize & ",NULL "
		''rw sqlStr
    	rsget.CursorLocation = adUseClient
        rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly', adCmdStoredProc
        IF Not (rsget.EOF OR rsget.BOF) THEN
            ArrRows = rsget.getRows()
        END IF
        rsget.close
        CALL WriteMakeFile(tFile,ArrRows)

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
