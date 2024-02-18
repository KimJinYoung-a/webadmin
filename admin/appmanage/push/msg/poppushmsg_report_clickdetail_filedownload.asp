<%@ codepage="65001" language="VBScript" %>
<% option Explicit %>
<%
session.codepage = 65001
response.Charset="UTF-8"

Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"

Server.ScriptTimeOut = 60
%>
<%
'###########################################################
' Description : 푸시 메시지 클릭 상세 로그 파일다운로드
' Hieditor : 2019.06.26 한용민 생성
'###########################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAppNotiopen.asp" -->
<!-- #include virtual="/lib/function_utf8.asp"-->
<!-- #include virtual="/lib/offshop_function_utf8.asp"-->
<!-- #include virtual="/lib/util/htmllib_utf8.asp" -->
<!-- #include virtual="/lib/classes/appmanage/push/apppush_msg_cls.asp" -->

<%
Dim multipskey,psKey,deviceid,regdate,refIP,appkey,pKey,targetKey,repeatpushyn, userid, menupos, page, arrList
dim cAppPushReport, i, resetyn, reload
	menupos = requestcheckvar(getNumeric(request("menupos")),10)
	page = requestcheckvar(getNumeric(request("page")),10)
    multipskey	= requestcheckvar(getNumeric(request("multipskey")),10)
    deviceid = requestcheckvar(request("deviceid"),512)
    appkey	= requestcheckvar(getNumeric(request("appkey")),1)
    targetKey	= requestcheckvar(getNumeric(request("targetKey")),10)
    repeatpushyn = requestcheckvar(request("repeatpushyn"),1)
    userid = requestcheckvar(request("userid"),32)
	reload = requestcheckvar(request("reload"),2)
	resetyn = requestcheckvar(request("resetyn"),1)

if page = "" then page = 1
if repeatpushyn="" then repeatpushyn="N"

if repeatpushyn="Y" then
	if targetKey="99999" then targetKey=""		' 일반푸시의 타켓전체일경우
end if

if multipskey="" and targetKey="" then
    session.codePage = 949
    response.end
end if

Dim sNow, sY, sM, sD, sH, sMi, sS, sDateName
	sNow = now()
	sY= Year(sNow)
	sM = Format00(2,Month(sNow))
	sD = Format00(2,Day(sNow))
	sH = Format00(2,Hour(sNow))
	sMi = Format00(2,Minute(sNow))
	sS = Format00(2,Second(sNow))
	sDateName = sY&sM&sD&sH&sMi&sS

Response.Expires=0
Response.ContentType = "text/csv"
Response.AddHeader "Content-Type", "text/csv;charset=UTF-8"
Response.AddHeader "Content-Disposition", "attachment; filename=TEN_" & sDateName & "_" & session.sessionID & ".txt"
Response.CacheControl = "public"
Response.Buffer = true    '버퍼사용여부

Dim sqlStr,FTotCnt, FTotPage, FCurrPage, fso, tFile, ArrRows, headLine
Const MaxPage   = 100
Const PageSize = 5000

Function WriteMakeFile(tFile, arrList)
    Dim intLoop,iRow
    Dim bufstr, tmpPrice
    Dim itemid,deliverytype, deliv

    iRow = UBound(arrList,2)
    intLoop=0
    For intLoop=0 to iRow
    	bufstr = ""

        bufstr = "" & Selectpushgubunname(arrList(9,intLoop)) & ""
        if arrList(9,intLoop)="N" or arrList(9,intLoop)="" or isnull(arrList(9,intLoop)) then
		    bufstr = bufstr & "," & arrList(2,intLoop) & ""
        else
            bufstr = bufstr & ",반복푸시("& arrList(8,i) &")"
        end if
        bufstr = bufstr & ",""" & chrbyte(arrList(11,i),20,"N") & """"
		bufstr = bufstr & "," & arrList(10,intLoop) & ""
		bufstr = bufstr & "," & arrList(4,intLoop) & ""
		bufstr = bufstr & "," & arrList(5,intLoop) & ""
		bufstr = bufstr & "," & Selectappname(arrList(6,intLoop)) & ""
        bufstr = bufstr & "," & arrList(3,intLoop) & "" & VbCrlf

        if intLoop mod 5000 = 0 then
            Response.Flush		' 버퍼리플래쉬
        end if
        response.write bufstr
    Next
End function

sqlStr = "exec db_AppNoti.dbo.usp_Ten_PushReport_click_Count '"& multipskey &"','"& deviceid &"','"& appkey &"','"& targetKey &"','"& repeatpushyn &"', '"& userid &"'" & vbcrlf

'response.write sqlStr & "<br>"
rsAppNotiget.CursorType = adOpenStatic
rsAppNotiget.LockType = adLockOptimistic
rsAppNotiget.CursorLocation = adUseClient
rsAppNotiget.Open sqlStr, dbAppNotiget, adOpenForwardOnly, adLockReadOnly
IF Not (rsAppNotiget.EOF OR rsAppNotiget.BOF) THEN
	FTotCnt = rsAppNotiget(0)
END IF
rsAppNotiget.close

IF FTotCnt > 0 THEN
	FTotPage =  CInt(FTotCnt\PageSize)
	If (FTotCnt\PageSize) <> (FTotCnt/PageSize) Then
		FTotPage = FTotPage + 1
	End If
    IF (FTotPage>MaxPage) THEn FTotPage=MaxPage

	headLine = "푸시구분,푸시번호,제목,고객ID,클릭일,IP,OS,디바이스ID" & VbCrlf
    response.write headLine
    i=0
    ' 페이징 하면서 루프 돌며 계속 뿌린다
    For i=0 to FTotPage-1
    	ArrRows = ""

        sqlStr = "exec db_AppNoti.dbo.usp_Ten_PushReport_click_List '"&CStr((PageSize*i) + 1)&"','"&CStr(PageSize*(i+1))&"','"& multipskey &"','"& deviceid &"','"& appkey &"','"& targetKey &"','"& repeatpushyn &"', '"& userid &"'" & vbcrlf

        'response.write sqlStr & "<br>"
        rsAppNotiget.CursorType = adOpenStatic
        rsAppNotiget.LockType = adLockOptimistic
        rsAppNotiget.pagesize = PageSize
        rsAppNotiget.CursorLocation = adUseClient
        rsAppNotiget.Open sqlStr, dbAppNotiget, adOpenForwardOnly, adLockReadOnly
        IF Not (rsAppNotiget.EOF OR rsAppNotiget.BOF) THEN
        	ArrRows = rsAppNotiget.getRows()
        END IF
        rsAppNotiget.close
       	CALL WriteMakeFile(tFile,ArrRows)
    NExt
END IF

session.codePage = 949
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAppNoticlose.asp" -->