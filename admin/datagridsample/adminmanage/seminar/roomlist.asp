<%@ codepage="65001" language="VBScript" %>
<% option Explicit %>
<% response.Charset="UTF-8" %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/md5.asp"-->
<!-- #include virtual="/lib/util/base64New.asp"-->
<!-- #include virtual="/lib/util/tenEncUtil.asp"-->
<!-- #include virtual="/lib/classes/seminar/seminarCls.asp"-->
<!-- #include virtual="/lib/util/JSON_2.0.4.asp"-->
<script language="jscript" runat="server" src="/lib/util/JSON_PARSER_JS.asp"></script>
<%
'###############################################
' PageName : /datagridsample/adminmanage/seminar/datatojson.asp
' Discription : datagridsample - seminarlist-view
' Response : response > 결과
' History : 2019.10.15 
'###############################################

'//헤더 출력
Response.ContentType = "application/json"
response.charset = "utf-8"

Dim sData : sData = Request("json")
Dim oJson
dim vSeminar , i
dim page :  page = 1
'// 전송결과 파징
'on Error Resume Next

'// json객체 선언
SET oJson = jsArray()
Dim contents_json , contents_object

IF (Err) THEN
	oJson("response") = getErrMsg("9999",sFDesc)
	oJson("faildesc") = "처리중 오류가 발생했습니다.1"
ELSE

    function purpose(v)
        Select Case v
            Case "1" purpose = "[강좌]"
            Case "2" purpose = "[회의]"
            Case "3" purpose = "[미팅]"
            Case "4" purpose = "[면접]"
            Case "5" purpose = "[기타]"
        End Select
    end function 
    
    dim roomNumber , startDate , endDate , groupName , userPurpose , userPhone , userCount , description , userName
    set vSeminar = New CSeminarRoomCalendar
        vSeminar.getReservationList
	
        IF vSeminar.FResultCount > 0 THEN
            FOR i=0 TO vSeminar.FResultCount-1
                roomNumber  = vSeminar.FItemList(i).Froomidx
                startDate   = vSeminar.FItemList(i).Fstart_date
                endDate     = vSeminar.FItemList(i).Fend_date
                groupName   = vSeminar.FItemList(i).Fgroupname
                userPurpose = vSeminar.FItemList(i).Fusepurpose
                userPhone   = vSeminar.FItemList(i).Fusercell
                userCount   = vSeminar.FItemList(i).FuseSu
                description = vSeminar.FItemList(i).Fetc
                userName    = vSeminar.FItemList(i).Fusername

                'UTC - TimeZone 계산
                Set oJson(null) = jsObject()
                    oJson(null)("text")              = purpose(userPurpose) & " "& groupName
                    oJson(null)("username")          = userName&" - "&userCount&"명"
                    oJson(null)("roomId")	         = roomNumber
                    oJson(null)("description")	     = description
                    oJson(null)("startDate")	     = FormatDate(DATEADD("h", -9, startDate),"0000-00-00T00:00Z")
                    oJson(null)("endDate")	         = FormatDate(DATEADD("h", -9, endDate),"0000-00-00T00:00Z")
            Next
        End If 

    set vSeminar = nothing

	'// 결과 출력
	IF (Err) then
		oJson("response") = getErrMsg("9999",sFDesc)
		oJson("faildesc") = "처리중 오류가 발생했습니다.2"
	end if
end if

'Json 출력(JSON)
oJson.flush
Set oJson = Nothing

if ERR then Call OnErrNoti()
'On Error Goto 0
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->