<%@ codepage="65001" language="VBScript" %>
<% option Explicit %>
<%
session.codepage = 65001
response.Charset="UTF-8"
%>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
Server.ScriptTimeOut = 60*10		' 10분
%>
<%
'###########################################################
' Description : 통합알림스케줄
' Hieditor : 2022.12.19 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib_utf8.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function_utf8.asp"-->
<!-- #include virtual="/lib/offshop_function_utf8.asp"-->
<%
dim nIdx, notiType, linkCode, sendType, userId, device, regDate, lastUpdate, replaceItemId, replaceMileage, isusing
dim adminUserid, menupos, mode, sqlStr, notiCnt
	menupos = requestcheckvar(getNumeric(trim(request("menupos"))),10)
	nIdx = requestcheckvar(getNumeric(trim(request("nIdx"))),10)
    notiType=requestcheckvar(trim(request("notiType")),32)
    linkCode=requestcheckvar(getNumeric(trim(request("linkCode"))),10)
    sendType=requestcheckvar(trim(request("sendType")),16)
    userId=requestcheckvar(trim(request("userId")),32)
    device=requestcheckvar(trim(request("device")),1)
    isusing=requestcheckvar(trim(request("isusing")),1)
    mode = RequestCheckVar(request("mode"),32)

adminUserid=session("ssBctId")

if (mode="NotiEdit") or (mode="NotiInsert") then
    if notiType="" or isnull(notiType) then
        response.write "<script type='text/javascript'>"
        response.write "	alert('구분을 선택해 주세요.');"
        response.write "</script>"
        session.codePage = 949
        dbget.close()	:	response.End
    end if
    if linkCode="" or isnull(linkCode) then
        response.write "<script type='text/javascript'>"
        response.write "	alert('관련코드를 등록해 주세요.');"
        response.write "</script>"
        session.codePage = 949
        dbget.close()	:	response.End
    end if
    if sendType="" or isnull(sendType) then
        response.write "<script type='text/javascript'>"
        response.write "	alert('발송구분을 선택해 주세요.');"
        response.write "</script>"
        session.codePage = 949
        dbget.close()	:	response.End
    end if
    if userId="" or isnull(userId) then
        response.write "<script type='text/javascript'>"
        response.write "	alert('고객아이디를 입력해 주세요.');"
        response.write "</script>"
        session.codePage = 949
        dbget.close()	:	response.End
    end if
    if device="" or isnull(device) then
        response.write "<script type='text/javascript'>"
        response.write "	alert('신청채널을 선택해 주세요.');"
        response.write "</script>"
        session.codePage = 949
        dbget.close()	:	response.End
    end if
    if isusing="" or isnull(isusing) then
        response.write "<script type='text/javascript'>"
        response.write "	alert('사용여부를 선택해 주세요.');"
        response.write "</script>"
        session.codePage = 949
        dbget.close()	:	response.End
    end if

end if

If mode = "NotiInsert" then
    notiCnt=0
    sqlstr = "SELECT COUNT(nIdx) as cnt"
    sqlstr = sqlstr & " FROM db_contents.dbo.tbl_IntegrateNotification as n with (nolock)"
    sqlstr = sqlstr & " WHERE isusing='Y'"
    sqlstr = sqlstr & " and userid= '"& userId &"'"
    sqlstr = sqlstr & " and notiType='"& notiType & "'"
    sqlstr = sqlstr & " and sendType='"& sendType & "'"
    sqlstr = sqlstr & " and linkCode="& linkCode & ""

    'response.write sqlstr & "<br>"
    rsget.CursorLocation = adUseClient
    rsget.Open sqlstr, dbget, adOpenForwardOnly, adLockReadOnly
        notiCnt = rsget("cnt")
    rsget.close

    if notiCnt>0 then
        response.write "<script type='text/javascript'>"
        response.write "	alert('이미 중복으로 신청된 내역이 있습니다.\n관련코드와 고객님 아이디로 검색하셔서 확인하세요.');"
        response.write "</script>"
        session.codePage = 949
        dbget.close()	:	response.End
    end if

    sqlStr = "insert into db_contents.dbo.tbl_IntegrateNotification (" & vbcrlf
    sqlStr = sqlStr & " notiType, linkCode, sendType, userId, device, regDate, lastUpdate, replaceItemId, replaceMileage, isusing" & vbcrlf
    sqlStr = sqlStr & " ) values (" & vbcrlf
    sqlStr = sqlStr & " N'"& notiType &"',"& linkCode &",N'"& sendType &"',N'"& userId &"',N'"& device &"',getdate(),getdate()" & vbcrlf
    sqlStr = sqlStr & " ,NULL,NULL,N'"& isusing &"'" & vbcrlf
    sqlStr = sqlStr & " )" & vbcrlf

    'response.write sqlStr & "<Br>"
    dbget.Execute sqlStr

    response.write "<script type='text/javascript'>alert('저장 되었습니다.');</script>"
    session.codePage = 949
    Response.write "<script type='text/javascript'>opener.location.reload();self.close();</script>"
    dbget.close()	:	response.End

elseIf mode = "NotiEdit" then
    if nidx="" or isnull(nidx) then
        response.write "<script type='text/javascript'>"
        response.write "	alert('수정을 위한 구분자가 없습니다.');"
        response.write "</script>"
        session.codePage = 949
        dbget.close()	:	response.End
    end if

    notiCnt=0
    sqlstr = "SELECT COUNT(nIdx) as cnt"
    sqlstr = sqlstr & " FROM db_contents.dbo.tbl_IntegrateNotification as n with (nolock)"
    sqlstr = sqlstr & " WHERE isusing='Y' and nidx<>"& nidx &""
    sqlstr = sqlstr & " and userid= '"& userId &"'"
    sqlstr = sqlstr & " and notiType='"& notiType & "'"
    sqlstr = sqlstr & " and sendType='"& sendType & "'"
    sqlstr = sqlstr & " and linkCode="& linkCode & ""

    'response.write sqlstr & "<br>"
    rsget.CursorLocation = adUseClient
    rsget.Open sqlstr, dbget, adOpenForwardOnly, adLockReadOnly
        notiCnt = rsget("cnt")
    rsget.close

    if notiCnt>0 then
        response.write "<script type='text/javascript'>"
        response.write "	alert('이미 중복으로 신청된 내역이 있습니다.\n관련코드와 고객님 아이디로 검색하셔서 확인하세요.');"
        response.write "</script>"
        session.codePage = 949
        dbget.close()	:	response.End
    end if

    sqlStr="update db_contents.dbo.tbl_IntegrateNotification" & vbcrlf
    sqlStr = sqlStr & " set notiType='"& notiType &"'" & vbcrlf
    sqlStr = sqlStr & " , linkCode="& linkCode &"" & vbcrlf
    sqlStr = sqlStr & " , sendType=N'"& sendType &"'" & vbcrlf
    sqlStr = sqlStr & " , userId=N'"& userId &"'" & vbcrlf
    sqlStr = sqlStr & " , device=N'"& device &"'" & vbcrlf
    sqlStr = sqlStr & " , lastUpdate=getdate()" & vbcrlf
    sqlStr = sqlStr & " , isusing=N'"& isusing &"' where" & vbcrlf
    sqlStr = sqlStr & " nidx="& nidx &""

    'response.write sqlStr & "<Br>"
    dbget.Execute sqlStr

    response.write "<script type='text/javascript'>alert('수정 되었습니다.');</script>"
    session.codePage = 949
    Response.write "<script type='text/javascript'>opener.location.reload();self.close();</script>"
    dbget.close()	:	response.End

else
    response.write "<script type='text/javascript'>alert('정의되지 않았음 "&mode&"');</script>"
    session.codePage = 949
    dbget.close()	:	response.End
end if
%>
<%
session.codePage = 949
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
