<%@  codepage="65001" language="VBScript" %>
<% option explicit %>
<% response.Charset="UTF-8" %>
<%
session.codePage = 65001
%>
<%
'###########################################################
' Description : 운송장전송주소오류관리
' Hieditor : 2022.06.27 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db_LogisticsOpen.asp" -->
<!-- #include virtual="/lib/util/htmllib_utf8.asp" -->
<!-- #include virtual="/lib/function_utf8.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/logistics/songjang/SongJangSendClass.asp"-->

<%
Dim i, sqlstr, menupos, idx, SongJangGubun, reqzipcode, reqzipaddr, reqaddress, mode, adminid
dim nm, tel_no, hp_no
	menupos = requestcheckvar(getNumeric(request("menupos")),10)
	mode = requestcheckvar(request("mode"),32)
    idx = requestcheckvar(getNumeric(request("idx")),10)
    SongJangGubun = requestcheckvar(request("SongJangGubun"),10)
	adminid = session("ssBctId")
    reqzipcode = requestcheckvar(trim(request("reqzipcode")),7)
    reqzipaddr = requestcheckvar(trim(request("reqzipaddr")),80)
    reqaddress = requestcheckvar(trim(request("reqaddress")),60)
    nm = requestcheckvar(trim(request("nm")),32)
    tel_no = requestcheckvar(trim(request("tel_no")),16)
    hp_no = requestcheckvar(trim(request("hp_no")),16)

dim refer
	refer = request.ServerVariables("HTTP_REFERER")

if mode="EDIT" then 
    If SongJangGubun = "" OR isnull(SongJangGubun) Then
        Response.Write "<script type='text/javascript'>alert('송장구분이 없습니다.');window.close()</script>"
        session.codePage = 949 : dbget.close() : dbget_Logistics.close() : Response.End
    End IF
    If idx = "" OR isnull(idx) Then
        Response.Write "<script type='text/javascript'>alert('로그번호가 없습니다.');window.close()</script>"
        session.codePage = 949 : dbget.close() : dbget_Logistics.close() : Response.End
    End IF
    If nm = "" OR isnull(nm) Then
        Response.Write "<script type='text/javascript'>alert('이름을 입력하세요.');window.close()</script>"
        session.codePage = 949 : dbget.close() : dbget_Logistics.close() : Response.End
    End IF
    If tel_no = "" OR isnull(tel_no) Then
        Response.Write "<script type='text/javascript'>alert('전화번호를 입력하세요.');window.close()</script>"
        session.codePage = 949 : dbget.close() : dbget_Logistics.close() : Response.End
    End IF
    If hp_no = "" OR isnull(hp_no) Then
        Response.Write "<script type='text/javascript'>alert('휴대폰번호를 입력하세요.');window.close()</script>"
        session.codePage = 949 : dbget.close() : dbget_Logistics.close() : Response.End
    End IF
    If reqzipcode = "" OR isnull(reqzipcode) Then
        Response.Write "<script type='text/javascript'>alert('우편번호를 입력하세요.');window.close()</script>"
        session.codePage = 949 : dbget.close() : dbget_Logistics.close() : Response.End
    End IF
    If reqzipaddr = "" OR isnull(reqzipaddr) Then
        Response.Write "<script type='text/javascript'>alert('주소를 입력하세요.');window.close()</script>"
        session.codePage = 949 : dbget.close() : dbget_Logistics.close() : Response.End
    End IF
    if reqzipcode<>"" then
        if checkNotValidHTML(reqzipcode) then
            Response.Write "<script type='text/javascript'>alert('우편번호에 유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요.');window.close()</script>"
            session.codePage = 949 : dbget.close() : Response.End
        end if
        
        reqzipcode = replace(reqzipcode,"'","""")
    end if
    if reqzipaddr<>"" then
        if checkNotValidHTML(reqzipaddr) then
            Response.Write "<script type='text/javascript'>alert('주소1에 유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요.');window.close()</script>"
            session.codePage = 949 : dbget.close() : Response.End
        end if
        
        reqzipaddr = replace(reqzipaddr,"'","""")
    end if
    if reqaddress<>"" then
        if checkNotValidHTML(reqaddress) then
            Response.Write "<script type='text/javascript'>alert('주소2에 유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요.');window.close()</script>"
            session.codePage = 949 : dbget.close() : Response.End
        end if
        
        reqaddress = replace(reqaddress,"'","""")
    end if

    if SongJangGubun = "GENERAL" then
        sqlstr = "update [db_aLogistics].[dbo].[tbl_Logistics_songjang_log]" & VbCrlf
    elseif SongJangGubun = "RETURN" then
        sqlstr = "update  [db_aLogistics].[dbo].[tbl_Logistics_songjang_log_return]" & VbCrlf
    end if

    sqlstr = sqlstr & " set ZIP_NO=N'"& reqzipcode &"', ADDR=N'"& reqzipaddr &"', ADDR_ETC=N'"& reqaddress &"'," & VbCrlf
    sqlstr = sqlstr & " nm=N'"& nm &"', tel_no=N'"& tel_no &"', hp_no=N'"& hp_no &"'," & VbCrlf
    sqlstr = sqlstr & " ISUPLOADED='N' where" & VbCrlf
    sqlstr = sqlstr & " ISUPLOADED='X' and idx="& idx &""

    'response.write sqlStr & "<br>"
    dbget_Logistics.execute sqlstr

    response.write "<script type='text/javascript'>"
    response.write "	alert('저장되었습니다.');"
    session.codePage = 949
    response.write "	opener.document.location.reload();"
    response.write "	window.close();"
    response.write "</script>"

elseif mode="DEL" then
    If idx = "" OR isnull(idx) Then
        Response.Write "<script type='text/javascript'>alert('로그번호가 없습니다.');window.close()</script>"
        session.codePage = 949 : dbget.close() : dbget_Logistics.close() : Response.End
    End IF

    if SongJangGubun = "GENERAL" then
        sqlstr = "delete from [db_aLogistics].[dbo].[tbl_Logistics_songjang_log] where" & VbCrlf
    elseif SongJangGubun = "RETURN" then
        sqlstr = "delete from [db_aLogistics].[dbo].[tbl_Logistics_songjang_log_return] where" & VbCrlf
    end if

    sqlstr = sqlstr & " idx="& idx &""

    'response.write sqlStr & "<br>"
    dbget_Logistics.execute sqlstr

    response.write "<script type='text/javascript'>"
    response.write "	alert('삭제 되었습니다.');"
    session.codePage = 949
    response.write "	opener.document.location.reload();"
    response.write "	window.close();"
    response.write "</script>"

else
    Response.Write "<script type='text/javascript'>alert('정상적인 경로로 접속해 주세요[999]');window.close()</script>"
    session.codePage = 949 : dbget.close() : Response.End
end if
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db_LogisticsClose.asp" -->
<%
session.codePage = 949
%>