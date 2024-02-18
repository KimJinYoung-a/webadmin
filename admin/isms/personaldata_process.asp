<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : 개인정보 문서 파기 관리
' History : 2019.08.13 한용민 생성
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/isms/personaldata_cls.asp"-->
<%
dim i, mode, sqlStr, userid, menupos, tmpidx, idxarr
    mode = requestcheckvar(request("mode"),32)
    menupos = requestcheckvar(request("menupos"),10)
    idxarr = request("idxarr")

userid = session("ssBctId")

dim refer
    refer = request.ServerVariables("HTTP_REFERER")

' 문서파기
if mode="downFileDelArr" then
    if request.form("idx").count<1 then
        response.write "<script type='text/javascript'>"
        response.write "    alert('선택값이 없습니다.');"
        response.write "    location.replace('"& refer &"');"
        response.write "</script>"
        dbget.close() : response.end
    end if

    for i=1 to request.form("idx").count
        tmpidx = request.form("idx")(i)

        sqlStr = "update db_log.dbo.tbl_ChkAllowIpLog" & vbcrlf
        sqlStr = sqlStr & " set downfiledelyn='Y'," & vbcrlf
        sqlStr = sqlStr & " downFileDelDate=getdate() where" & vbcrlf
        sqlStr = sqlStr & " qryuserid='"& userid &"' and downfiledelyn='N' and idx in ("& tmpidx &")" & vbcrlf

        'response.write sqlStr & "<br>"
        dbget.execute sqlStr
    next

    response.write "<script type='text/javascript'>"
    response.write "    alert('문서가 폐기 되었습니다. 확인서를 작성해 주세요.');"
    response.write "    location.replace('"& refer &"');"
    response.write "</script>"
    dbget.close() : response.end

' 확인서작성
elseif mode="downFileconfirmArr" then
    idxarr = trim(idxarr)
    if replace(idxarr,",","")="" then
        response.write "<script type='text/javascript'>"
        response.write "    alert('구분자가 없습니다.');"
        response.write "    location.replace('"& refer &"');"
        response.write "</script>"
        dbget.close() : response.end
    end if

    if right(idxarr,1)="," then idxarr = left(idxarr,len(idxarr)-1)

    sqlStr = "update db_log.dbo.tbl_ChkAllowIpLog" & vbcrlf
    sqlStr = sqlStr & " set downFileconfirmYN='Y'," & vbcrlf
    sqlStr = sqlStr & " downFileconfirmDelDate=getdate() where" & vbcrlf
    sqlStr = sqlStr & " qryuserid='"& userid &"' and downfiledelyn='Y' and downFileconfirmYN='N' and idx in ("& idxarr &")" & vbcrlf

    'response.write sqlStr & "<br>"
    'response.end
    dbget.execute sqlStr

    response.write "<script type='text/javascript'>"
    response.write "    alert('고객정보 파기 확인서가 작성 되었습니다.');"
    response.write "    opener.location.reload(); self.close();"
    response.write "</script>"
    dbget.close() : response.end

else
    response.write "<script type='text/javascript'>"
    response.write "    alert('구분자가 없습니다.');"
    response.write "    location.replace('"& refer &"');"
    response.write "</script>"
    dbget.close() : response.end
end if
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->