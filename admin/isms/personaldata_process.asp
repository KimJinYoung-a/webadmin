<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : �������� ���� �ı� ����
' History : 2019.08.13 �ѿ�� ����
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

' �����ı�
if mode="downFileDelArr" then
    if request.form("idx").count<1 then
        response.write "<script type='text/javascript'>"
        response.write "    alert('���ð��� �����ϴ�.');"
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
    response.write "    alert('������ ��� �Ǿ����ϴ�. Ȯ�μ��� �ۼ��� �ּ���.');"
    response.write "    location.replace('"& refer &"');"
    response.write "</script>"
    dbget.close() : response.end

' Ȯ�μ��ۼ�
elseif mode="downFileconfirmArr" then
    idxarr = trim(idxarr)
    if replace(idxarr,",","")="" then
        response.write "<script type='text/javascript'>"
        response.write "    alert('�����ڰ� �����ϴ�.');"
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
    response.write "    alert('������ �ı� Ȯ�μ��� �ۼ� �Ǿ����ϴ�.');"
    response.write "    opener.location.reload(); self.close();"
    response.write "</script>"
    dbget.close() : response.end

else
    response.write "<script type='text/javascript'>"
    response.write "    alert('�����ڰ� �����ϴ�.');"
    response.write "    location.replace('"& refer &"');"
    response.write "</script>"
    dbget.close() : response.end
end if
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->