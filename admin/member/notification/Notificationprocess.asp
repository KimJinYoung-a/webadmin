<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : �˸�����
' Hieditor : 2023.03.30 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/admin/MemberCls.asp" -->

<%
dim userId, notificationType, idx, menupos, mode, i, existscount, adminid, sqlStr, existsuseridYN
	userId = requestcheckvar(request("userId"),32)
    notificationType = requestcheckvar(request("notificationType"),32)
    idx = requestCheckvar(getNumeric(Request("idx")),10)
    menupos = requestCheckvar(getNumeric(Request("menupos")),10)
    mode = requestcheckvar(request("mode"),32)

dim ref
	ref = request.ServerVariables("HTTP_REFERER")

existscount=0
adminid		= session("ssBctId")
existsuseridYN="N"

sqlStr = "select count(ut.userid) as existscount"
sqlStr = sqlStr + " from db_partner.dbo.tbl_user_tenbyten ut with (nolock)"
sqlStr = sqlStr + " where ut.userid='"& userid &"'"

'response.write sqlStr & "<br>"
rsget.CursorLocation = adUseClient
rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
if Not rsget.Eof then
    if rsget("existscount")>0 then existsuseridYN="Y"
end if
rsget.Close

if existsuseridYN="N" then
    response.write "<script type='text/javascript'>alert('�������� �ʴ� ���� ���̵� �Դϴ�.');history.back();</script>"
    dbget.close()	:	response.End
end if

if mode = "NotificationReg" then
    sqlStr = "select count(nu.idx) as existscount"
    sqlStr = sqlStr + " from db_partner.dbo.notificationUser nu with (nolock)"
    sqlStr = sqlStr + " where nu.isusing='Y' and nu.userid='"& userid &"' and notificationType='"& notificationType &"'"

    'response.write sqlStr & "<br>"
    rsget.CursorLocation = adUseClient
    rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
    if Not rsget.Eof then
        existscount = rsget("existscount")
    end if
    rsget.Close

    if existscount>0 then
        response.write "<script type='text/javascript'>alert('�ش� �˸��� �̹� ��ϵǾ� �ֽ��ϴ�.');history.back();</script>"
        dbget.close()	:	response.End
    end if
	
    sqlStr = "insert into db_partner.dbo.notificationUser(" & vbcrlf
    sqlStr = sqlStr & " userid,notificationType,isusing,regdate,lastupdate,reguserid,lastuserid" & vbcrlf
    sqlStr = sqlStr & " ) values (" & vbcrlf
    sqlStr = sqlStr & " '" & userid & "', '" & notificationType & "', 'Y', getdate(), getdate(), '" & adminid & "', '" & adminid & "'" & vbcrlf
    sqlStr = sqlStr & " )"

    'response.write sqlStr & "<br>"
    'response.end
    dbget.Execute sqlStr

    response.write "<script type='text/javascript'>alert('���� �Ǿ����ϴ�.');</script>"
    response.write "<script type='text/javascript'>opener.location.reload(); opener.focus(); location.replace('/admin/member/notification/NotificationUser.asp?userId="&userid&"&menupos="&menupos&"');</script>"
    dbget.close()	:	response.End

'//����
elseif mode = "NotificationUserDel" then
    if idx="" or isnull(idx) then
        response.write "<script type='text/javascript'>alert('��뿩�ΰ� �����ϴ�');history.back();</script>"
        dbget.close()	:	response.End
    end if

    sqlStr = "update db_partner.dbo.notificationUser set isusing='N',lastupdate=getdate(),lastuserid='"&adminid&"' where isusing='Y' and idx = "& idx &""

    'response.write sqlStr &"<Br>"
    dbget.execute sqlStr

    response.write "<script type='text/javascript'>alert('���� �Ǿ����ϴ�.');</script>"
    response.write "<script type='text/javascript'>opener.location.reload(); opener.focus(); location.replace('/admin/member/notification/NotificationUser.asp?userId="&userid&"&menupos="&menupos&"');</script>"
    dbget.close()	:	response.End

end if
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->