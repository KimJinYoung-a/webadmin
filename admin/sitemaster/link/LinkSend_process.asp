<%@ language=vbscript %>
<% option explicit %>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
%>
<%
'#######################################################
'	History	:  2019.10.16 �ѿ�� ����
'	Description : Link �߼�
'#######################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/offshop_function.asp"-->
<%
dim linkidx, title, linkurl, isusing, adminid, menupos, mode, sqlStr
	linkidx = requestCheckVar(getNumeric(request("linkidx")),10)
    title = trim(requestCheckVar(request("title"),128))
    linkurl = trim(requestCheckVar(request("linkurl"),512))
    isusing = requestCheckVar(request("isusing"),1)
    menupos = requestCheckVar(getNumeric(request("menupos")),10)
    mode = requestCheckVar(request("mode"),32)
    adminid=session("ssBctId")

dim refer
    refer = request.ServerVariables("HTTP_REFERER")

if mode="RegLink" then
    if title="" or linkurl="" or isusing="" then
        response.write "<script type='text/javascript'>"
        response.write "    alert('��ũ�� or ������ũ or ��뿩�ο� ���� �����ϴ�.');"
        response.write "    parent.location.replace('"& refer &"')"
        response.write "</script>"
        dbget.close() : response.end
    end if

    if linkidx<>"" then
        if title <> "" and not(isnull(title)) then
            title = ReplaceBracket(title)
        end If
        if linkurl <> "" and not(isnull(linkurl)) then
            linkurl = ReplaceBracket(linkurl)
        end If

        sqlStr = "update db_sitemaster.dbo.tbl_Link_SendList" & vbcrlf
        sqlStr = sqlStr & " set title='"& html2db(title) &"'" & vbcrlf
        sqlStr = sqlStr & " , linkurl='"& html2db(linkurl) &"'" & vbcrlf
        sqlStr = sqlStr & " , isusing='"& isusing &"'" & vbcrlf
        sqlStr = sqlStr & " , lastupdate=getdate()" & vbcrlf
        sqlStr = sqlStr & " , lastadminid='"& adminid &"' where" & vbcrlf
        sqlStr = sqlStr & " linkidx="& linkidx &"" & vbcrlf

        'response.write sqlStr & "<Br>"
        dbget.execute sqlStr

        response.write "<script type='text/javascript'>"
        response.write "    alert('���� �Ǿ����ϴ�.');"
        response.write "    parent.opener.location.reload();"
        response.write "    parent.location.replace('"& refer &"')"
        response.write "</script>"
        dbget.close() : response.end
    else
        if title <> "" and not(isnull(title)) then
            title = ReplaceBracket(title)
        end If
        if linkurl <> "" and not(isnull(linkurl)) then
            linkurl = ReplaceBracket(linkurl)
        end If

        sqlStr="insert into db_sitemaster.dbo.tbl_Link_SendList (" & vbcrlf
        sqlStr = sqlStr & " title, linkurl, isusing, viewcount, regdate, lastupdate, lastadminid) values (" & vbcrlf
        sqlStr = sqlStr & " '"& html2db(title) &"','"& html2db(linkurl) &"','"& isusing &"',0,getdate(),getdate(),'"& adminid &"'" & vbcrlf
        sqlStr = sqlStr & " )" & vbcrlf

        'response.write sqlStr & "<Br>"
        dbget.execute sqlStr

		sqlStr = "select IDENT_CURRENT('db_sitemaster.dbo.tbl_Link_SendList') as linkidx"

        'response.write sqlStr & "<Br>"
        rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		if  not rsget.EOF  then
			linkidx = rsget("linkidx")
		end if
		rsget.Close

        if linkidx="" then
            response.write "<script type='text/javascript'>"
            response.write "    alert('���������� ��ϵ��� �ʾҽ��ϴ�. ������ ���� �ϼ���.');"
            response.write "    parent.location.replace('"& refer &"')"
            response.write "</script>"
            dbget.close() : response.end
        end if

        response.write "<script type='text/javascript'>"
        response.write "    alert('���� �Ǿ����ϴ�.');"
        response.write "    parent.opener.location.reload();"
        response.write "    parent.location.replace('/admin/sitemaster/link/LinkSend_reg.asp?linkidx="& linkidx &"')"
        response.write "</script>"
        dbget.close() : response.end

    end if
else
    response.write "<script type='text/javascript'>"
    response.write "    alert('�������� ��ΰ� �ƴմϴ�.');"
    response.write "    parent.location.replace('"& refer &"')"
    response.write "</script>"
    dbget.close() : response.end
end if
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
