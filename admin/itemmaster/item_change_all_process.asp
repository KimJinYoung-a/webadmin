<%@ language=vbscript %>
<% option explicit %>
<%
session.codePage = 949
Response.CharSet = "EUC-KR"
%>
<%
'###########################################################
' Description : ��ǰ�ϰ�����[������]
' History : 2021.11.15 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/itemcls_2008.asp"-->
<%
dim menupos, makerid, mode, mwdiv, sqlstr, deliverytype, resultrows, margin, vSCMChangeSQL, vChangeContents, toMakerid, toBrandName
    menupos = requestCheckvar(getNumeric(trim(request("menupos"))),10)
    mode = requestCheckvar(trim(request("mode")),32)
    makerid = requestCheckvar(trim(request("makerid")),32)
    mwdiv = requestCheckvar(trim(request("mwdiv")),1)
    margin = requestCheckvar(trim(request("margin")),10)
    toMakerid = requestCheckvar(trim(request("toMakerid")),32)

vChangeContents = "- HTTP_REFERER : " & request.ServerVariables("HTTP_REFERER") & vbCrLf
resultrows=0
if mwdiv="M" or mwdiv="W" then
    deliverytype="1"
else
    deliverytype="2"
end if

if mode="makerchmwdiv" then
    if not(C_ADMIN_AUTH or C_MD_AUTH or C_SYSTEM_Part) then
        response.write "<script type='text/javascript'>alert('������ �����ϴ�. MD��,������ ��Ʈ�� �̻� ���� ���� �մϴ�.[0]');</script>"
        dbget.close() : response.end
    end if
    if makerid="" or isnull(makerid) then
        response.write "<script type='text/javascript'>alert('�귣��ID�� �����ϴ�.[0]');</script>"
        dbget.close() : response.end
    end if
    if mwdiv="" or isnull(mwdiv) then
        response.write "<script type='text/javascript'>alert('�����Ͻ� �귣��ID�� ������ �ּ���.[0]');</script>"
        dbget.close() : response.end
    end if
    if deliverytype="" or isnull(deliverytype) then
        response.write "<script type='text/javascript'>alert('��۱����� �����ϴ�.[0]');</script>"
        dbget.close() : response.end
    end if

    sqlstr = "update db_item.dbo.tbl_item" & vbcrlf
    sqlstr = sqlstr & " set lastupdate=getdate()" & vbcrlf
    sqlstr = sqlstr & " , isextusing='N'" & vbcrlf
    sqlstr = sqlstr & " ,mwdiv='"& ucase(mwdiv) &"'" & vbcrlf
    sqlstr = sqlstr & " ,deliverytype='"& deliverytype &"' where" & vbcrlf
    sqlstr = sqlstr & " makerid in ('"& makerid &"')" & vbcrlf
    sqlstr = sqlstr & " and itemdiv not in (21,23)"	    ' ��, B2B����

    'response.write sqlstr & "<Br>"
    dbget.execute sqlstr, resultrows

    sqlstr = "update db_temp.dbo.tbl_wait_item" & vbcrlf
    sqlstr = sqlstr & " set isextusing='N'" & vbcrlf
    sqlstr = sqlstr & " , mwdiv='"& ucase(mwdiv) &"'" & vbcrlf
    sqlstr = sqlstr & " , deliverytype='"& deliverytype &"' where" & vbcrlf
    sqlstr = sqlstr & " makerid in ('"& makerid &"')" & vbcrlf
    sqlstr = sqlstr & " and itemdiv not in (21,23)"	    ' ��, B2B����

    'response.write sqlstr & "<Br>"
    dbget.execute sqlstr

    vChangeContents = vChangeContents & "- ��౸�� : mwdiv = " & ucase(mwdiv) & vbCrLf
    vChangeContents = vChangeContents & "- ��۱��� : deliverytype = " & deliverytype & vbCrLf
    vChangeContents = vChangeContents & "- ���޸��Ǹſ��� : isextusing = N" & vbCrLf

    '### ���� �α� ����(item)
    vSCMChangeSQL = "INSERT INTO [db_log].[dbo].[tbl_scm_change_log] (userid, gubun, pk_idx, menupos, contents, refip)"
    vSCMChangeSQL = vSCMChangeSQL & "   select"
    vSCMChangeSQL = vSCMChangeSQL & "   '" & session("ssBctId") & "', 'item', i.itemid, '" & Request("menupos") & "', '" & vChangeContents & "'"
    vSCMChangeSQL = vSCMChangeSQL & "   , '" & Request.ServerVariables("REMOTE_ADDR") & "'"
    vSCMChangeSQL = vSCMChangeSQL & "   from db_item.dbo.tbl_item i with (nolock)"
    vSCMChangeSQL = vSCMChangeSQL & "   where makerid in ('"& makerid &"')"
    vSCMChangeSQL = vSCMChangeSQL & "   and itemdiv not in (21,23)"	    ' ��, B2B����

    'response.write vSCMChangeSQL & "<Br>"
    dbget.execute vSCMChangeSQL

    response.write "<script type='text/javascript'>"
    response.write "    alert('��౸�к��� "& resultrows &"�� ó���Ǿ����ϴ�.');"
    response.write "    parent.location.reload();"
    response.write "</script>"
    dbget.close() : response.end

elseif mode="makerchmargin" then
    if not(C_ADMIN_AUTH or C_MD_AUTH or C_SYSTEM_Part) then
        response.write "<script type='text/javascript'>alert('������ �����ϴ�. MD��,������ ��Ʈ�� �̻� ���� ���� �մϴ�.[1]');</script>"
        dbget.close() : response.end
    end if
    if makerid="" or isnull(makerid) then
        response.write "<script type='text/javascript'>alert('�귣��ID�� �����ϴ�.[1]');</script>"
        dbget.close() : response.end
    end if
    if margin="" or isnull(margin) or margin="0" then
        response.write "<script type='text/javascript'>alert('�����Ͻ� ���� % �� �Է��� �ּ���.[1]');</script>"
        dbget.close() : response.end
    end if

    sqlstr = "update o" & vbcrlf
    sqlstr = sqlstr & " set o.optaddbuyprice=convert(int,o.optaddprice-(convert(int,o.optaddprice*"& margin &")/100))" & vbcrlf	' �ɼ��߰��ݾ� ���԰�
    sqlstr = sqlstr & " from db_item.dbo.tbl_item as i with(noLock)" & vbcrlf
    sqlstr = sqlstr & " join db_item.dbo.tbl_item_option as o with(noLock)" & vbcrlf
    sqlstr = sqlstr & "     on i.itemid=o.itemid" & vbcrlf
    sqlstr = sqlstr & " where i.makerid in ('"& makerid &"')" & vbcrlf
    sqlstr = sqlstr & " and o.optaddprice>0" & vbcrlf	' �ɼ� �߰��ݾ��� �ִ� ���

    'response.write sqlstr & "<Br>"
    dbget.execute sqlstr

    sqlstr = "update db_item.dbo.tbl_item" & vbcrlf
    sqlstr = sqlstr & " set lastupdate=getdate()" & vbcrlf
    sqlstr = sqlstr & " , isextusing='N'" & vbcrlf
    sqlstr = sqlstr & " , orgsuplycash = convert(int,orgprice-(convert(int,orgprice*"& margin &")/100))" & vbcrlf
    sqlstr = sqlstr & " , buycash = convert(int,sellcash-(convert(int,sellcash*"& margin &")/100))" & vbcrlf
    sqlstr = sqlstr & " , sailsuplycash = iif(sailyn='Y',convert(int,sailprice-(convert(int,sailprice*"& margin &")/100)),sailsuplycash) where" & vbcrlf
    sqlstr = sqlstr & " makerid in ('"& makerid &"')" & vbcrlf
    sqlstr = sqlstr & " and itemdiv not in (21,23)"	    ' ��, B2B����

    'response.write sqlstr & "<Br>"
    dbget.execute sqlstr, resultrows

    sqlstr = "update db_temp.dbo.tbl_wait_item" & vbcrlf
    sqlstr = sqlstr & " set isextusing='N'" & vbcrlf
    sqlstr = sqlstr & " , buycash = convert(int,sellcash-(convert(int,sellcash*"& margin &")/100)) where" & vbcrlf
    sqlstr = sqlstr & " makerid in ('"& makerid &"')" & vbcrlf
    sqlstr = sqlstr & " and itemdiv not in (21,23)"	    ' ��, B2B����

    'response.write sqlstr & "<Br>"
    dbget.execute sqlstr

    vChangeContents = vChangeContents & "- ���޸��Ǹſ��� : isextusing = N" & vbCrLf

    '### ���� �α� ����(item)
    vSCMChangeSQL = "INSERT INTO [db_log].[dbo].[tbl_scm_change_log] (userid, gubun, pk_idx, menupos, contents, refip)"
    vSCMChangeSQL = vSCMChangeSQL & "   select"
    vSCMChangeSQL = vSCMChangeSQL & "   '" & session("ssBctId") & "', 'item', i.itemid, '" & Request("menupos") & "'"
    vSCMChangeSQL = vSCMChangeSQL & "   , '- �Һ��ڸ��԰� : orgsuplycash = ' + convert(nvarchar,convert(int,orgprice-(convert(int,orgprice*"& margin &")/100))) + '" & vbcrlf & "- �ǸŸ��԰� : buycash = ' + convert(nvarchar,convert(int,orgprice-(convert(int,orgprice*"& margin &")/100))) + '" & vbcrlf & vChangeContents & "'"
    vSCMChangeSQL = vSCMChangeSQL & "   , '" & Request.ServerVariables("REMOTE_ADDR") & "'"
    vSCMChangeSQL = vSCMChangeSQL & "   from db_item.dbo.tbl_item i with (nolock)"
    vSCMChangeSQL = vSCMChangeSQL & "   where makerid in ('"& makerid &"')"
    vSCMChangeSQL = vSCMChangeSQL & "   and itemdiv not in (21,23)"	    ' ��, B2B����

    'response.write vSCMChangeSQL & "<Br>"
    dbget.execute vSCMChangeSQL

    response.write "<script type='text/javascript'>"
    response.write "    alert('�������� "& resultrows &"�� ó���Ǿ����ϴ�.');"
    response.write "    parent.location.reload();"
    response.write "</script>"
    dbget.close() : response.end

elseif mode="makerchsellyn_n" then
    if not(C_ADMIN_AUTH or C_MD_AUTH or C_SYSTEM_Part) then
        response.write "<script type='text/javascript'>alert('������ �����ϴ�. MD��,������ ��Ʈ�� �̻� ���� ���� �մϴ�.[1]');</script>"
        dbget.close() : response.end
    end if
    if makerid="" or isnull(makerid) then
        response.write "<script type='text/javascript'>alert('�귣��ID�� �����ϴ�.[1]');</script>"
        dbget.close() : response.end
    end if

    vChangeContents = vChangeContents & "- �Ǹſ��� : sellyn = N" & vbCrLf

    '### ���� �α� ����(item)
    vSCMChangeSQL = "INSERT INTO [db_log].[dbo].[tbl_scm_change_log] (userid, gubun, pk_idx, menupos, contents, refip)"
    vSCMChangeSQL = vSCMChangeSQL & "   select"
    vSCMChangeSQL = vSCMChangeSQL & "   '" & session("ssBctId") & "', 'item', i.itemid, '" & Request("menupos") & "'"
    vSCMChangeSQL = vSCMChangeSQL & "   , '"& vChangeContents &"'"
    vSCMChangeSQL = vSCMChangeSQL & "   , '" & Request.ServerVariables("REMOTE_ADDR") & "'"
    vSCMChangeSQL = vSCMChangeSQL & "   from db_item.dbo.tbl_item i with (nolock)"
    vSCMChangeSQL = vSCMChangeSQL & "   where i.sellyn='Y' and i.makerid='"& makerid &"'"

    'response.write vSCMChangeSQL & "<Br>"
    dbget.execute vSCMChangeSQL

    sqlstr = "update db_item.dbo.tbl_item" & vbcrlf
    sqlstr = sqlstr & " set sellyn='N', isextusing='N'," & vbcrlf
    sqlstr = sqlstr & " lastupdate=GETDATE() where" & vbcrlf
    sqlstr = sqlstr & " sellyn='Y' and makerid='"& makerid &"'"

    'response.write sqlstr & "<Br>"
    dbget.execute sqlstr, resultrows

    sqlstr = "update db_temp.dbo.tbl_wait_item" & vbcrlf
    sqlstr = sqlstr & " set sellyn='N'" & vbcrlf
    sqlstr = sqlstr & " , isextusing='N'" & vbcrlf
    sqlstr = sqlstr & " , lastupdate=getdate() where" & vbcrlf
    sqlstr = sqlstr & " makerid='blueelephant10'"

    'response.write sqlstr & "<Br>"
    dbget.execute sqlstr

    response.write "<script type='text/javascript'>"
    response.write "    alert('�Ǹž������� "& resultrows &"�� ó���Ǿ����ϴ�.');"
    response.write "    parent.location.reload();"
    response.write "</script>"
    dbget.close() : response.end

elseif mode="MoveMaker" then
    if not(C_ADMIN_AUTH or C_MD_AUTH or C_SYSTEM_Part) then
        response.write "<script type='text/javascript'>alert('������ �����ϴ�. MD��,������ ��Ʈ�� �̻� ���� ���� �մϴ�.[1]');</script>"
        dbget.close() : response.end
    end if
    if makerid="" or isnull(makerid) then
        response.write "<script type='text/javascript'>alert('��� �귣��ID�� �����ϴ�.[1]');</script>"
        dbget.close() : response.end
    end if
    if toMakerid="" or isnull(toMakerid) then
        response.write "<script type='text/javascript'>alert('�̵��� �귣��ID�� �����ϴ�.[2]');</script>"
        dbget.close() : response.end
    end if

    '�̵��� �귣�� Ȯ��
    sqlstr = "select socname from db_user.dbo.tbl_user_c where userid='" & toMakerid & "'"
	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
    IF Not (rsget.EOF OR rsget.BOF) THEN
		toBrandName = rsget("socname")
    END IF
	rsget.Close
    
    if toBrandName="" or isNull(toBrandName) then
        response.write "<script type='text/javascript'>alert('�̵��� �귣��ID�� �������� �ʽ��ϴ�.[3]');</script>"
        dbget.close() : response.end
    end if

    vChangeContents = vChangeContents & "- �귣���̵� : from " & makerid & " to " & toMakerid & vbCrLf

    '### ���� �α� ����(item)
    vSCMChangeSQL = "INSERT INTO [db_log].[dbo].[tbl_scm_change_log] (userid, gubun, pk_idx, menupos, contents, refip)"
    vSCMChangeSQL = vSCMChangeSQL & "   select"
    vSCMChangeSQL = vSCMChangeSQL & "   '" & session("ssBctId") & "', 'item', i.itemid, '" & Request("menupos") & "'"
    vSCMChangeSQL = vSCMChangeSQL & "   , '"& vChangeContents &"'"
    vSCMChangeSQL = vSCMChangeSQL & "   , '" & Request.ServerVariables("REMOTE_ADDR") & "'"
    vSCMChangeSQL = vSCMChangeSQL & "   from db_item.dbo.tbl_item i with (nolock)"
    vSCMChangeSQL = vSCMChangeSQL & "   where i.makerid='"& makerid &"'"

    dbget.execute vSCMChangeSQL

    sqlstr = "update db_item.dbo.tbl_item" & vbcrlf
    sqlstr = sqlstr & " set makerid='"& toMakerid &"', " & vbcrlf
    sqlstr = sqlstr & " brandname='" & toBrandName & "', " & vbcrlf
    sqlstr = sqlstr & " lastupdate=GETDATE() " & vbcrlf
    sqlstr = sqlstr & " where makerid='"& makerid &"'"

    dbget.execute sqlstr, resultrows

    response.write "<script type='text/javascript'>"
    response.write "    alert('" & toMakerid & "(��)�� "& resultrows &"�� ó���Ǿ����ϴ�.');"
    response.write "    parent.location.reload();"
    response.write "</script>"
    dbget.close() : response.end

else
    response.write "<script type='text/javascript'>alert('���ٰ�ΰ� �߸� �Ǿ����ϴ�[900].');</script>"
    dbget.close() : response.end
end if
%>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->