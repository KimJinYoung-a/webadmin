<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_aslistcls.asp" -->
<%
dim id, nextstate,reguserid
dim confirmregmsg, confirmfinishmsg, confirmfinishuserid, mode
dim sitegubun

id          = request("id")
mode        = request("mode")
nextstate   = request("nextstate")
reguserid   = session("ssBctID")
confirmregmsg       = html2db(request("confirmregmsg"))
confirmfinishmsg    = html2db(request("confirmfinishmsg"))
confirmfinishuserid = session("ssBctID")
sitegubun      		= RequestCheckVar(request("sitegubun"),32)


dim ocsaslist
set ocsaslist = New CCSASList
ocsaslist.FRectCsAsID = id

if (id<>"") then
	if (sitegubun = "academy") then
		ocsaslist.GetOneCSASMasterAcademy
	else
		'10x10
		ocsaslist.GetOneCSASMaster
	end if
end if


''Ȯ�ο�û���� :
dim OCsConfirm
set OCsConfirm = new CCSASList
OCsConfirm.FRectCsAsID = id

if id<>"" then
	if (sitegubun = "academy") then
		OCsConfirm.GetOneCsConfirmItemAcademy
	else
		'10x10
		OCsConfirm.GetOneCsConfirmItem
	end if
end if


if (ocsaslist.FResultCount<1) then
    response.write "<script>alert('��ȿ���� ���� �����Դϴ�.');</script>"
    response.write "<script>window.close();</script>"
    dbget.close()	:	response.End
end if

''Ȯ�ο�û������ ���� ���=> �������¸� ��� ����.
''Ȯ�ο�û������ �ִ� ���=> Ȯ�� ��û ���¿����� ����/�Ϸ� ����.
if (OCsConfirm.FResultCount<1) then
    if (ocsaslist.FOneItem.FCurrstate>="B006") then
        response.write "<script>alert('�Ϸ� ���� ���� ������ ��� �����մϴ�.');</script>"
        response.write "<script>history.back();</script>"
        dbget.close()	:	response.End
    end if
else
    if (ocsaslist.FOneItem.FCurrstate>="B006") then
        response.write "<script>alert('�Ϸ� ���� ���� ������ ����/�Ϸ� �����մϴ�.');</script>"
        response.write "<script>history.back();</script>"
        dbget.close()	:	response.End
    end if
end if


dim IsEditMode
IsEditMode = (OCsConfirm.FResultCount>0)



dim sqlStr
dim TBL_AS_CONFIRM, TBL_AS_LIST, TBL_REFUND_INFO

if (sitegubun = "academy") then
	TBL_AS_CONFIRM 	= "[db_academy].[dbo].tbl_academy_as_confirm"
	TBL_AS_LIST		= "[db_academy].[dbo].tbl_academy_as_list"
	TBL_REFUND_INFO = "[db_academy].[dbo].tbl_academy_as_refund_info"
else
	'10x10
	TBL_AS_CONFIRM 	= "[db_cs].[dbo].tbl_new_as_confirm"
	TBL_AS_LIST		= "[db_cs].[dbo].tbl_new_as_list"
	TBL_REFUND_INFO = "[db_cs].[dbo].tbl_as_refund_info"
end if

function ExecuteQuery(sitegubun, strsql)

	if (sitegubun = "academy") then
		dbACADEMYget.Execute strsql
	else
		'10x10
		dbget.Execute strsql
	end if

end function

if IsEditMode and (mode<>"reInput") then
    if mode="finish" then
        sqlStr = "update " + TBL_AS_CONFIRM + ""                      & VbCrlf
        sqlStr = sqlStr + " set confirmfinishmsg='" & confirmfinishmsg & "'"    & VbCrlf
        sqlStr = sqlStr + " ,confirmfinishuserid='" & confirmfinishuserid & "'" & VbCrlf
        sqlStr = sqlStr + " ,confirmfinishdate=getdate()"                          & VbCrlf
        sqlStr = sqlStr + " where asid=" & id

        Call ExecuteQuery(sitegubun, sqlStr)
        'dbget.Execute sqlStr

        ''Ȯ�� �Ϸ�� -> ������ �����.
        sqlStr = "update " + TBL_AS_LIST + ""                         & VbCrlf
        sqlStr = sqlStr + " set currstate='" & nextstate & "'"                  & VbCrlf
        sqlStr = sqlStr + " where id=" & id

        Call ExecuteQuery(sitegubun, sqlStr)
        'dbget.Execute sqlStr

        ''
        sqlStr = " update " + TBL_REFUND_INFO + "" + VbCrlf
        sqlStr = sqlStr + " set IBK_TIDX=NULL" + VbCrlf
        sqlStr = sqlStr + " where asid=" + CStr(id)

        Call ExecuteQuery(sitegubun, sqlStr)
        'dbget.Execute sqlStr

        response.write "<script>alert('Ȯ�� �Ϸ� �Ǿ����ϴ�. ���´� ������ ����˴ϴ�.');</script>"
        response.write "<script>opener.location.reload();</script>"
        response.write "<script>window.close();</script>"



    else

        sqlStr = "update " + TBL_AS_CONFIRM + ""                      & VbCrlf
        sqlStr = sqlStr + " set confirmregmsg='" & confirmregmsg & "'"          & VbCrlf
        sqlStr = sqlStr + " ,confirmreguserid='" & reguserid & "'"              & VbCrlf
        sqlStr = sqlStr + " where asid=" & id

        Call ExecuteQuery(sitegubun, sqlStr)
        'dbget.Execute sqlStr


        response.write "<script>alert('���� �Ǿ����ϴ�.');</script>"
        response.write "<script>history.back();</script>"
    end if
else
    '''IBK ��������� üũ : ���г����� SKIP
    Dim NotFinishedExists : NotFinishedExists= false
    sqlStr = "select count(*) as CNT from db_log.dbo.tbl_IBK_ERP_ICHE_DATA"                      & VbCrlf
    sqlStr = sqlStr + " where TEN_CSID=" & id
    sqlStr = sqlStr + " and IsNULL(PROC_YN,'Y')='Y'"                      & VbCrlf

    rsget.CursorLocation = adUseClient
    rsget.Open sqlStr, dbget, adOpenForwardOnly
    if Not rsget.Eof then
        NotFinishedExists = rsget("CNT")>0
    end if
    rsget.Close

    if (NotFinishedExists) then
        response.write "<script>alert('ȯ��ó���� �Ϸ���� �ʾҰų�, �̹� ó���� ������ �ֽ��ϴ�.!!');</script>"
        response.write "<script>history.back();</script>"
    end if


    if (mode="reInput") then
        sqlStr = "update " + TBL_AS_CONFIRM + ""                     & VbCrlf
        sqlStr = sqlStr + " set confirmregmsg='" & confirmregmsg & "'"          & VbCrlf
        sqlStr = sqlStr + " ,confirmreguserid='" & reguserid & "'"              & VbCrlf
        sqlStr = sqlStr + " ,confirmfinishmsg=convert(varchar(2000),confirmfinishmsg)+char(13)+'====================='"    & VbCrlf
        sqlStr = sqlStr + " ,confirmfinishuserid=NULL" & VbCrlf
        sqlStr = sqlStr + " ,confirmfinishdate=NULL"                          & VbCrlf
        sqlStr = sqlStr + " where asid=" & id

        Call ExecuteQuery(sitegubun, sqlStr)
        'dbget.Execute sqlStr

    else
        sqlStr = "insert into " + TBL_AS_CONFIRM + ""                     & VbCrlf
        sqlStr = sqlStr + " (asid, confirmregmsg, confirmreguserid)"                & VbCrlf
        sqlStr = sqlStr + " values(" & id                                           & VbCrlf
        sqlStr = sqlStr + " ,'" & confirmregmsg & "'"                               & VbCrlf
        sqlStr = sqlStr + " ,'" & reguserid & "'"                                   & VbCrlf
        sqlStr = sqlStr + ")"

        Call ExecuteQuery(sitegubun, sqlStr)
        'dbget.Execute sqlStr
    end if

    sqlStr = "update " + TBL_AS_LIST + ""                             & VbCrlf
    sqlStr = sqlStr + " set currstate='" & nextstate & "'"                      & VbCrlf
    sqlStr = sqlStr + " where id=" & id

    Call ExecuteQuery(sitegubun, sqlStr)
    'dbget.Execute sqlStr

    ''���ϳ�¥ NULL
    sqlStr = " update " + TBL_REFUND_INFO + "" + VbCrlf
    sqlStr = sqlStr + " set upfiledate=NULL" + VbCrlf
    sqlStr = sqlStr + " where asid=" + CStr(id)

    Call ExecuteQuery(sitegubun, sqlStr)
    'dbget.Execute sqlStr

    response.write "<script>alert('�����Ǿ����ϴ�. ���´� Ȯ�ο�û���� ����˴ϴ�.');</script>"
    response.write "<script>opener.location.reload();</script>"
    response.write "<script>window.close();</script>"
end if

set ocsaslist = Nothing
set OCsConfirm = Nothing
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
