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


''확인요청정보 :
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
    response.write "<script>alert('유효하지 않은 내역입니다.');</script>"
    response.write "<script>window.close();</script>"
    dbget.close()	:	response.End
end if

''확인요청내역이 없는 경우=> 접수상태만 등록 가능.
''확인요청내역이 있는 경우=> 확인 요청 상태에서만 수정/완료 가능.
if (OCsConfirm.FResultCount<1) then
    if (ocsaslist.FOneItem.FCurrstate>="B006") then
        response.write "<script>alert('완료 이전 상태 에서만 등록 가능합니다.');</script>"
        response.write "<script>history.back();</script>"
        dbget.close()	:	response.End
    end if
else
    if (ocsaslist.FOneItem.FCurrstate>="B006") then
        response.write "<script>alert('완료 이전 상태 에서만 수정/완료 가능합니다.');</script>"
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

        ''확인 완료시 -> 접수로 변경됨.
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

        response.write "<script>alert('확인 완료 되었습니다. 상태는 접수로 변경됩니다.');</script>"
        response.write "<script>opener.location.reload();</script>"
        response.write "<script>window.close();</script>"



    else

        sqlStr = "update " + TBL_AS_CONFIRM + ""                      & VbCrlf
        sqlStr = sqlStr + " set confirmregmsg='" & confirmregmsg & "'"          & VbCrlf
        sqlStr = sqlStr + " ,confirmreguserid='" & reguserid & "'"              & VbCrlf
        sqlStr = sqlStr + " where asid=" & id

        Call ExecuteQuery(sitegubun, sqlStr)
        'dbget.Execute sqlStr


        response.write "<script>alert('수정 되었습니다.');</script>"
        response.write "<script>history.back();</script>"
    end if
else
    '''IBK 등록중인지 체크 : 실패내역은 SKIP
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
        response.write "<script>alert('환불처리가 완료되지 않았거나, 이미 처리된 내역이 있습니다.!!');</script>"
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

    ''파일날짜 NULL
    sqlStr = " update " + TBL_REFUND_INFO + "" + VbCrlf
    sqlStr = sqlStr + " set upfiledate=NULL" + VbCrlf
    sqlStr = sqlStr + " where asid=" + CStr(id)

    Call ExecuteQuery(sitegubun, sqlStr)
    'dbget.Execute sqlStr

    response.write "<script>alert('접수되었습니다. 상태는 확인요청으로 변경됩니다.');</script>"
    response.write "<script>opener.location.reload();</script>"
    response.write "<script>window.close();</script>"
end if

set ocsaslist = Nothing
set OCsConfirm = Nothing
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
