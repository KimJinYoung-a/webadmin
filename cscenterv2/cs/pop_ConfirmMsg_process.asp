<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/cscenterv2/lib/incSessionAdminCS.asp" -->
<!-- #include virtual="/cscenterv2/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/cscenterv2/lib/function.asp"-->
<!-- #include virtual="/cscenterv2/lib/classes/cs/cs_aslistcls.asp"-->
<%
dim id, nextstate,reguserid
dim confirmregmsg, confirmfinishmsg, confirmfinishuserid, mode
id          = RequestCheckvar(request("id"),10)
mode        = RequestCheckvar(request("mode"),16)
nextstate   = RequestCheckvar(request("nextstate"),4)
reguserid   = session("ssBctID")
confirmregmsg       = html2db(request("confirmregmsg"))
confirmfinishmsg    = html2db(request("confirmfinishmsg"))
confirmfinishuserid = session("ssBctID")
if confirmregmsg <> "" then
	if checkNotValidHTML(confirmregmsg) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
	response.write "</script>"
	response.End
	end if
end If
if confirmfinishmsg <> "" then
	if checkNotValidHTML(confirmfinishmsg) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
	response.write "</script>"
	response.End
	end if
end if

dim ocsaslist
set ocsaslist = New CCSASList
ocsaslist.FRectCsAsID = id

if (id<>"") then
    ocsaslist.GetOneCSASMaster
end if


''확인요청정보 :
dim OCsConfirm
set OCsConfirm = new CCSASList
OCsConfirm.FRectCsAsID = id

if id<>"" then
    OCsConfirm.GetOneCsConfirmItem
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

if IsEditMode and (mode<>"reInput") then
    if mode="finish" then
        sqlStr = "update " & TABLE_CS_CONFIRM & ""                      & VbCrlf
        sqlStr = sqlStr + " set confirmfinishmsg='" & confirmfinishmsg & "'"    & VbCrlf
        sqlStr = sqlStr + " ,confirmfinishuserid='" & confirmfinishuserid & "'" & VbCrlf
        sqlStr = sqlStr + " ,confirmfinishdate=getdate()"                          & VbCrlf
        sqlStr = sqlStr + " where asid=" & id

        dbget.Execute sqlStr

        ''확인 완료시 -> 접수로 변경됨.
        sqlStr = "update " & TABLE_CSMASTER & ""                         & VbCrlf
        sqlStr = sqlStr + " set currstate='" & nextstate & "'"                  & VbCrlf
        sqlStr = sqlStr + " where id=" & id

        dbget.Execute sqlStr

        ''
        sqlStr = " update " & TABLE_CS_REFUND & "" + VbCrlf
        sqlStr = sqlStr + " set IBK_TIDX=NULL" + VbCrlf
        sqlStr = sqlStr + " where asid=" + CStr(id)

        dbget.Execute sqlStr

        response.write "<script>alert('확인 완료 되었습니다. 상태는 접수로 변경됩니다.');</script>"
        response.write "<script>opener.location.reload();</script>"
        response.write "<script>window.close();</script>"



    else

        sqlStr = "update " & TABLE_CS_CONFIRM & ""                      & VbCrlf
        sqlStr = sqlStr + " set confirmregmsg='" & confirmregmsg & "'"          & VbCrlf
        sqlStr = sqlStr + " ,confirmreguserid='" & reguserid & "'"              & VbCrlf
        sqlStr = sqlStr + " where asid=" & id

        dbget.Execute sqlStr


        response.write "<script>alert('수정 되었습니다.');</script>"
        response.write "<script>history.back();</script>"
    end if
else
    '''IBK 등록중인지 체크 : 실패내역은 SKIP
    Dim NotFinishedExists : NotFinishedExists= false
    sqlStr = "select count(*) as CNT from db_log.dbo.tbl_IBK_ERP_ICHE_DATA"                      & VbCrlf
    sqlStr = sqlStr + " where TEN_CSID=" & id
    sqlStr = sqlStr + " and IsNULL(PROC_YN,'Y')='Y'"                      & VbCrlf

    rsget_CS.CursorLocation = adUseClient
    rsget_CS.Open sqlStr, dbget_CS, adOpenForwardOnly
    if Not rsget_CS.Eof then
        NotFinishedExists = rsget_CS("CNT")>0
    end if
    rsget_CS.Close

    if (NotFinishedExists) then
        response.write "<script>alert('환불처리가 완료되지 않았거나, 이미 처리된 내역이 있습니다.!!');</script>"
        response.write "<script>history.back();</script>"
    end if


    if (mode="reInput") then
        sqlStr = "update " & TABLE_CS_CONFIRM & ""                     & VbCrlf
        sqlStr = sqlStr + " set confirmregmsg='" & confirmregmsg & "'"          & VbCrlf
        sqlStr = sqlStr + " ,confirmreguserid='" & reguserid & "'"              & VbCrlf
        sqlStr = sqlStr + " ,confirmfinishmsg=convert(varchar(2000),confirmfinishmsg)+char(13)+'====================='"    & VbCrlf
        sqlStr = sqlStr + " ,confirmfinishuserid=NULL" & VbCrlf
        sqlStr = sqlStr + " ,confirmfinishdate=NULL"                          & VbCrlf
        sqlStr = sqlStr + " where asid=" & id

        dbget.Execute sqlStr

    else
        sqlStr = "insert into " & TABLE_CS_CONFIRM & ""                     & VbCrlf
        sqlStr = sqlStr + " (asid, confirmregmsg, confirmreguserid)"                & VbCrlf
        sqlStr = sqlStr + " values(" & id                                           & VbCrlf
        sqlStr = sqlStr + " ,'" & confirmregmsg & "'"                               & VbCrlf
        sqlStr = sqlStr + " ,'" & reguserid & "'"                                   & VbCrlf
        sqlStr = sqlStr + ")"
    ''rw sqlStr
        dbget.Execute sqlStr
    end if

    sqlStr = "update " & TABLE_CSMASTER & ""                             & VbCrlf
    sqlStr = sqlStr + " set currstate='" & nextstate & "'"                      & VbCrlf
    sqlStr = sqlStr + " where id=" & id

    dbget.Execute sqlStr

    ''파일날짜 NULL
    sqlStr = " update " & TABLE_CS_REFUND & "" + VbCrlf
    sqlStr = sqlStr + " set upfiledate=NULL" + VbCrlf
    sqlStr = sqlStr + " where asid=" + CStr(id)

    dbget.Execute sqlStr

    response.write "<script>alert('접수되었습니다. 상태는 확인요청으로 변경됩니다.');</script>"
    response.write "<script>opener.location.reload();</script>"
    response.write "<script>window.close();</script>"
end if

set ocsaslist = Nothing
set OCsConfirm = Nothing
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->
