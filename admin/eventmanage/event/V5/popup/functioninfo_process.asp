<%@ language=vbscript %>
<% option explicit %>
<%
Response.Expires = 0   
Response.AddHeader "Pragma","no-cache"   
Response.AddHeader "Cache-Control","no-cache,must-revalidate"   

'###############################################
' PageName : functioninfo_process.asp
' Discription : I형(통합형) 이벤트 기능정보 등록 프로세스
' History : 2019.02.15 정태훈
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->

<%
dim eCode, strSql
dim eComment, eBbs, eItemps, eisblogurl
dim comm_isusing, comm_text, freebie_img, comm_start, comm_end
dim eval_isusing, eval_text, eval_freebie_img, eval_start, eval_end
dim board_isusing, board_text, board_freebie_img, board_start, board_end
dim refer, ePdate

refer = request.ServerVariables("HTTP_REFERER")
eCode = requestCheckVar(Request.Form("evt_code"),10)

eComment = requestCheckVar(Request.Form("chComm"),1)
eBbs = requestCheckVar(Request.Form("chBbs"),1)
eItemps	= requestCheckVar(Request.Form("chItemps"),1)
eisblogurl	= requestCheckVar(Request.Form("isblogurl"),1)

comm_isusing = requestCheckVar(Request.Form("comm_isusing"),1)
comm_text = Request.Form("comm_text")
freebie_img = Request.Form("freebie_img")
comm_start = requestCheckVar(Request.Form("comm_start"),10)
comm_end = requestCheckVar(Request.Form("comm_end"),10)

eval_isusing = requestCheckVar(Request.Form("eval_isusing"),1)
eval_text = Request.Form("eval_text")
eval_freebie_img = Request.Form("eval_freebie_img")
eval_start = requestCheckVar(Request.Form("eval_start"),10)
eval_end = requestCheckVar(Request.Form("eval_end"),10)

board_isusing = requestCheckVar(Request.Form("board_isusing"),1)
board_text = Request.Form("board_text")
board_freebie_img = Request.Form("board_freebie_img")
board_start = requestCheckVar(Request.Form("board_start"),10)
board_end = requestCheckVar(Request.Form("board_end"),10)

if comm_text <> "" then
	if checkNotValidHTML(comm_text) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
	response.write "</script>"
	response.End
	end if
end if

if freebie_img <> "" then
	if checkNotValidHTML(freebie_img) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
	response.write "</script>"
	response.End
	end if
end if

if eval_text <> "" then
	if checkNotValidHTML(eval_text) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
	response.write "</script>"
	response.End
	end if
end if

if eval_freebie_img <> "" then
	if checkNotValidHTML(eval_freebie_img) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
	response.write "</script>"
	response.End
	end if
end if

if board_text <> "" then
	if checkNotValidHTML(board_text) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
	response.write "</script>"
	response.End
	end if
end if

if board_freebie_img <> "" then
	if checkNotValidHTML(board_freebie_img) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
	response.write "</script>"
	response.End
	end if
end if

if eComment ="" then eComment = 0
if eBbs ="" then eBbs = 0
if eItemps ="" then eItemps = 0
if eisblogurl ="" then eisblogurl = 0
'당첨일 삭제 여부 체크
if eComment=0 and eBbs=0 and eItemps=0 then ePdate=""
'--------------------------------------------------------
' 데이터 처리
' I : 이벤트 개요등록, U: 개요수정, disply등록/수정
'--------------------------------------------------------
	
	'트랜잭션 (1.master수정/2.disply수정/3.MDTheme수정)
	dbget.beginTrans

        if eComment=0 and eBbs=0 and eItemps=0 then
            '--1.master 수정
            strSql = "UPDATE [db_event].[dbo].[tbl_event]" & vbCrlf
            strSql = strSql + " SET evt_prizedate='" & ePdate & "'" & vbCrlf
            strSql = strSql + " WHERE evt_code=" & eCode
            dbget.execute strSql
        end if

        '--2.disply 수정
        strSql = "UPDATE [db_event].[dbo].[tbl_event_display]" & vbCrlf
        strSql = strSql + " SET iscomment=" & eComment & vbCrlf
        strSql = strSql + ", isbbs=" & eBbs & vbCrlf
        strSql = strSql + ", isitemps=" & eItemps & vbCrlf
        strSql = strSql + ", isGetBlogURL=" & eisblogurl & vbCrlf
        strSql = strSql + " where evt_code=" & eCode
        dbget.execute strSql

        if Err.Number <> 0 then
            dbget.RollBackTrans 
            Call sbAlertMsg ("데이터 처리에 문제가 발생하였습니다.[1]", "back", "")
            response.End 
        end if

        '--3.MDTheme 수정
        strSql = "UPDATE [db_event].[dbo].[tbl_event_md_theme]" & vbCrlf
        strSql = strSql + " SET gift_isusing=0" & vbCrlf
        If comm_isusing <> "" Then
        strSql = strSql & " ,comm_isusing='" & comm_isusing & "'"
        strSql = strSql & " ,comm_text='" & html2db(comm_text) & "'"
        If freebie_img <> "" Then
        strSql = strSql & " ,freebie_img='" & freebie_img & "'"
        End If
        strSql = strSql & " ,comm_start='" & comm_start & "'"
        strSql = strSql & " ,comm_end='" & comm_end & "'"
        End If
        If eval_isusing <> "" Then
        strSql = strSql & " ,eval_isusing='" & eval_isusing & "'"
        strSql = strSql & " ,eval_text='" & html2db(eval_text) & "'"
        If eval_freebie_img <> "" Then
        strSql = strSql & " ,eval_freebie_img='" & eval_freebie_img & "'"
        End If
        strSql = strSql & " ,eval_start='" & eval_start & "'"
        strSql = strSql & " ,eval_end='" & eval_end & "'"
        End If
        If board_isusing <> "" Then
        strSql = strSql & " ,board_isusing='" & board_isusing & "'"
        strSql = strSql & " ,board_text='" & html2db(board_text) & "'"
        If board_freebie_img <> "" Then
        strSql = strSql & " ,board_freebie_img='" & board_freebie_img & "'"
        End If
        strSql = strSql & " ,board_start='" & board_start & "'"
        strSql = strSql & " ,board_end='" & board_end & "'"
        End If
        strSql = strSql + " where evt_code=" & eCode
        dbget.execute strSql

        if Err.Number <> 0 then
            dbget.RollBackTrans 
            Call sbAlertMsg ("데이터 처리에 문제가 발생하였습니다.[2]", "back", "")
            response.End 
        end if
    '===========================================================
	dbget.CommitTrans
    
	response.write "<script type='text/javascript'>"
    response.write "    window.document.domain = ""10x10.co.kr"";"
	response.write "	opener.document.location.replace('/admin/eventmanage/event/v5/event_register.asp?eC=" + Cstr(eCode) + "&togglediv=4&viewset='+opener.document.frmEvt.viewset.value);"
    'response.write "    location.replace('" + refer + "');"
    response.write "    self.close();"
	response.write "</script>"
	dbget.close()	:	response.End
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->