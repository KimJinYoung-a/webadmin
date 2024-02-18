<%@ language=vbscript %>
<% option explicit %>
<%
Response.Expires = 0   
Response.AddHeader "Pragma","no-cache"   
Response.AddHeader "Cache-Control","no-cache,must-revalidate"   

'###############################################
' PageName : culturestation_process.asp
' Discription : 컬쳐스테이션 컨텐츠 등록 프로세스
' History : 2019.02.20 정태훈
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->

<%

dim eCode, eMode, strSql, eName, subcopyK, evt_kind, evt_comment
dim m_main_content, m_cmt_desc, evt_mainimg, evt_type, themecolor
dim refer
refer = request.ServerVariables("HTTP_REFERER")
eCode = requestCheckVar(Request.Form("evt_code"),10)
eMode = requestCheckVar(Request.Form("imod"),2)
eName = requestCheckVar(Request.Form("sEN"),128)
evt_type = requestCheckVar(Request.Form("evt_type"),10)
evt_kind = requestCheckVar(Request.Form("evt_kind"),10)
subcopyK = requestCheckVar(Request.Form("subcopyK"),128)
evt_mainimg = requestCheckVar(Request.Form("evt_mainimg"),128)
m_main_content	= html2db(replace(request("maincontent"),"'",""""))
m_cmt_desc		= html2db(replace(request("maincontent2"),"'",""""))
themecolor = requestCheckVar(Request.Form("DFcolorCD"),2)
evt_comment = requestCheckVar(Request.Form("evt_comment"),64)

if eCode="" then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 데이터 입니다. 다시 시도해 주세요.');history.back();"
	response.write "</script>"
	response.End
end if

if m_cmt_desc <> "" then
	if checkNotValidHTML(m_cmt_desc) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
	response.write "</script>"
	response.End
	end if
end If

if eName <> "" then
	if checkNotValidHTML(eName) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
	response.write "</script>"
	response.End
	end if
end If

if subcopyK <> "" then
	if checkNotValidHTML(subcopyK) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
	response.write "</script>"
	response.End
	end if
end If

select case eMode
case "CU"
	dbget.beginTrans

        '===========================================================
        '--2.disply 수정
        strSql = "UPDATE [db_event].[dbo].[tbl_event_display]" & vbCrlf
        strSql = strSql + " SET evt_mainimg='" & evt_mainimg & "'" & vbCrlf
        strSql = strSql + ", evt_html='" & html2db(m_main_content) & "'" & vbCrlf
        strSql = strSql + ", evt_html_mo='" & html2db(m_cmt_desc) & "'" & vbCrlf
		strSql = strSql + ", eventtype_pc='" & evt_type & "'" & vbCrlf
		strSql = strSql + ", eventtype_mo='" & evt_kind & "'" & vbCrlf
		strSql = strSql + ", themecolor='" & themecolor & "'" & vbCrlf
		strSql = strSql + ", evt_comment='" & evt_comment & "'" & vbCrlf
        strSql = strSql + ", evt_template_mo='12' ,evt_template='12'" & vbCrlf
        strSql = strSql + " where evt_code=" & eCode
        dbget.execute strSql

        if Err.Number <> 0 then
            dbget.RollBackTrans 
            Call sbAlertMsg ("데이터 처리에 문제가 발생하였습니다.[1]", "back", "")
            response.End
        end if
    '===============================================================
	dbget.CommitTrans
    
	response.write "<script type='text/javascript'>"
	response.write "    window.document.domain = ""10x10.co.kr"";"
	response.write "	opener.document.location.replace('/admin/eventmanage/event/v5/event_register.asp?eC=" + Cstr(eCode) + "&togglediv=1&viewset='+opener.document.frmEvt.viewset.value);"
    'response.write "    location.replace('" + refer + "');"
	response.write "    self.close();"
	response.write "</script>"
	dbget.close()	:	response.End
case else
	Call sbAlertMsg ("데이터 처리에 문제가 발생하였습니다.", "back", "")
end select
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->