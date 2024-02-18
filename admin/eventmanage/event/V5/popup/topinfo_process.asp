<%@ language=vbscript %>
<% option explicit %>
<%
Response.Expires = 0   
Response.AddHeader "Pragma","no-cache"   
Response.AddHeader "Cache-Control","no-cache,must-revalidate"   

'###############################################
' PageName : topinfo_process.asp
' Discription : I형(통합형) 이벤트 탑 배너, 정보 등록 프로세스
' History : 2019.01.29 정태훈
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->

<%

dim eCode, eMode, strSql
dim evt_template_mo, evt_template, title_mo, title_pc
dim subcopyK, subsEN, chkWide, mdbntype, mdbntypemo
dim themecolor, textbgcolor, themecolormo, textbgcolormo
dim evt_html, evt_mainimg, evt_html_mo, evt_mainimg_mo
dim GroupItemType, contentsAlign, eSlideYN_W, eSlideYN_M
dim refer, eventtype_pc, eventtype_mo, chkFull
dim blnexec, blnexec_mo, eexecfile, eexecfile_mo, copyhide

refer = request.ServerVariables("HTTP_REFERER")
eCode = requestCheckVar(Request.Form("evt_code"),10)
eMode = requestCheckVar(Request.Form("imod"),2)
evt_template_mo = requestCheckVar(Request.Form("evt_template_mo"),2)
evt_template = requestCheckVar(Request.Form("evt_template"),2)
title_mo = requestCheckVar(Request.Form("title_mo"),120)
title_pc = requestCheckVar(Request.Form("title_pc"),120)
subsEN = requestCheckVar(Request.Form("subsEN"),120)
subcopyK = requestCheckVar(Request.Form("subcopyK"),120)
chkWide = requestCheckVar(Request.Form("chkWide"),10)
chkFull = requestCheckVar(Request.Form("chkFull"),10)
mdbntype = requestCheckVar(Request.Form("mdbntype"),1)
mdbntypemo = requestCheckVar(Request.Form("mdbntypemo"),1)
themecolor  		= requestCheckVar(Request.Form("DFcolorCD"),3)
textbgcolor  	= requestCheckVar(Request.Form("DFcolorCD2"),3)
themecolormo  	= requestCheckVar(Request.Form("DFcolorCDMo"),3)
textbgcolormo  	= requestCheckVar(Request.Form("DFcolorCDMo2"),3)
GroupItemType  	= requestCheckVar(Request.Form("GroupItemType"),1)
contentsAlign  	= requestCheckVar(Request.Form("contentsAlign"),1)
eSlideYN_W	= requestCheckVar(Request.Form("slide_w_flag"),1)	'슬라이드 사용/pc
eSlideYN_M	= requestCheckVar(Request.Form("slide_m_flag"),1)	'슬라이드 사용/mo
copyhide	= requestCheckVar(Request.Form("copyhide"),1)	'모바일 카피 / 서브카피 숨김 유무

evt_html = html2db(Request.Form("tHtml"))		'화면설정html 코드
evt_mainimg = Request.Form("main")
evt_html_mo = html2db(Request.Form("tHtml_mo"))
evt_mainimg_mo = Request.Form("main_mo")

eventtype_pc = requestCheckVar(Request.Form("pc_evttype"),3)
eventtype_mo = requestCheckVar(Request.Form("mo_evttype"),3)

blnexec     = requestCheckVar(Request.Form("rdoEF"),1)
blnexec_mo  = requestCheckVar(Request.Form("rdoEF_mo"),1)

'if evt_template_mo="11" then eventtype_mo=""
'if evt_template="10" then eventtype_pc=""

'response.write chkFull & "<br>"
'response.write chkWide & "<br>"
'response.end
IF blnexec = "" THEN blnexec = 0    
IF blnexec_mo = "" THEN blnexec_mo = 0
IF chkFull = ""	THEN chkFull = 1
IF chkWide = ""	THEN chkWide = 0

if contentsAlign="" then contentsAlign=1
if contentsAlign=1 then chkFull=1
if contentsAlign=2 then chkWide=1
if eventtype_pc="" then
	if contentsAlign="2" then
		eventtype_pc="50"
	else
		eventtype_pc="20"
	end if
end if

if blnexec = "1" then
	eexecfile   =  requestCheckVar(Request.Form("sEFP"),128)
else
	eexecfile = ""  
end if
if blnexec_mo = "1" then
	eexecfile_mo=  requestCheckVar(Request.Form("sEFP_mo"),128)
else
	eexecfile_mo = ""
end If

if title_mo <> "" then
	if checkNotValidHTML(title_mo) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
	response.write "</script>"
	response.End
	end if
end If

if title_pc <> "" then
	if checkNotValidHTML(title_pc) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
	response.write "</script>"
	response.End
	end if
end If

if subsEN <> "" then
	if checkNotValidHTML(subsEN) then
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

if evt_mainimg <> "" then
	if checkNotValidHTML(evt_mainimg) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
	response.write "</script>"
	response.End
	end if
end If

if evt_mainimg_mo <> "" then
	if checkNotValidHTML(evt_mainimg_mo) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
	response.write "</script>"
	response.End
	end if
end If

if eCode="" then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 데이터 입니다. 다시 시도해 주세요.');history.back();"
	response.write "</script>"
	response.End
end if

select case eMode
case "TU"
	dbget.beginTrans

		'--1.master 수정
		strSql = "UPDATE [db_event].[dbo].[tbl_event]" & vbCrlf
        strSql = strSql + " SET evt_subcopyK='" & subcopyK & "'" & vbCrlf
        strSql = strSql + ", evt_subname='" & subsEN & "'" & vbCrlf
        strSql = strSql + " WHERE evt_code=" & eCode
		dbget.execute strSql

        if Err.Number <> 0 then
            dbget.RollBackTrans 
            Call sbAlertMsg ("데이터 처리에 문제가 발생하였습니다.[1]", "back", "")
            response.End 
        end if

        '===========================================================
        '--2.disply 수정
        strSql = "UPDATE [db_event].[dbo].[tbl_event_display]" & vbCrlf
        strSql = strSql + " SET evt_fullyn=" & chkFull & "" & vbCrlf
		strSql = strSql + ", evt_template_mo='" & evt_template_mo & "'" & vbCrlf
        strSql = strSql + ", evt_template='" & evt_template & "'" & vbCrlf
        strSql = strSql + ", evt_wideyn=" & chkWide & "" & vbCrlf
        strSql = strSql + ", mdbntype='" & mdbntype & "'" & vbCrlf
        strSql = strSql + ", mdbntypemo='" & mdbntypemo & "'" & vbCrlf
        strSql = strSql + ", themecolor='" & themecolor & "'" & vbCrlf
        strSql = strSql + ", textbgcolor='" & textbgcolor & "'" & vbCrlf
        strSql = strSql + ", themecolormo='" & themecolormo & "'" & vbCrlf
        strSql = strSql + ", textbgcolormo='" & textbgcolormo & "'" & vbCrlf
        strSql = strSql + ", evt_html='" & evt_html & "'" & vbCrlf
        strSql = strSql + ", evt_html_mo='" & evt_html_mo & "'" & vbCrlf
        strSql = strSql + ", evt_mainimg='" & evt_mainimg & "'" & vbCrlf
        strSql = strSql + ", evt_mainimg_mo='" & evt_mainimg_mo & "'" & vbCrlf
		strSql = strSql + ", evt_slide_w_flag='" & eSlideYN_W & "'" & vbCrlf
		strSql = strSql + ", evt_slide_m_flag='" & eSlideYN_M & "'" & vbCrlf
		strSql = strSql + ", eventtype_pc='" & eventtype_pc & "'" & vbCrlf
		strSql = strSql + ", eventtype_mo='" & eventtype_mo & "'" & vbCrlf
		strSql = strSql + ", evt_isExec=" & blnexec & vbCrlf
		strSql = strSql + ", evt_execFile='" & eexecfile & "'" & vbCrlf
		strSql = strSql + ", evt_isExec_mo=" & blnexec_mo & vbCrlf
		strSql = strSql + ", evt_execFile_mo='" & eexecfile_mo & "'" & vbCrlf
		strSql = strSql + ", videoType='" & copyhide & "'" & vbCrlf
        strSql = strSql + " where evt_code=" & eCode
        dbget.execute strSql

        if Err.Number <> 0 then
            dbget.RollBackTrans 
            Call sbAlertMsg ("데이터 처리에 문제가 발생하였습니다.[2]", "back", "")
            response.End 
        end if

        '--3.theme 수정
        strSql = "UPDATE [db_event].[dbo].[tbl_event_md_theme]" & vbCrlf
        strSql = strSql + " SET title_mo='" & title_mo & "'" & vbCrlf
        strSql = strSql + " , title_pc='" & title_pc & "'" & vbCrlf
		strSql = strSql + " , GroupItemType='" & GroupItemType & "'" & vbCrlf
		strSql = strSql + " , contentsAlign='" & contentsAlign & "'" & vbCrlf
        strSql = strSql + " where evt_code=" & eCode
        dbget.execute strSql

        if Err.Number <> 0 then
            dbget.RollBackTrans 
            Call sbAlertMsg ("데이터 처리에 문제가 발생하였습니다.[3]", "back", "")
            response.End 
        end if
    '===========================================================
	dbget.CommitTrans
    
	response.write "<script type='text/javascript'>"
	response.write "    window.document.domain = ""10x10.co.kr"";"
	response.write "	opener.document.location.replace('/admin/eventmanage/event/v5/event_register.asp?eC=" + Cstr(eCode) + "&togglediv=2&viewset='+opener.document.frmEvt.viewset.value);"
    'response.write "    location.replace('" + refer + "');"
    response.write "    self.close();"
	response.write "</script>"
	dbget.close()	:	response.End
case else
	Call sbAlertMsg ("데이터 처리에 문제가 발생하였습니다.", "back", "")
end select
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->